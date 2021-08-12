const ui = DocumentApp.getUi();
const body = DocumentApp.getActiveDocument().getBody();
let formUrl = ''; // set on document open, updateForm

function onOpen(e) {
  ui.createMenu('Form builder')
    .addSubMenu(
      ui.createMenu('Insert new question')
        // https://developers.google.com/apps-script/reference/forms/item-type
        .addItem('Short answer', 'insertShortAnswerQuestion')
        .addItem('Long answer (paragraph)', 'insertLongAnswerQuestion')
        .addItem('Dropdown list (single-select)', 'insertDropdownQuestion')
        .addItem('Multiple choice (single-select)', 'insertMultipleChoiceQuestion')
        .addItem('Checkbox (multi-select)', 'insertCheckboxQuestion')
        .addItem('Linear scale', 'insertLinearScaleQuestion')
        )
    .addSeparator()
    .addItem('Update form', 'updateForm')
    .addToUi();

    formUrl = setFormUrl();
}

function setFormUrl() {
  // https://developers.google.com/apps-script/reference/slides/table?hl=en
  const formUrlTable = body.getTables()[0];
  if (!formUrlTable || !formUrlTable.getNumRows()) {
    return [];
  }
  return formUrlTable.getCell(0, 1).getText()
}

// Question generation

function insertQuestion(type) {
  var rowsData = [['Question', ''], ['Type', type]];
  if (type === 'Linear scale') {
    rowsData.push(['Lower bound', '1: Least label'])
    rowsData.push(['Upper bound', '5: Most label'])
  } else if (!['Short answer', 'Long answer'].includes(type)) {
    rowsData.push(['Options', ''])
  }
  rowsData.push(['Required?', ''])
  const table = body.appendTable(rowsData);

  let style = {};
  style[DocumentApp.Attribute.BOLD] = true;
  for (let i = 0; i < table.getNumRows(); i++) {
    table.getCell(i, 0).setAttributes(style).setWidth(80)
  }
}
function insertShortAnswerQuestion() {
  insertQuestion('Short answer')
}
function insertLongAnswerQuestion() {
  insertQuestion('Long answer')
}
function insertDropdownQuestion() {
  insertQuestion('Dropdown list')
}
function insertMultipleChoiceQuestion() {
  insertQuestion('Multiple choice')
}
function insertCheckboxQuestion() {
  insertQuestion('Checkbox')
}
function insertLinearScaleQuestion() {
  insertQuestion('Linear scale')
}

// Update the form

function updateForm() {
  if (!formUrl) {
    formUrl = setFormUrl();
  }
  const form = FormApp.openByUrl(formUrl);

  createSection(form);

  const questionTables = body.getTables().slice(1);

  questionTables.forEach((table) => {
    processQuestion(form, table);
  });
}

function processQuestion(form, table) {
  const question = table.getCell(0, 1).getText()
  const type = table.getCell(1, 1).getText()
  const isRequired = ['y', 'yes'].includes(table.getCell(table.getNumRows() - 1, 1).getText().toLowerCase())

  if (!question && !type) {
    return;
  }

  let item = null;

  if (type === 'Short answer') {
    item = form.addTextItem();
    item.setTitle(question);
  } else if (type === 'Long answer') {
    item = form.addParagraphTextItem();
    item.setTitle(question);
  } else if (type === 'Dropdown list') {
    item = form.addListItem();
    addQuestionWithOptions({ item, table, allowCustomOther: false })
  } else if (type === 'Multiple choice') {
    item = form.addMultipleChoiceItem();
    addQuestionWithOptions({ item, table, allowCustomOther: true })
  } else if (type === 'Checkbox') {
    item = form.addCheckboxItem();
    addQuestionWithOptions({ item, table, allowCustomOther: true })
  } else if (type === 'Linear scale') {
    item = form.addScaleItem();
    const lower = table.getCell(2, 1).getText().split(': ')
    const upper = table.getCell(3, 1).getText().split(': ')
    item.setTitle(question)
      .setBounds(lower[0], upper[0])
      .setLabels(lower[1], upper[1])
  }

  item.setRequired(isRequired)
}

function addQuestionWithOptions({ item, table, allowCustomOther }) {
  const question = table.getCell(0, 1).getText()
  const options = table.getCell(2, 1).getText()

  let showOther = allowCustomOther && options.toLowerCase().includes('other');
  let choices = options.trim().split('\n').reduce((result, option) => {
    if (
      // If "other" is a special option for custom answer, skip adding "other" as an set option to choose
      (option.toLowerCase() === 'other' && !allowCustomOther) ||
      (option.toLowerCase() !== 'other')) {
      result.push(item.createChoice(option));
    }
    return result;
  }, []);

  item.setTitle(question).setChoices(choices)
  if (allowCustomOther) {
    item.showOtherOption(showOther);
  }
}

function createSection(form) {
  const section = form.addPageBreakItem().setTitle('Page Two');
}
