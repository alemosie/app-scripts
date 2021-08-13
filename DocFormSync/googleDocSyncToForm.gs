const ui = DocumentApp.getUi();
const body = DocumentApp.getActiveDocument().getBody();
let formUrl = ''; // set on document open, updateForm

function onOpen(e) {
  ui.createMenu('✨ Form builder')
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
  const questionTables = body.getTables().slice(1);

  if (!formUrl) {
    formUrl = setFormUrl();
  }
  const form = FormApp.openByUrl(formUrl);
  const existingFormQuestions = form.getItems()

  if (existingFormQuestions.length > 0) {
    for (let i = 0; i < existingFormQuestions.length; i++) {
      const formItem = existingFormQuestions[i];
      const tableQuestion = questionTables[i];
      updateExistingQuestion(formItem, tableQuestion);
    }
  } else {
    questionTables.forEach((table) => {
      addNewQuestion(form, table);
    });
  }
}



// Helpers

function parseTableQuestion(table) {
  let questionData = {};
  const numRows = table.getNumRows();
  for (let i = 0; i < table.getNumRows(); i++) {
    const key = table.getCell(i, 0).getText().replace('?', '').toLowerCase();
    const value = table.getCell(i, 1).getText();

    if (key !== 'lower bound' && key !== 'upper bound') {
      questionData[key] = value;
    } else {
      // Handle lower and upper bound data
      const boundType = key.split(' ')[0]; // 'lower' or 'upper'
      const values = value.split(': ');
      questionData[boundType] = {
        bound: values[0],
        label: values[1]
      }
    }
  }
  questionData.required = ['y', 'yes'].includes(questionData.required.toLowerCase());
  return questionData;
}

function getAllowCustomOther(type) {
  return ['Multiple choice', 'Checkbox'].includes(type);
}

function generateFormQuestionOptions({ item, options, allowCustomOther }) {
  return options.trim().split('\n').reduce((result, option) => {
    if (
      // If "other" is a special option for custom answer, skip adding "other" as an set option to choose
      (option.toLowerCase() === 'other' && !allowCustomOther) ||
      (option.toLowerCase() !== 'other')) {
      result.push(item.createChoice(option));
    }
    return result;
  }, []);
}




// Edit question

function updateExistingQuestion(formQuestion, tableQuestion) {
  const tableQuestionData = parseTableQuestion(tableQuestion);
  Logger.log(`Table/form: ${tableQuestionData.question} / ${formQuestion.getTitle()}`,)

  // Every item is generic by default, so the type must be set
  switch (tableQuestionData.type) {
    case 'Short answer':
      formItem = formQuestion.asTextItem();
      break;
    case 'Long answer':
      formItem = formQuestion.asParagraphTextItem();
      break;
    case 'Dropdown list':
      formItem = formQuestion.asListItem();
      break;
    case 'Multiple choice':
      formItem = formQuestion.asMultipleChoiceItem();
      break;
    case 'Checkbox':
      formItem = formQuestion.asCheckboxItem();
      break;
    case 'Linear scale':
      formItem = formQuestion.asScaleItem();
      break;
  }

  // Change title if necessary
  if (formItem.getTitle() !== tableQuestionData.question) {
    Logger.log(`Changing title from ${formItem.getTitle()} to ${tableQuestionData.question}`)
    formItem.setTitle(tableQuestionData.question);
  }

  // It's easier to reset the all the choices than see which have changed
  if (tableQuestionData.options) {
    Logger.log(`Processing options`)
    // https://developers.google.com/apps-script/reference/forms/choice?hl=en
    setQuestionChoices({ item: formItem, options: tableQuestionData.options, allowCustomOther: getAllowCustomOther() })
  } else if (tableQuestionData.lower || tableQuestionData.upper) {
    Logger.log(`Processing bounds and labels`)
    formItem
      .setBounds(tableQuestionData.lower.bound, tableQuestionData.upper.bound)
      .setLabels(tableQuestionData.lower.label, tableQuestionData.upper.label)
  }

  Logger.log('Updating isRequired')
  formItem.setRequired(tableQuestionData.required);

  Logger.log(`Finished processing ${tableQuestionData.question}`)
}

// Add question

function addNewQuestion(form, table) {
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
    addQuestionWithOptions({ item, table })
  } else if (type === 'Multiple choice') {
    item = form.addMultipleChoiceItem();
    addQuestionWithOptions({ item, table })
  } else if (type === 'Checkbox') {
    item = form.addCheckboxItem();
    addQuestionWithOptions({ item, table })
  } else if (type === 'Linear scale') {
    item = form.addScaleItem();
    const lower = table.getCell(2, 1).getText().split(': ')
    const upper = table.getCell(3, 1).getText().split(': ')
    item.setTitle(question)
      .setBounds(lower[0], upper[0])
      .setLabels(lower[1], upper[1])
  }

  item.setRequired(isRequired);
}

function setQuestionChoices({ item, options, allowCustomOther }) {
  let choices = generateFormQuestionOptions({ item, options, allowCustomOther })
  let showOther = allowCustomOther && options.toLowerCase().includes('other');

  item.setChoices(choices)
  if (allowCustomOther) {
    item.showOtherOption(showOther);
  }

  return item;
}

function addQuestionWithOptions({ item, table }) {
  const question = table.getCell(0, 1).getText()
  const type = table.getCell(1, 1).getText()
  const options = table.getCell(2, 1).getText()
  const allowCustomOther = getAllowCustomOther(type);

  let showOther = allowCustomOther && options.toLowerCase().includes('other');
  let choices = generateFormQuestionOptions({ item, options, allowCustomOther })

  item.setTitle(question).setChoices(choices)
  if (allowCustomOther) {
    item.showOtherOption(showOther);
  }
}

function createSection(form) {
  const section = form.addPageBreakItem().setTitle('Page Two');
}
