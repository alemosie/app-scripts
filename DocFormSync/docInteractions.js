const ui = DocumentApp.getUi();
const body = DocumentApp.getActiveDocument().getBody();

let form = ''; // Set on doc open, when updates are run

const onOpen = () => {
  ui.createMenu('✨ Form builder')
    .addSubMenu(
      ui.createMenu('Insert new question')
        .addItem('Short answer', 'insertShortAnswerQuestionTemplate')
        .addItem('Long answer (paragraph)', 'insertLongAnswerQuestionTemplate')
        .addItem('Dropdown list (single-select)', 'insertDropdownQuestionTemplate')
        .addItem('Multiple choice (single-select)', 'insertMultipleChoiceQuestionTemplate')
        .addItem('Checkbox (multi-select)', 'insertCheckboxQuestionTemplate')
        .addItem('Linear scale', 'insertLinearScaleQuestionTemplate')
    )
    .addSeparator()
    .addItem('Doc ➡️ Form', 'syncDocToForm')
    .addToUi();

  const formDataTable = getFormDataTable();
  form = getForm(formDataTable);
  updateFormTitleOnDoc(form, formDataTable);
}

const getFormDataTable = () => {
  return body.getTables()[0];
}

const getForm = (formDataTable) => {
  if (!formDataTable || !formDataTable?.getNumRows() === 0) {
    throw 'Cannot find form URL'
  }
  const formUrl = formDataTable.getCell(0, 1).getText()
  return FormApp.openByUrl(formUrl);
}

const updateFormTitleOnDoc = (form, formDataTable) => {
  formDataTable.getCell(1, 1).setText(form.getTitle())
}

// CREATE AND INSERT QUESTION TEMPLATES

const insertQuestionTemplateIntoDocument = (type) => {
  Logger.log(`Inserting ${type} question template`)

  let rowsData = [
    ['Question', ''],
    ['Type', type]
  ];

  if (type === 'Linear scale') {
    rowsData.push(['Lower bound', '1'])
    rowsData.push(['Lower label', 'Least'])
    rowsData.push(['Upper bound', '5'])
    rowsData.push(['Upper label', 'Most'])
  }

  else if (!['Short answer', 'Long answer'].includes(type)) {
    rowsData.push(['Options', ''])
  }

  rowsData.push(['Required?', ''])

  let table = null;

  // Insert table at cursor position or if no cursor, at bottom
  const cursor = DocumentApp.getActiveDocument().getCursor();
  if (cursor) {
    const element = cursor.getElement();
    if (element.getParent().getType() === DocumentApp.ElementType.BODY_SECTION) {
      table = body.insertTable(element.getParent().getChildIndex(element) + 1, rowsData);
    } else {
      let tableParent = element.getParent();
      while (tableParent.getType() !== DocumentApp.ElementType.TABLE) {
        tableParent = tableParent.getParent();
      }
      table = body.insertTable(body.getChildIndex(tableParent.getNextSibling()) + 1, rowsData)
    }
  } else {
    table = body.appendTable(rowsData);
  }

  let style = {};
  style[DocumentApp.Attribute.BOLD] = true;
  for (let i = 0; i < table.getNumRows(); i++) {
    table.getCell(i, 0).setAttributes(style).setWidth(80)
  }
}
const insertShortAnswerQuestionTemplate = () => {
  insertQuestionTemplateIntoDocument('Short answer')
}
const insertLongAnswerQuestionTemplate = () => {
  insertQuestionTemplateIntoDocument('Long answer')
}
const insertDropdownQuestionTemplate = () => {
  insertQuestionTemplateIntoDocument('Dropdown list')
}
const insertMultipleChoiceQuestionTemplate = () => {
  insertQuestionTemplateIntoDocument('Multiple choice')
}
const insertCheckboxQuestionTemplate = () => {
  insertQuestionTemplateIntoDocument('Checkbox')
}
const insertLinearScaleQuestionTemplate = () => {
  insertQuestionTemplateIntoDocument('Linear scale')
}