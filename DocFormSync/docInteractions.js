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
    .addItem('Doc ➡ Form', 'syncDocToForm')
    .addToUi();

  form = getForm();
}

const getForm = () => {
  const formUrlTable = body.getTables()[0];
  // TODO: Better logic for identifying URL table
  if (!formUrlTable || !formUrlTable?.getNumRows() >= 1) {
    throw 'Cannot find form URL'
  }
  const formUrl = formUrlTable.getCell(0, 1).getText()
  return FormApp.openByUrl(formUrl);
}

// CREATE AND INSERT QUESTION TEMPLATES

const insertQuestionTemplateIntoDocument = (type) => {
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

  const table = body.appendTable(rowsData);

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