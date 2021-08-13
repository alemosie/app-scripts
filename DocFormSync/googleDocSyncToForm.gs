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

  // Delete any extra questions
  if (questionTables.length < existingFormQuestions.length) {
    const questionsToDelete = existingFormQuestions.slice(questionTables.length)
    Logger.log(`Deleting ${questionsToDelete.length} questions`)
    questionsToDelete.forEach((question) => {
      form.deleteItem(question)
    });
  }

  for (let i = 0; i < questionTables.length; i++) {
    const tableQuestion = questionTables[i];

    // If there's an existing question in that slot, edit it
    if (i < existingFormQuestions.length) {
      const formItem = existingFormQuestions[i];

      try {
        updateExistingQuestion(formItem, tableQuestion);
      }
      // If the question has a different type, we have to delete, create new, and move
      catch {
        try {
          Logger.log(`Replacing ${formItem.getTitle()}`);
          replaceExistingQuestion(i, form, tableQuestion);
        } catch(err) {
          throw err;
        }
      }

    // Otherwise, add a new question
    } else {
      addNewQuestion(form, tableQuestion);
    }
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

function setQuestionChoices({ item, options, allowCustomOther }) {
  let choices = options.trim().split('\n').reduce((result, option) => {
    if (
      // If "other" is a special option for custom answer, skip adding "other" as an set option to choose
      (option.toLowerCase() === 'other' && !allowCustomOther) ||
      (option.toLowerCase() !== 'other')) {
      result.push(item.createChoice(option));
    }
    return result;
  }, []);
  item.setChoices(choices);

  let showOther = allowCustomOther && options.toLowerCase().includes('other');
  if (allowCustomOther) {
    item.showOtherOption(showOther);
  }

  return item;
}



// Edit question

function updateExistingQuestion(formQuestion, tableQuestion) {
  const tableQuestionData = parseTableQuestion(tableQuestion);
  Logger.log(`Table/form: ${tableQuestionData.question} / ${formQuestion.getTitle()}`,)

  if (formQuestion.getTitle() !== tableQuestionData.question) {
    Logger.log(`Updating title to ${tableQuestionData.question}`)
    formItem.setTitle(tableQuestionData.question);
  }

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
      default:
        throw 'Invalid question type. Please choose one of: Short answer, Long answer, Dropdown list, Multiple choice, Checkbox, Linear scale'
  }

  // It's easier to reset the all the choices than see which have changed
  if (tableQuestionData.options) {
    Logger.log(`Processing options`)
    const allowCustomOther = getAllowCustomOther(tableQuestionData.type);
    // https://developers.google.com/apps-script/reference/forms/choice?hl=en
    setQuestionChoices({
      item: formItem,
      options: tableQuestionData.options,
      allowCustomOther
    });
  } else if (tableQuestionData.lower || tableQuestionData.upper) {
    Logger.log(`Processing bounds and labels`)
    formItem
      .setBounds(tableQuestionData.lower.bound, tableQuestionData.upper.bound)
      .setLabels(tableQuestionData.lower.label, tableQuestionData.upper.label)
  }

  Logger.log('Updating isRequired')
  formItem.setRequired(tableQuestionData.required);

  Logger.log(`Finished processing ${tableQuestionData.question}`)
  return returnItem;
}

// Add question

function addNewQuestion(form, tableQuestion) {
  const tableQuestionData = parseTableQuestion(tableQuestion);
  const allowCustomOther = getAllowCustomOther(tableQuestionData.type);

  if (!tableQuestionData.question && !tableQuestionData.type) {
    return;
  }

  Logger.log(`Adding ${tableQuestionData.question}`)

  let formItem = null;
  switch (tableQuestionData.type) {
    case 'Short answer':
      formItem = form.addTextItem();
      break;
    case 'Long answer':
      formItem = form.addParagraphTextItem();
      break;
    case 'Dropdown list':
      formItem = form.addListItem();
      setQuestionChoices({
        item: formItem,
        options: tableQuestionData.options,
        allowCustomOther
      });
      break;
    case 'Multiple choice':
      formItem = form.addMultipleChoiceItem();
      setQuestionChoices({
        item: formItem,
        options: tableQuestionData.options,
        allowCustomOther
      });
      break;
    case 'Checkbox':
      formItem = form.addCheckboxItem();
      setQuestionChoices({
        item: formItem,
        options: tableQuestionData.options,
        allowCustomOther
      });
      break;
    case 'Linear scale':
      formItem = form.addScaleItem();
      break;
    default:
      throw 'Invalid question type. Please choose one of: Short answer, Long answer, Dropdown list, Multiple choice, Checkbox, Linear scale'
  }

  formItem.setTitle(tableQuestionData.question);
  formItem.setRequired(tableQuestionData.required);

  Logger.log(`Finished adding ${tableQuestionData.question}`)
  return formItem;
}

function replaceExistingQuestion(index, form, tableQuestion) {
  form.deleteItem(index);
  const newQuestion = addNewQuestion(form, tableQuestion);

  // Exception: The parameters (FormApp.MultipleChoiceItem,number) don't match the method signature for FormApp.Form.moveItem.
  // We need to access the item before it has been assigned a type to move it to the right place
  const newQuestionIndex = newQuestion.getIndex()
  const newQuestionItem = form.getItems()[newQuestionIndex]
  form.moveItem(newQuestionItem, index);
}
