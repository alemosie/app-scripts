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

function getFormItemsFromDocument() {
  const formElements = [];
  let withinQuestionSection = false; // Only slurp tables that are under "Questions" header

  for (let i = 0; i < body.getNumChildren(); i++) {
    const child = body.getChild(i);
    const childType = child.getType();

    // Create form sections out of every Heading 3
    if (
      childType === DocumentApp.ElementType.PARAGRAPH &&
      child.getHeading() === DocumentApp.ParagraphHeading.HEADING2 &&
      child.asParagraph().getText() == 'Questions'
    ) {
      withinQuestionSection = true;
    } else if (childType === DocumentApp.ElementType.PARAGRAPH && child.getHeading() === DocumentApp.ParagraphHeading.HEADING3) {
      formElements.push(child.asParagraph())
    } else if (withinQuestionSection && childType === DocumentApp.ElementType.TABLE) {
      formElements.push(child.asTable())
    }
  }
  return formElements;
}

function updateForm() {
  const formItemsFromDocument = getFormItemsFromDocument();

  if (!formUrl) {
    formUrl = setFormUrl();
  }
  const form = FormApp.openByUrl(formUrl);
  const existingFormItems = form.getItems()

  // Delete any extra questions
  if (formItemsFromDocument.length < existingFormItems.length) {
    const itemsToDelete = existingFormItems.slice(formItemsFromDocument.length)
    Logger.log(`Deleting ${itemsToDelete.length} questions`)
    itemsToDelete.forEach((question) => {
      form.deleteItem(question)
    });
  }

  for (let i = 0; i < formItemsFromDocument.length; i++) {
    const documentItem = formItemsFromDocument[i];

    // If there's an existing question in that slot, edit it
    if (i < existingFormItems.length) {
      const formItem = existingFormItems[i];

      try {
        updateExistingItem(formItem, documentItem);
      }
      // If the question has a different type, we have to delete, create new, and move
      catch {
        try {
          Logger.log(`Replacing ${formItem.getTitle()}`);
          replaceExistingItem(i, form, documentItem);
        } catch(err) {
          throw err;
        }
      }

    // Otherwise, add a new question
    } else {
      addNewItem(form, documentItem);
    }
  }
}

// Questions helpers

function parseDocumentQuestion(table) {
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
  Logger.log(questionData)
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

function updateExistingItem(formItem, documentItem) {
  // For non-tables, create a section (page break)
  if (documentItem.getType() === DocumentApp.ElementType.PARAGRAPH) {
    formItem = formItem.asPageBreakItem();
    formItem.setTitle(documentItem.getText());
    return formItem;
  }

  // Else, process the table elements as questions
  const documentQuestionData = parseDocumentQuestion(documentItem);
  if (!documentQuestionData.question && !documentQuestionData.type) {
    return;
  }

  if (formItem.getTitle() !== documentQuestionData.question) {
    Logger.log(`Updating title to ${documentQuestionData.question}`)
    formItem.setTitle(documentQuestionData.question);
  }

  // Every item is generic by default, so the type must be set
    switch (documentQuestionData.type) {
      case 'Short answer':
        formItem = formItem.asTextItem();
        break;
      case 'Long answer':
        formItem = formItem.asParagraphTextItem();
        break;
      case 'Dropdown list':
        formItem = formItem.asListItem();
        break;
      case 'Multiple choice':
        formItem = formItem.asMultipleChoiceItem();
        break;
      case 'Checkbox':
        formItem = formItem.asCheckboxItem();
        break;
      case 'Linear scale':
        formItem = formItem.asScaleItem();
        break;
      default:
        throw 'Invalid question type. Please choose one of: Short answer, Long answer, Dropdown list, Multiple choice, Checkbox, Linear scale'
  }

  // It's easier to reset the all the choices than see which have changed
  if (documentQuestionData.options) {
    Logger.log(`Processing options`)
    const allowCustomOther = getAllowCustomOther(documentQuestionData.type);
    // https://developers.google.com/apps-script/reference/forms/choice?hl=en
    setQuestionChoices({
      item: formItem,
      options: documentQuestionData.options,
      allowCustomOther
    });
  } else if (documentQuestionData.lower || documentQuestionData.upper) {
    Logger.log(`Processing bounds and labels`)
    formItem.setBounds(documentQuestionData.lower.bound, documentQuestionData.upper.bound)
      .setLabels(documentQuestionData.lower.label, documentQuestionData.upper.label)
  }

  Logger.log('Updating isRequired')
  formItem.setRequired(documentQuestionData.required);

  Logger.log(`Finished processing ${documentQuestionData.question}`)
  return returnItem;
}

// Add question

function addNewItem(form, documentItem) {
  // For non-tables, create a section (page break)
  if (documentItem.getType() === DocumentApp.ElementType.PARAGRAPH) {
    formItem = form.addPageBreakItem();
    formItem.setTitle(documentItem.getText());
    return formItem;
  }

  const documentQuestionData = parseDocumentQuestion(documentItem);
  const allowCustomOther = getAllowCustomOther(documentQuestionData.type);

  if (!documentQuestionData.question && !documentQuestionData.type) {
    return;
  }

  Logger.log(`Adding ${documentQuestionData.question}`)

  let formItem = null;
  switch (documentQuestionData.type) {
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
        options: documentQuestionData.options,
        allowCustomOther
      });
      break;
    case 'Multiple choice':
      formItem = form.addMultipleChoiceItem();
      setQuestionChoices({
        item: formItem,
        options: documentQuestionData.options,
        allowCustomOther
      });
      break;
    case 'Checkbox':
      formItem = form.addCheckboxItem();
      setQuestionChoices({
        item: formItem,
        options: documentQuestionData.options,
        allowCustomOther
      });
      break;
    case 'Linear scale':
      formItem = form.addScaleItem();
      formItem.setBounds(documentQuestionData.lower.bound, documentQuestionData.upper.bound)
        .setLabels(documentQuestionData.lower.label, documentQuestionData.upper.label)
      break;
    default:
      throw 'Invalid question type. Please choose one of: Short answer, Long answer, Dropdown list, Multiple choice, Checkbox, Linear scale'
  }

  formItem.setTitle(documentQuestionData.question);
  formItem.setRequired(documentQuestionData.required);

  Logger.log(`Finished adding ${documentQuestionData.question}`)
  return formItem;
}

function replaceExistingItem(index, form, documentItem) {
  form.deleteItem(index);
  const newItem = addNewItem(form, documentItem);

  // Exception: The parameters (FormApp.MultipleChoiceItem,number) don't match the method signature for FormApp.Form.moveItem.
  // We need to access the item before it has been assigned a type to move it to the right place
  const newItemIndex = newItem.getIndex()
  const newElementItem = form.getItems()[newItemIndex]
  form.moveItem(newElementItem, index);
}
