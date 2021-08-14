// STEP 1: FIND ALL ELEMENTS IN DOCUMENT TO SYNC TO DOC

const getItemsToSyncFromDoc = () => {
  let itemsToSync = [];

  // In the document, section headers are heading 2
  const isFormSectionHeader = (element) => {
    return (
      element.getType() === DocumentApp.ElementType.PARAGRAPH &&
      element.getHeading() === DocumentApp.ParagraphHeading.HEADING2
    )
  }

  // As we traverse the body elements, we only want to identify items
  // that are within the "Questions" section for the form sync
  const isQuestionsHeader = (element) => {
    return (
      element.getType() === DocumentApp.ElementType.PARAGRAPH &&
      element.getHeading() === DocumentApp.ParagraphHeading.HEADING1 &&
      element.asParagraph().getText() == 'Questions'
    )
  }
  let withinQuestionSection = false;

  for (let i = 0; i < body.getNumChildren(); i++) {
    const child = body.getChild(i);
    const childType = child.getType();

    if (isQuestionsHeader(child)) {
      withinQuestionSection = true;
      // Our form section headers (page breaks)
    } else if (withinQuestionSection && isFormSectionHeader(child)) {
      itemsToSync.push(child.asParagraph())
      // Our form questions
    } else if (withinQuestionSection && childType === DocumentApp.ElementType.TABLE) {
      itemsToSync.push(child.asTable())
    }
  }

  if (itemsToSync.length === 0) {
    throw 'Could not find any questions in the document to sync to the form'
  }
  return itemsToSync;
}

// STEP 2: PARSE ELEMENTS WE FOUND IN THE DOCUMENT

const parseDocumentQuestionItem = (table) => {
  // Convert a question table in the document into an object
  // that we'll use to populate the form version of the question

  const formatKey = (text) => {
    let key = text.replace('?', '').toLowerCase();
    if (key.includes(' ')) {
      // camelCase keys
      const keySplit = key.split(' ');
      return keySplit[0] + keySplit[1][0].toUpperCase() + keySplit[1].substr(1,)
    } else {
      return key;
    }
  }

  let questionData = {};
  for (let i = 0; i < table.getNumRows(); i++) {
    const key = formatKey(table.getCell(i, 0).getText())
    const value = table.getCell(i, 1).getText();
    questionData[key] = value;
  }

  // Users can denote that a question be required with either "Y" or "Yes" (case insensitive)
  const isValidRequired = (input) => ['y', 'yes'].includes(input.toLowerCase())
  questionData.required = isValidRequired(questionData?.required);

  return questionData;
}

const containsScaleData = (data) => {
  const { lowerBound, upperBound, lowerLabel, upperLabel } = data;
  return (lowerBound && upperBound && lowerLabel && upperLabel)
}

// STEP 3a: INSERT NEW FORM ITEMS

const setFormItemChoices = ({ item, options, allowCustomOther }) => {
  Logger.log(`Adding choices to question: ${item.getTitle()}`)

  let choices = options.trim().split('\n').reduce((result, option) => {
    if (
      // If "other" is a special option for custom answer, skip adding "other" as a set option to choose
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

const addNewFormItem = (form, documentItem) => {
  let formItem = null;

  // For non-tables, create a section (page break)
  if (documentItem.getType() === DocumentApp.ElementType.PARAGRAPH) {
    const sectionText = documentItem.getText()
    Logger.log(`Adding section to form: ${sectionText}`)
    formItem = form.addPageBreakItem();
    formItem.setTitle(sectionText);
    return formItem;
  }

  // If not a section, add a question
  const documentQuestionData = parseDocumentQuestionItem(documentItem);
  if (!documentQuestionData.question && !documentQuestionData.type) {
    return;
  }

  Logger.log(`Adding question to form: ${documentQuestionData.question}`)

  const itemObject = getItemByType(documentQuestionData.type);
  formItem = itemObject.addItemToForm(form)
  if (documentQuestionData.options) {
    Logger.log(`Processing options`)
    setFormItemChoices({
      item: formItem,
      options: documentQuestionData.options,
      allowCustomOther: itemObject.allowCustomOther
    });
  } else if (containsScaleData(documentQuestionData)) {
    Logger.log(`Processing bounds and labels`)
    formItem.setBounds(documentQuestionData.lowerBound, documentQuestionData.upperBound)
      .setLabels(documentQuestionData.lowerLabel, documentQuestionData.upperLabel)
  }

  formItem.setTitle(documentQuestionData.question);
  formItem.setRequired(documentQuestionData.required);

  Logger.log(`Finished adding question to form: ${documentQuestionData.question}`)
  return formItem;
}

// STEP 3b: EDIT EXISTING ITEMS

const updateExistingFormItem = (formItem, documentItem) => {
  // For non-tables, update as section (page break)
  if (documentItem.getType() === DocumentApp.ElementType.PARAGRAPH) {
    formItem = formItem.asPageBreakItem();
    formItem.setTitle(documentItem.getText());
    return formItem;
  }

  // Else, process the table elements as questions
  const documentQuestionData = parseDocumentQuestionItem(documentItem);
  if (!documentQuestionData.question && !documentQuestionData.type) {
    return;
  }

  Logger.log(`Updating existing form question: ${formItem.getTitle()}`)

  if (formItem.getTitle() !== documentQuestionData.question) {
    Logger.log(`Updating title to ${documentQuestionData.question}`)
    formItem.setTitle(documentQuestionData.question);
  }

  // Every item is generic by default, so the type must be set
  const itemObject = getItemByType(documentQuestionData.type);
  formItem = itemObject.convertToFormItem(formItem);

  // It's easier to reset the all the choices than see which have changed
  if (documentQuestionData.options) {
    Logger.log(`Processing options`)
    setFormItemChoices({
      item: formItem,
      options: documentQuestionData.options,
      allowCustomOther: itemObject.allowCustomOther
    });
  } else if (containsScaleData(documentQuestionData)) {
    Logger.log(`Processing bounds and labels`)
    formItem.setBounds(documentQuestionData.lowerBound, documentQuestionData.upperBound)
      .setLabels(documentQuestionData.lowerLabel, documentQuestionData.upperLabel)
  }

  Logger.log('Updating isRequired')
  formItem.setRequired(documentQuestionData.required);

  Logger.log(`Finished updating form question: ${documentQuestionData.question} (${formItem.getTitle()})`)
  return returnItem;
}

const replaceExistingFormItem = (index, form, documentItem) => {
  form.deleteItem(index);
  const newItem = addNewFormItem(form, documentItem);

  // Exception: The parameters (FormApp.MultipleChoiceItem,number) don't match the method signature for FormApp.Form.moveItem.
  // We need to access the item before it has been assigned a type to move it to the right place
  const newItemIndex = newItem.getIndex()
  const newElementItem = form.getItems()[newItemIndex]
  form.moveItem(newElementItem, index);
}



// Update the form

const syncDocToForm = () => {
  Logger.log(`Running: doc -> form`)
  const formItemsFromDoc = getItemsToSyncFromDoc();

  if (!form) {
    form = getForm();
  }
  const existingFormItems = form.getItems()

  // Delete any extra questions
  if (formItemsFromDoc.length < existingFormItems.length) {
    const itemsToDelete = existingFormItems.slice(formItemsFromDoc.length)
    Logger.log(`Deleting ${itemsToDelete.length} questions`)
    itemsToDelete.forEach((question) => {
      form.deleteItem(question)
    });
  }

  for (let i = 0; i < formItemsFromDoc.length; i++) {
    const documentItem = formItemsFromDoc[i];

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
          replaceExistingFormItem(i, form, documentItem);
        } catch (err) {
          throw err;
        }
      }

      // Otherwise, add a new question
    } else {
      addNewFormItem(form, documentItem);
    }
  }
}