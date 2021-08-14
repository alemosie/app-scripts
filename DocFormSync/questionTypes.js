const QUESTION_TYPES = [
  {
    documentQuestionType: 'Short answer',
    formItemType: FormApp.ItemType.TEXT,
    addItemToForm: (form) => form.addTextItem(),
    convertToFormItem: (item) => item.asTextItem()
  },
  {
    documentQuestionType: 'Long answer',
    formItemType: FormApp.ItemType.PARAGRAPH_TEXT,
    addItemToForm: (form) => form.addParagraphTextItem(),
    convertToFormItem: (item) => item.asParagraphTextItem()
  },
  {
    documentQuestionType: 'Dropdown list',
    formItemType: FormApp.ItemType.LIST,
    addItemToForm: (form) => form.addListItem(),
    convertToFormItem: (item) => item.asListItem(),
    allowCustomOther: false,
  },
  {
    documentQuestionType: 'Multiple choice',
    formItemType: FormApp.ItemType.MULTIPLE_CHOICE,
    addItemToForm: (form) => form.addMultipleChoiceItem(),
    convertToFormItem: (item) => item.asMultipleChoiceItem(),
    allowCustomOther: true,
  },
  {
    documentQuestionType: 'Checkbox',
    formItemType: FormApp.ItemType.CHECKBOX,
    addItemToForm: (form) => form.addCheckboxItem(),
    convertToFormItem: (item) => item.asCheckboxItem(),
    allowCustomOther: true,
  },
  {
    documentQuestionType: 'Linear scale',
    formItemType: FormApp.ItemType.SCALE,
    addItemToForm: (form) => form.addScaleItem(),
    convertToFormItem: (item) => item.asScaleItem()
  }
]

const getItemByType = (type) => {
  for (let i = 0; i < QUESTION_TYPES.length; i++) {
    const { documentQuestionType, itemType } = QUESTION_TYPES[i];
    if (type === documentQuestionType || type === itemType) {
      return QUESTION_TYPES[i]
    }
  }

  // If no results, surface an error back to the user that the question type is invalid
  throw 'Invalid question type'
}