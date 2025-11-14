const EMAIL_VALIDATION = CardService.newValidation().setInputType(CardService.InputType.EMAIL)
function createAddEmailCard() {
  const card = CardService.newCardBuilder()
  const header = CardService.newCardHeader().setTitle("Add emails to a tag")
  card.setHeader(header)

  const textInput = CardService.newTextInput()
    .setFieldName('customer-email')
    .setTitle('customerEmail')
    .setValidation(EMAIL_VALIDATION)
}
function createLabelsSelectionInput() {
  const labels = GmailApp.getUserLabels().filter((label) => {
    return label.getName().includes('Client Communication')
  })
  const selectionInput = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setTitle('Labels')
    .setFieldName('label-name')
  
  for(const label of labels) {
    const labelName = label.getName()
    selectionInput.addItem(labelName.split('/')[1], labelName, false)
  }
  return selectionInput
}