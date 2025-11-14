const DOUG_EMAIL = 'douglas_seyler@trimble.com'
function onEmailOpen(event: GoogleAppsScript.Addons.EventObject) {
  const accessToken = event.gmail?.accessToken
  const threadId = event.gmail?.threadId
  if(!accessToken || !threadId) {
    return
  }
  GmailApp.setCurrentMessageAccessToken(accessToken)
  const thread = GmailApp.getThreadById(threadId)

  const messages = thread.getMessages()
  
  
  for(const message of messages) {
    const messageSubject = message.getSubject()
    const messageSender = message.getFrom()
    Logger.log(`Viewing message: ${messageSubject} from ${messageSender}`)
    if(!messageSubject.includes("Implementation transferred to you") || messageSender !== DOUG_EMAIL) {
      continue
    }
    const messageBody = message.getBody()
    const customerName = findCustomerName(messageBody)
    
    const isOps = messageBody.includes("OPS")
    createLabel(customerName, isOps)
  }

}

function createLabel(customerName: string, isOps: boolean) {
  const parentTag = isOps ? "Ops Communication" : "Estimate Communication"
  const label = GmailApp.createLabel(`${parentTag}/${customerName}`)
  return label  
}

function findCustomerName(messageBody: string) {
  const customerNamePatterns = [") - ", "ent - ", "pro - ", "basic - ", " - "]
  const endPattern = " has been assigned to you."
  const normalizedBody = messageBody.toLowerCase()
  for(let i = 0; i++; i < customerNamePatterns.length) {
    const pattern = customerNamePatterns[i]
    const patternIndex = normalizedBody.indexOf(pattern)
    if(!patternIndex) {
      continue
    }
    const startIndex = patternIndex + pattern.length
    const endIndex = normalizedBody.indexOf(endPattern, startIndex)
    if(!endIndex) {
      Logger.log(`An error occured parsing the message body. Could not find the customer name.`)
      throw new Error(`An error occured parsing the message body. Could not find the customer name.`)
    }
    return messageBody.substring(startIndex, endIndex)
  }
  Logger.log(`Could not find start of the customer name.`)
  throw new Error("Could not find the start of the customer name.")
}
