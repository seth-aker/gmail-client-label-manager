const CACHE_KEY = 'emailLabels'
type Label = string
type Email = string

function onEmailOpen(event: GoogleAppsScript.Addons.EventObject) {
  const accessToken = event.gmail?.accessToken
  const threadId = event.gmail?.threadId
  if(!accessToken || !threadId) {
    return
  }
  GmailApp.setCurrentMessageAccessToken(accessToken)
  const openThread = GmailApp.getThreadById(threadId)

  const emailLabelCache = new Map<Email, Label>(JSON.parse(PropertiesService.getUserProperties().getProperty(CACHE_KEY) ?? ''))

  let threadEmails: Set<string> = getAddressesInThread(openThread)

  for(const email of threadEmails) {
    const labelName = emailLabelCache.get(email)
    if(labelName) {
      const label = GmailApp.getUserLabelByName(labelName)
      openThread.addLabel(label)
      break;
    }  
  }
  if(openThread.getLabels().length > 0)  return;

  // Iterate through existing labels
  for(const label of GmailApp.getUserLabels()) {
    const threads = label.getThreads()
    // get all threads using this label
    for(const thread of threads) {
      const addresses = getAddressesInThread(thread)
      // get the addresses associated with said label and add them to the cache if they aren't there already
      for(const address of addresses) {
        if(!emailLabelCache.has(address)) {
          emailLabelCache.set(address, label.getName())
        }
        if(openThread.getLabels().length < 1 && threadEmails.has(address)) {
          openThread.addLabel(label)
        }
      }
    }
  }
  
  updateEmailLabelCache(emailLabelCache)

  // for(const message of messages) {
  //   const messageSubject = message.getSubject()
  //   const messageSender = message.getFrom()
  //   Logger.log(`Viewing message: ${messageSubject} from ${messageSender}`)
  //   if(!messageSubject.includes("Implementation transferred to you") || messageSender !== DOUG_EMAIL) {
  //     continue
  //   }
  //   const messageBody = message.getBody()
  //   const customerName = findCustomerName(messageBody)
    
  //   const isOps = messageBody.includes("OPS")
  //   createLabel(customerName, isOps)
  // }

}
function getAddressesInThread(thread: GoogleAppsScript.Gmail.GmailThread) {
  const messages = thread.getMessages()
  const addresses = new Set<string>()
  for(const message of messages) {
    const fromEmails = message.getFrom()
    const toEmails = message.getTo().split(',')
    const ccs = message.getCc().split(',')
    const emails = [fromEmails, ...toEmails, ...ccs].filter(each => !each.includes('@trimble.com'))
    emails.forEach(email => addresses.add(email))
  }
  return addresses
}

function updateEmailLabelCache(cache: Map<Email, Label>) {
  const userProperties = PropertiesService.getUserProperties()
  const cacheString = JSON.stringify(cache)
  userProperties.setProperty(CACHE_KEY, cacheString)
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
