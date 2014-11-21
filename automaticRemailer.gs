// function to get the number of threads that have unread messages in them.
function InboxUnreadCount()
{
  var unreadthreads = GmailApp.getInboxUnreadCount()
  
  return unreadthreads
}

// function to return all inbox threads.
function totalInboxThreads()
{
  var threads = GmailApp.getInboxThreads();
  return threads

}

function unreadMails(inboxthreads)
// returns a array of unread threads
{
  var inboxthread_length = inboxthreads.length
  var mailIdArray = []
  for (i = 0; i < inboxthread_length; i++)
  {
    if (inboxthreads[i].isUnread() == true)
    {
      mailIdArray.push(inboxthreads[i].getId())
    }
  }
  return mailIdArray
}

function forwardMails(unreadMailArray,listOfReceipients)
// forwarding of unread mails only
{
  var unreadMailCount = unreadMailArray.length
  var noOfReceipients = listOfReceipients.length
  
  for (i = 0; i < unreadMailCount; i++)
  {
    var mailToSend = GmailApp.getMessageById(unreadMailArray[i])
    for (j = 0; j < noOfReceipients ; j++)
    {
      mailToSend.forward(listOfReceipients[j])
    }
  }
}

function markReadForwardedMails(unreadMailArray)
// marks the given mail array as read.
{
  var unreadMailCount = unreadMailArray.length
  for (i = 0; i < unreadMailCount; i++)
  {
    GmailApp.markThreadRead(GmailApp.getThreadById((unreadMailArray[i])))
  }
}

function forwardMarkRead()
// function which combines the forward and markread functions
// this function will have a trigger.
{
  var mailIdlist = [<list of receipients go here>]
  var threads = totalInboxThreads()
  var unreadmail = unreadMails(threads)
  var currentDate = new Date()
  var stringDate = currentDate.getDate() + "-" + (1 + currentDate.getMonth()) + "-" + currentDate.getFullYear() + "-" + currentDate.getHours() + "-" + currentDate.getMinutes() + "-" + currentDate.getSeconds()
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = spreadsheet.getSheets()[0]
  var req_range1 = sheet.getRange(1, 2, 1, 1) //last run time
  var req_range2 = sheet.getRange(2, 2, 1, 1) //number of forwardedmails
  if (unreadmail.length > 0)
  {
    forwardMails(unreadmail,mailIdlist)
    markReadForwardedMails(unreadmail)
  }  
  
  else
  {
    Logger.log(unreadmail.length)
  }
  Logger.log(stringDate)
  req_range1.setValue(stringDate)
  req_range2.setValue(Math.floor(unreadmail.length))
}
