
const ss = SpreadsheetApp.getActiveSpreadsheet();

const TAB_NAME = {
  PERMISSIONS: "Permissions",
  GROUPS: "Groups"
};

const PERMISSIONS_FIRST_DATA_ROW = 2;

const VARS = {
  PERMISSIONS: getScriptVars(TAB_NAME.PERMISSIONS),
  GROUPS: getScriptVars(TAB_NAME.GROUPS)
};

const EMAIL_FWD_ADDRS = buildEmailFwdAddrs();
const EMAIL_PLUS_ADDRS = buildEmailPlusAddrs();
const TITLE_PREFIX = `Water Volleyball`;
const TITLES = buildTitles();

const MAX_EMAIL_THREADS_TO_GET = 1;
const MAX_MEMBERS_PER_MAILING_LIST = 1000;
const MAX_PEOPLE_GET_BATCH_SIZE = 50;
const ALL_EMAILS_LIST = "All Active Members";
var ALL_MEMBER_EMAILS = getEmailsForMailingList(ALL_EMAILS_LIST);

/**
 * Function getScriptVars
 * 
 * Reads the script variables from the "Variables" tab and returns them.
 * 
 * @param    {string}  tabName  Name of a tab to get variables from.
 * @returns  {object}           An object containing the file IDs by file name.
 */
function getScriptVars(tabName) {
  const tab = ss.getSheetByName(tabName);
  const values = tab.getRange(PERMISSIONS_FIRST_DATA_ROW, 1, tab.getLastRow(), tab.getLastColumn()).getValues();
  var obj = {};
  var varName, value;

  for (let idx = 0; idx < values.length; idx++) {
    varName = values[idx][0].toString().trim();
    value = values[idx][1].toString().trim();
    if (varName.length > 0) {
      obj[varName] = value;
    }
  }

  return obj;
}


/**
 * Function getScriptVar
 * 
 * Gets the value of a script variable that was read from the spreadsheet.
 * 
 * @param    {object} tabName  Name of the tab containing the variables
 * @param    {string} varName  Name of the variable to retrieve
 * @returns  {string}          The value of a variable contained in vars.
 */
function getScriptVar(tabName, varName) {
  const varsIdx = tabName.toUpperCase();
  if (!VARS.hasOwnProperty(varsIdx)) {
    const msg = "Could not find the " + varsIdx + " tab name in VARS";
    throw new Error(msg);
  }
  if (!VARS[varsIdx].hasOwnProperty(varName)) {
    const msg = "Could not find the " + varName + " entry in the " + tabName + " tab";
    throw new Error(msg);
  }

  return VARS[varsIdx][varName];
}


/**
 * Function buildEmailFwdAddrs
 * 
 * Builds a forward lookup of the email forwarding addresses.
 * The key of the table entries is the plus address for the steering committee position.
 * The value of the table entries is the email address that is forwarded to.
 * 
 * @returns   {object}  Object containing the email forwarding addresses.
 */
function buildEmailFwdAddrs() {
  const filters = Gmail.Users.Settings.Filters.list(`me`).filter;
  var addrs = {};

  filters.forEach(filter => {
    addrs[filter.criteria.to] = filter.action.forward;
  });

  return addrs;
}


/**
 * Function buildEmailPlusAddrs
 * 
 * Builds a reverse lookup of the email forwarding addresses.
 * The key of the table entries is the email address that is forwarded to.
 * The value of the table entries is the plus address for the steering committee position.
 * 
 * @returns   {object}  Object containing the email forwarding addresses.
 */
function buildEmailPlusAddrs() {
  const filters = Gmail.Users.Settings.Filters.list(`me`).filter;
  var addrs = {};

  filters.forEach(filter => {
    addrs[filter.action.forward] = filter.criteria.to;
  });

  return addrs;
}



/**
 * Function buildTitles
 * 
 * Builds a reverse lookup of titles by email forward address.
 * 
 * @returns   {object}  Object containing the titles.
 */
function buildTitles() {
  const sendAs = Gmail.Users.Settings.SendAs.list(`me`).sendAs;
  var titles = {};

  sendAs.forEach(item => {
    if (EMAIL_FWD_ADDRS.hasOwnProperty(item.sendAsEmail)) {
      var title = item.displayName;

      // Trim off the prefix
      if (title.startsWith(TITLE_PREFIX)) {
        title = title.substring(TITLE_PREFIX.length, title.length).trim();
      }

      titles[EMAIL_FWD_ADDRS[item.sendAsEmail]] = title;
    }
  });

  return titles;
}


/**
 * Function checkSendPermissions
 * 
 * Checks whether the sender has permissions to send to a specified group of members.
 * 
 * @param   sender  {string}  Email address of the sender
 * @param   group   {string}  Name of the group to send the email to
 * 
 * @returns {bool}            True if the sender has permission to send emails to the specified group
 */
function checkSendPermissions(sender, group) {
  // If the sender's email isn't recognized, then they don't have permission to send
  if (!TITLES.hasOwnProperty(sender)) return false;

  // Get the permissions for the sender
  const senderPermissions = getScriptVar(TAB_NAME.PERMISSIONS, TITLES[sender]).split(`,`);

  // Check if the group is in the sender's permissions
  var foundPermission = false;
  senderPermissions.forEach(_permission => {
    const permission = _permission.toUpperCase().trim();
    if ((permission == `ALL`) || (permission == group.toUpperCase().trim())) {
      foundPermission = true;
    }
  });

  // If we're here, the permission wasn't found.
  // The sender isn't allowed to send to this group.
  return foundPermission;
}


/**
 * Function getEmailsForMailingList
 * 
 * Gets the email addresses for all contacts in the specified mailing list.
 *  
 * @param   {string}  mailingList  Name of a mailing list or an individual email address
 * @return  {string}               Comma separated email addresses of contacts in the mailing lists
 */
function getEmailsForMailingList(mailingList) {
  var emailAddresses = [];

  // Check if each item is an email address by the presence of an '@' character
  if (mailingList.includes(`@`)) {
    const emailAddr = mailingList; // mailingList is actually an email address

    // Check that there is an email contact with this mailing address
    if (!ALL_MEMBER_EMAILS.includes(emailAddr)) {
      throw `Email address ${emailAddr} does not belong to an active member`;
    }

    // Just return the email address
    emailAddresses.push(emailAddr);
    return emailAddresses;
  }

  // Find the ContactGroup whose name matches the mailingList
  var groupResourceName = ``;
  People.ContactGroups.list({pageSize: 100}).contactGroups.forEach(contactGroup => {
    if (contactGroup.name == mailingList) {
      groupResourceName = contactGroup.resourceName;
    }
  });

  // Throw an error if no ContactGroup was found for the mailingList
  if (groupResourceName == ``) {
    throw `No ContactGroup found for ${mailingList}`;
    return [];
  }

  // Get the ContactGroup
  // Done this way because memberResourceNames is only populated by get(), not list()
  const group = People.ContactGroups.get(groupResourceName, {maxMembers: MAX_MEMBERS_PER_MAILING_LIST});

  // Collect all email addresses in the contact group
  for (var batchStart = 0; batchStart < group.memberResourceNames.length; batchStart += MAX_PEOPLE_GET_BATCH_SIZE) {
    const batchResourceNames = group.memberResourceNames.slice(batchStart, batchStart + MAX_PEOPLE_GET_BATCH_SIZE);
    const batchResponse = People.People.getBatchGet({resourceNames: batchResourceNames, personFields: `emailAddresses`});
    batchResponse.responses.forEach(response => {
      response.person.emailAddresses.forEach(emailAddr => {
        if (emailAddr.value) {
          emailAddresses.push(emailAddr.value);
        }
      });
    });
  }

  return emailAddresses;
}


/**
 * Function getEmailsForMailingLists
 * 
 * Gets the email addresses for all contacts in the specified mailing list(s).
 * 
 * The mailingLists can include individual emails so long as these emails belong to existing contacts.
 * 
 * @param   {string}  mailingLists  Comma separated names of mailing lists and/or individual emails
 * @return  {string}                Comma separated email addresses of contacts in the mailing lists
 */
function getEmailsForMailingLists(mailingLists) {
  var emailAddresses = [];

  // Split the string into a list
  const mailingListItems = mailingLists.split(",");

  // Collect email addresses from each list
  mailingListItems.forEach(item_ => {
    const item = item_.trim();
    emailAddresses = emailAddresses.concat(getEmailsForMailingList(item));
  });

  return emailAddresses;
}


/**
 * Function getNextEmail2Send
 * 
 * Searches the inbox for emails with +send in the "To:" address.
 * Returns the first email that it finds that matches the criteria.
 */
function getNextEmail2Send() {
  /*
  const inboxEmails = Gmail.Users.Messages.list(`me`)[`messages`];
  Logger.log(`Listed messages: ${JSON.stringify(inboxEmails)}`);
  const msg = Gmail.Users.Messages.get(`me`, inboxEmails[0].id);
  Logger.log(`First message: ${JSON.stringify(msg)}`);
  */
  var threads, threadNum = 0, messages;

  for (let start = 0, threads = GmailApp.getInboxThreads(0, MAX_EMAIL_THREADS_TO_GET);
       (threads.length > 0);
       start += MAX_EMAIL_THREADS_TO_GET, threads = GmailApp.getInboxThreads(start, MAX_EMAIL_THREADS_TO_GET))
  {
    Logger.log(`start = ${start}, threads.length = ${threads.length}`);
    for (let threadIdx = 0; threadIdx < threads.length; threadIdx++) {
      messages = threads[threadIdx].getMessages();
      for (let msgNum = 0; msgNum < messages.length; msgNum++) {
        const msg = messages[msgNum];
        var to = msg.getTo();
        Logger.log(`Thread ${threadNum}, Message ${msgNum}:\n From: ${msg.getFrom()}\n To: ${to}\n Subject: ${msg.getSubject()}\n Body: ${msg.getBody()}`);

        to = to.split(`@`)[0]; // Strip @gmail.com
        if (to.includes(`+send-`)) {
          const groups = to.split(`+send-`)[1].split(`-`);
          Logger.log(`Groups: ${JSON.stringify(groups)}`);
          return msg;
        }

        msgNum++;
      }
      threadNum++;
    }
  }
}


/**
 * Function forwardEmails
 * 
 * Checks for emails to forward.
 * Checks that the sender of each email has permissions to send to the groups that they want to send to.
 * Creates a new email on the sender's behalf, copies the original email and sends it to the email groups.
 */
function forwardEmails() {
  var message;

  for (message = getNextEmail2Send(); message != null; message = getNextEmail2Send()) {
    const sender = message.getFrom();
    Logger.log(`Forwarding message from sender: ${sender}`);

    // Get the groups to send to
    const groups = message.getTo().split(`@`)[0].split(`+send-`)[1].split(`-`);
    Logger.log(`Forwarding to groups: ${groups}`);

    // Check that the sender has permission to send to these groups
    var cantSendGroups = [];
    for (let grpIdx = 0; grpIdx < groups.length; grpIdx++) {
      const group = groups[grpIdx];
      if (!checkSendPermissions(sender, group)) {
        cantSendGroups.push(group);
      }
    }

    // Send the email
    if (cantSendGroups.length == 0) {
      Logger.log(`TODO: Send the email`);
    } else {
      Logger.log(`${sender} doesn't have permission to send to the following groups: ${groups.join(`, `)}`);
      Logger.log(`TODO: Send rejection reply`);
    }

    // Delete the original email
    Logger.log(`TODO: Delete the original email`);
    break;
  }
}


function test() {
  Logger.log(`EMAIL_FWD_ADDRS: ${JSON.stringify(EMAIL_FWD_ADDRS)}`);
  Logger.log(`EMAIL_PLUS_ADDRS: ${JSON.stringify(EMAIL_PLUS_ADDRS)}`);
  Logger.log(`TITLES: ${JSON.stringify(TITLES)}`);
  Logger.log(`Chair permissions: ${VARS.GROUPS[VARS.PERMISSIONS[TITLES["MillieWVBChair@gmail.com"]]]}`);
  Logger.log(`Millie may send to all? ${checkSendPermissions("MillieWVBChair@gmail.com", "All")}`);
  Logger.log(`Millie may send to l6? ${checkSendPermissions("MillieWVBChair@gmail.com", "l6")}`);
  Logger.log(`Elaine may send to all? ${checkSendPermissions("elainehunt4@gmail.com", "All")}`);
  Logger.log(`Elaine may send to l1? ${checkSendPermissions("elainehunt4@gmail.com", "l1")}`);
  Logger.log(`Elaine may send to l2? ${checkSendPermissions("elainehunt4@gmail.com", "l2")}`);
  Logger.log(`Elaine may send to sc? ${checkSendPermissions("elainehunt4@gmail.com", "sc")}`);
}
