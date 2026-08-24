
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
  Logger.log(`sendAs: ${JSON.stringify(sendAs)}`);

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
