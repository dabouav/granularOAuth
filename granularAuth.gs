/**
 * @fileoverview Functions to aid App Script powered scripts, Add-ons and
 * Chat Apps to check for required OAuth scopes, and direct the end-user to 
 * re-authorize if some or all scopes listed in the manifest have not been 
 * granted.
 * 
 * (c) Dave Abouav, 2025 (https://github.com/dabouav)
 * 
 * Free to use for any purpose, with or without attribution. 
 *
 * Example use:
 * 
 * // Check if user has granted/authorized *all* scopes listed in this 
 * // project's manifest, and if not, provide them a link to grant the
 * // remaining ones. This check should be done at the start of any entry 
 * // point into your app that a brand new user may execute (i.e. Menu 
 * // items for them to start using your app). 
 * 
 * function menuItemStartAction() {
 *
 *   if (authHandleMissingScopeGrants()) {
 *     // user needs to grant access to some or all scopes before normal
 *     // app execution can continue.
 *     return;
 *   }
 *   
 *  // continue with normal execution
 *  handleStartAction();
 * }
 */

const AUTH_APP_NAME = '<CHANGE TO NAME OF YOUR APP>';

// IMPORTANT: To avoid requesting unnecessary scopes, list ONLY the 
// host applications this add-on is designed for.
// For example, for a Sheets-only add-on, use: [SpreadsheetApp]
const ADD_ON_CONTAINERS = [SpreadsheetApp, DocumentApp, SlidesApp, FormApp];

/**
 * Call this function at the start of any entry point into your app that
 * a brand new user may execute (i.e. Menu items for them to start 
 * using your app). This is really main function you'll use from this file.
 *
 * return {Boolean} True if the re-auth prompt has been shown, in which case
 *   the caller should return/abort. Returns false otherwise.
 */
function authHandleMissingScopeGrants() {
  if (!authUserHasGrantedAllScopes()) { 
    _showAllScopesRequiredMessage();
    return true;
  }

  return false;
}


/**
 * Checks if *all* scopes listed in this app's manifest have been granted by
 * the user. 
 *
 * return {Boolean}.
 */
function authUserHasGrantedAllScopes() {
  let authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
  let authStatus = authInfo.getAuthorizationStatus();

  if (authStatus === ScriptApp.AuthorizationStatus.NOT_REQUIRED) {
    return true;
  }

  // One or more OAuth scopes still need to be granted by the user
  return false;
}

/* Checks if the passed scope is missing from the user granted scopes.
 *
 * @param {String} scopeId Full path of OAuth scope being asked about.
 *   (i.e. 'https://www.googleapis.com/auth/drive.file')
 * 
 * return {Boolean}
 */
function authCheckForMissingScope(scopeId) {
  let authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
  let grantedScopes = authInfo.getAuthorizedScopes();

  if (grantedScopes.indexOf(scopeId) === -1) {
    // passed scope is not present in list or granted/authorizes scopes
    return true;
  }

  return false;
}

//// INTERNAL helper functions below this point ////

/* Checks if the script.container.ui is present among the user granted scopes, and ergo
 * this app can show the user HTML based content in a modal or the sidepanel.
 *
 * return {Boolean}
 */
function _appCanDisplayUIContent() {
  let uiScopeIsMissing = authCheckForMissingScope('https://www.googleapis.com/auth/script.container.ui');

  if (uiScopeIsMissing) {
    return false;
  }

  return true;
}

/**
 * Shows a message box asking the user to re-authorize remaining scopes 
 * needed by the app, with a link to click on.
 *
 * return {Boolean}.
 */
function _showAllScopesRequiredMessage() {
  let authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
  let reAuthUrl = authInfo.getAuthorizationUrl();

  let title = 'Further Authorization Required';

  if (_appCanDisplayUIContent()) {
    let message = `To operate properly, ${AUTH_APP_NAME} requires that you approve <u>all</u>  \
                      of its listed permissions (scopes) during installation. \
                      <p>Please click the link below to re-authorize ${AUTH_APP_NAME}, and \
                      be sure to select all of the permissions listed there.
                      Afterwards, close this window and access ${AUTH_APP_NAME} again.</p> \
                      <a target="_blank" href="${reAuthUrl}">Authorize Now</a>`;

    _showHtmlNotification(title, message);
  }
  else {
    const ui = _getActiveUi();
    let message = `${title}\\n\\nTo operate properly, ${AUTH_APP_NAME} requires that you approve all  \
                      of its listed permissions (scopes) during installation. \
                      \\n\\nPlease visit the link below in a new tab to re-authorize ${AUTH_APP_NAME}, and \
                      be sure to select all of the permissions listed there.
                      Afterwards, close this window and access ${AUTH_APP_NAME} again.\\n\\n${reAuthUrl}`;
    if (ui) {
        // Use the cross-compatible Ui.alert() for the fallback message.
        ui.alert(title, message, ui.ButtonSet.OK);
    } else {
        // This case is highly unlikely in a functioning add-on.
        console.error("Could not display the authorization message because no active UI was found.");
    }
  }
}


/**
 * Shows an HTML based message box to the user. For Editor based Add-ons only.
 * Please modify the last line of this function depending on which type of Editor
 * you are showing this message in (Spreadsheet, Doc, Form, etc).
 * 
 * If the scope for showing third-party/sidepanel content doesn't exist, then
 * falls back to native Browser.msgBox(), from which user will have to copy/paste
 * re-auth link.
 */
function _showHtmlNotification(title, message) {
  let html = _createHtmlMessageBox(title, message);

  const ui = _getActiveUi();

  if (ui) {
    ui.showModalDialog(html, title);
  } else {
    // Fallback if no UI can be determined, though this is unlikely in an add-on.
    console.error("Could not display the authorization message in a dialog.");
  }
}

/**
 * Gets the UI object for the active container based on the ADD_ON_CONTAINERS constant.
 * This developer-configured approach prevents requesting unnecessary scopes.
 * @returns {GoogleAppsScript.Base.Ui|null} The UI object for the active editor, or null if none is found.
 */
function _getActiveUi() {
  const activeApp = ADD_ON_CONTAINERS.find((app) => {
    try {
      // The getUi() method will only succeed in the active container.
      // This check confirms the app object exists and has the getUi method.
      return app && typeof app.getUi === 'function' && app.getUi();
    } catch (f) {
      // An error will be thrown if trying to access an inactive service UI.
      return false;
    }
  });

  if (activeApp) {
    return activeApp.getUi();
  }
  return null;
};

/**
 * Creates the HTML output for the message box.
 * @param {string} title The title for the HTML document.
 * @param {string} message The HTML message content.
 * @returns {GoogleAppsScript.HTML.HtmlOutput} The configured HtmlOutput object.
 */
function _createHtmlMessageBox(title, message) {

  let html = HtmlService.createHtmlOutput()
                        .setSandboxMode(HtmlService.SandboxMode.IFRAME)
                        .setWidth(500).setHeight(250);

  html.setTitle(title);
  
  let h = `<html> \
           <body style="margin-top:0px;"> \
           <div style="font-face:Arial,Times;font-family:Sans-Serif;font-size:18px;\
             padding:30px 0px 30px 0px;vertical-align:top;">      
           ${message}</div> \
           </body></html>`;
  
  html.append(h);
  
  return html;
}
