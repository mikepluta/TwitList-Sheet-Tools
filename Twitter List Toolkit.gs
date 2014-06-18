/**
 * Twitter List Management Toolkit for Google Sheets
 * 
 * Mike Pluta 06.18.2014
 * @mikepluta
 * mike.pluta@gmail.com
 *
 **
 *
 * Prior to use, you will need to create a Twitter API Key and Secret.  To do this, go to
 * https://apps.twitter.com/ and login with your twitter credentials.
 * On this page, press the "Create New App" button
 * Fill in the Application Name, Description and Website.  In the Callback URL field,
 * enter "https://spreadsheets.google.com/". Read the 'Rules of the Road', select the
 * checkbox indicating that you agree and press the 'Create Your Twitter Application Button'
 *
 * on the resulting pages, go to the API Keys Tab and select the 'Generate API Key' Button
 * make note of the resulting API Key and API Secret (NOT the Access Token or Access Token
 * Secret)
 *
 * From the 'Twitter' menu on the Google Sheet, select the 'Configure' option and fill in
 * the API Key, API Secret and your user screen name and press the 'Save Configuration'
 * button.  Verify that the authentication credentials are correct by selecting the
 * 'Authorize' item from the 'Twitter' menu.
 *
 **/

var API_KEY_PROPERTY_NAME          = "twitterAPIKey";
var API_SECRET_PROPERTY_NAME       = "twitterAPISecret";
var AUTH_SCREEN_NAME_PROPERTY_NAME = "twitterScreenName";
var AUTH_USER_ID_PROPERTY_NAME     = "twitterUserId";

/**
 * Return String OAuth API key to use when tweeting.
 */
function getAPIKey() {
  var userProperties = PropertiesService.getUserProperties();
  var key = userProperties.getProperty(API_KEY_PROPERTY_NAME);
  if (key == null) {
    key = "";
  }
  return key;
}

/**
 * Set String OAuth API key to use when tweeting.
 */
function setAPIKey(key) {
  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty(API_KEY_PROPERTY_NAME, key);
}

/**
 * Return String OAuth API secret to use when tweeting.
 */
function getAPISecret() {
  var userProperties = PropertiesService.getUserProperties();
  var secret = userProperties.getProperty(API_SECRET_PROPERTY_NAME);
  if (secret == null) {
    secret = "";
  }
  return secret;
}

/**
 * Set String OAuth API secret to use when tweeting.
 */
function setAPISecret(secret) {
  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty(API_SECRET_PROPERTY_NAME, secret);
}

/**
 * Return String Authorized User Screen Name
 */
function getScreenName() {
  var userProperties = PropertiesService.getUserProperties();
  var screen_name = userProperties.getProperty(AUTH_SCREEN_NAME_PROPERTY_NAME);
  if (screen_name == null) {
    screen_name = "";
  }
  return screen_name;
}

/**
 * Set String Authorized User Screen Name
 */
function setScreenName(screen_name) {
  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty(AUTH_SCREEN_NAME_PROPERTY_NAME, screen_name);
}

/**
 * Return String Authorized User ID
 */
function getUserId() {
  var userProperties = PropertiesService.getUserProperties();
  var user_id = userProperties.getProperty(AUTH_USER_ID_PROPERTY_NAME);
  if (user_id == null) {
    user_id = "";
  }
  return user_id;
}

/**
 * Set String Authorized User ID
 */
function setUserId(user_id) {
  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty(AUTH_USER_ID_PROPERTY_NAME, user_id);
}

/**
 * Return bool True if all of the configuration properties are set,
 *              false if otherwise.
 */
function isConfigured() {
  return getAPIKey() != "" && getAPISecret() != "" && getScreenName() != "";
}

/** Retrieve config params from the UI and store them. */
function saveConfiguration(e) {
  setAPIKey(e.parameter.APIKey);
  setAPISecret(e.parameter.APISecret);
  setScreenName(e.parameter.screenName);
  var app = UiApp.getActiveApplication();
  app.close();
  return app;
}

/**
 * Configure all UI components and display a dialog to allow the user to 
 * configure OAuth API Key and API Secret.
 */
function renderConfigurationDialog() {
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  var app = UiApp.createApplication().setTitle(
      "Configure the OAuth Credentials");
  app.setStyleAttribute("padding", "10px");
  
  var helpLabel = app.createLabel(
      "From here, you can configure the OAuth credentials.  You will need the " +
      "API Key and the API Secret, not the actual tokens.");
  helpLabel.setStyleAttribute("text-align", "justify");

  var APIKeyLabel = app.createLabel("Twitter OAuth API Key:");
  var APIKey      = app.createTextBox();
  APIKey.setName("APIKey");
  APIKey.setWidth("90%");
  APIKey.setText(getAPIKey());

  var APISecretLabel = app.createLabel("Twitter OAuth API Secret:");
  var APISecret      = app.createTextBox();
  APISecret.setName("APISecret");
  APISecret.setWidth("90%");
  APISecret.setText(getAPISecret());
  
  var screenNameLabel = app.createLabel("Authorized User Screen Name:");
  var screenName      = app.createTextBox();
  screenName.setName("screenName");
  screenName.setWidth("90%");
  screenName.setText(getScreenName());
  
  var userIDLabel = app.createLabel("Authorized User ID:");
  var userID      = app.createLabel(getUserId());
  userID.setWidth("90%");
  
  var saveHandler = app.createServerClickHandler("saveConfiguration");
  var saveButton = app.createButton("Save Configuration", saveHandler);
  
  var listPanel = app.createGrid(4, 2);
  listPanel.setStyleAttribute("margin-top", "10px")
  listPanel.setWidth("100%");
  listPanel.setWidget(0, 0, APIKeyLabel);
  listPanel.setWidget(0, 1, APIKey);
  listPanel.setWidget(1, 0, APISecretLabel);
  listPanel.setWidget(1, 1, APISecret);
  listPanel.setWidget(2, 0, screenNameLabel);
  listPanel.setWidget(2, 1, screenName);
  listPanel.setWidget(3, 0, userIDLabel);
  listPanel.setWidget(3, 1, userID);

  // Ensure that all form fields get sent along to the handler
  saveHandler.addCallbackElement(listPanel);
  
  var dialogPanel = app.createFlowPanel();
  dialogPanel.add(helpLabel);
  dialogPanel.add(listPanel);
  dialogPanel.add(saveButton);
  app.add(dialogPanel);
  doc.show(app);
}

/** Controller to render approvers UI and apply configuration. */
function configure() {
  renderConfigurationDialog();
}
  
/**
 * Authorize against Twitter.  This method must be run prior to 
 * clicking any link in a script email.  If you click a link in an
 * email, you will get a message stating:
 * "Authorization is required to perform that action."
 */
function authorize() {
  var URI   = "https://api.twitter.com/1.1/account/verify_credentials.json";
  var query = "?skip_status=true";
  
  var oauthConfig = UrlFetchApp.addOAuthService("twitter");
  oauthConfig.setAccessTokenUrl("https://api.twitter.com/oauth/access_token");
  oauthConfig.setRequestTokenUrl("https://api.twitter.com/oauth/request_token");
  oauthConfig.setAuthorizationUrl("https://api.twitter.com/oauth/authorize");
  oauthConfig.setConsumerKey(getAPIKey());
  oauthConfig.setConsumerSecret(getAPISecret());

  var requestData = {"method"             : "GET",
                     "MuteHttpExceptions" : "true",
                     "oAuthServiceName"   : "twitter",
                     "oAuthUseToken"      : "always"
                    };

  try {
    var result = UrlFetchApp.fetch(URI+query, requestData);
    var o      = JSON.parse(result.getContentText());
    setUserId(o.id_str);
  } catch(e) {
    Logger.log(e);
  }

}

function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [{name : "Export Follower IDs",          functionName : "exportFollowerIDs"},
                     {name : "Export Friend IDs",            functionName : "exportFriendIDs"},
//                     {name : "Hydrate IDs",                  functionName : "hydrateIDs"},
                     {name : "Export Users from a List",     functionName : "exportUsersFromList"},
                     {name : "Add Users to a List",          functionName : "addToList"},
                     {name : "Delete Users from a List",     functionName : "deleteFromList"},
                     {name : "Download Tweets from a List",  functionName : "downloadTweetsFromList"},
                     {name : "Download Tweets from a Query", functionName : "downloadTweetsFromQuery"},
//                     {name : "Tweet",                        functionName : "sendTweet"},
                     {name : "Get List Names & IDs",         functionName : "getLists"},
                     {name : "Configure",                    functionName : "configure"},
                     {name : "Authorize",                    functionName : "authorize"}];
  ss.addMenu("Twitter", menuEntries);
}

function testit() {
  var ids = [1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26];
  for (var i in ids) {
    Logger.log("index = " + i);
    Logger.log("value = " + ids[i]);
  }
}

function exportFollowerIDs() {
  // Followers are defined to be those who follow me
  var dataSheet  = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var newMember  = "";
  var screenName = getScreenName();
  var batchSize  = 1000;
  var cursor     = "-1";
  var rowNum     = 1;
  var colNum     = 1;
  
  //get list name to add users to (must already exist)
  //var list_id = Browser.inputBox("Enter list name from which users are to be exported:");
  
  var URI = "https://api.twitter.com/1.1/followers/ids.json"

  authorize();

  var requestData = {"method"             : "GET",
                     "MuteHttpExceptions" : "true",
                     "oAuthServiceName"   : "twitter",
                     "oAuthUseToken"      : "always"   };

  var query = "?screen_name=" + screenName +
              "&stringify_ids=true" +
              "&count=" + batchSize +
              "&cursor=" + cursor;

  do {
    var result = UrlFetchApp.fetch(URI+query, requestData);
    var o = JSON.parse(result.getContentText());
    Logger.log(o);
    for (fID in o.ids) {
      dataSheet.getRange(++rowNum,colNum).setValue(o.ids[fID]);
    }
//    for (var i = 0; i < o.users.length; ++i) {
//      dataSheet.getRange(++rowNum, colNum).setValue(o.users[i].screen_name);
//      dataSheet.getRange(rowNum, colNum+1).setValue(JSON.stringify(o.users[i]));
//    }
    cursor = o.next_cursor_str;
    var query = "?screen_name=" + screenName +
                "&stringify_ids=true" +
                "&count=" + batchSize +
                "&cursor=" + cursor;
  } while (cursor!="0");
}

function exportFriendIDs() {
  // Friends are defined as those who I follow
  var dataSheet  = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var newMember  = "";
  var screenName = getScreenName();
  var batchSize  = 1000;
  var cursor     = "-1";
  var rowNum     = 1;
  var colNum     = 1;
  
  //get list name to add users to (must already exist)
  //var list_id = Browser.inputBox("Enter list name from which users are to be exported:");
  
  var URI = "https://api.twitter.com/1.1/friends/ids.json"

  authorize();

  var requestData = {"method"             : "GET",
                     "MuteHttpExceptions" : "true",
                     "oAuthServiceName"   : "twitter",
                     "oAuthUseToken"      : "always"   };

  var query = "?screen_name=" + screenName +
              "&stringify_ids=true" +
              "&count=" + batchSize +
              "&cursor=" + cursor;

  do {
    var result = UrlFetchApp.fetch(URI+query, requestData);
    var o = JSON.parse(result.getContentText());
    Logger.log(o);
    for (fID in o.ids) {
      dataSheet.getRange(++rowNum,colNum).setValue(o.ids[fID]);
    }
//    for (var i = 0; i < o.users.length; ++i) {
//      dataSheet.getRange(++rowNum, colNum).setValue(o.users[i].screen_name);
//      dataSheet.getRange(rowNum, colNum+1).setValue(JSON.stringify(o.users[i]));
//    }
    cursor = o.next_cursor_str;
    var query = "?screen_name=" + screenName +
                "&stringify_ids=true" +
                "&count=" + batchSize +
                "&cursor=" + cursor;
  } while (cursor!="0");
}

function hydrateIDs() {
  // This needs a LOT more work
  var dataSheet  = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var dataRange  = dataSheet.getRange(2, 1, dataSheet.getMaxRows(), 1);
  var newMember  = "";
  var screenName = getScreenName();
  var batchSize  = 25;
  
  authorize();

  var requestData = {"method"             : "POST",
                     "MuteHttpExceptions" : "true",
                     "oAuthServiceName"   : "twitter",
                     "oAuthUseToken"      : "always"   };
  var URI = "https://api.twitter.com/1.1/users/lookup.json"

  objects = getRowsData(dataSheet,dataRange,1);
  var batchCount = Math.ceil(objects.length / batchSize);
  for (var b = 0; b < batchCount; ++b) {
    var query = "?include_entities=false" +
                "&user_id=";
    for (var i = b * batchSize; (i < objects.length) && (i < ((b+1)*batchSize)); ++i) {
      var rowData = objects[i];
      if (rowData.columnToLoad != "" ) {
        newMember = rowData.columnToLoad;
        query = query + newMember + ",";
        dataSheet.getRange(i+2, 2).setValue("Added");
      }
    }
    query.slice(0,-1);
    var result = UrlFetchApp.fetch(URI+query, requestData);
    for (var i = b * batchSize; (i < objects.length) && (i < ((b+1)*batchSize)); ++i) {
      if (rowData.columnToLoad != "" ) {
        newMember = rowData.columnToLoad;
        query = query + newMember + ",";
        dataSheet.getRange(i+2, 2).setValue("Added");
      }
    }
  }
  
}

function exportUsersFromList() {
  var dataSheet  = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var newMember  = "";
  var screenName = getScreenName();
  var batchSize  = 25;
  var cursor     = "-1";
  var rowNum     = 1;
  var colNum     = 1;
  
  //get list name to add users to (must already exist)
  var list_id = Browser.inputBox("Enter list name from which users are to be exported:");
  
  var URI = "https://api.twitter.com/1.1/lists/members.json"

  authorize();

  var requestData = {"method"             : "GET",
                     "MuteHttpExceptions" : "true",
                     "oAuthServiceName"   : "twitter",
                     "oAuthUseToken"      : "always"   };

  var query = "?slug=" + list_id +
              "&owner_screen_name=" + screenName +
              "&skip_status=true" +
              "&include_entities=false" +
              "&cursor=" + cursor;

  do {
    var result = UrlFetchApp.fetch(URI+query, requestData);
    var o = JSON.parse(result.getContentText());
    for (var i = 0; i < o.users.length; ++i) {
      dataSheet.getRange(++rowNum, colNum).setValue(o.users[i].screen_name);
      dataSheet.getRange(rowNum, colNum+1).setValue(JSON.stringify(o.users[i]));
    }
    cursor = o.next_cursor_str;
    var query = "?slug="                  + list_id +
                "&owner_screen_name="     + screenName +
                "&skip_status=true"       +
                "&include_entities=false" +
                "&cursor="                + cursor;
  } while (cursor!="0");
}

function addToList() {
  var dataSheet  = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var dataRange  = dataSheet.getRange(2, 1, dataSheet.getMaxRows(), 1);
  var newMember  = "";
  var screenName = getScreenName();
  var batchSize  = 25;
  
  //get list name to add users to (must already exist)
  var list_id = Browser.inputBox("Enter list name to which users are to be added:");
  
  var URI = "https://api.twitter.com/1.1/lists/members/create_all.json"

  authorize();

  var requestData = {"method"             : "POST",
                     "MuteHttpExceptions" : "true",
                     "oAuthServiceName"   : "twitter",
                     "oAuthUseToken"      : "always"   };
  objects = getRowsData(dataSheet,dataRange,1);
  var batchCount = Math.ceil(objects.length / batchSize);
  for (var b = 0; b < batchCount; ++b) {
    var query = "?slug=" + list_id +
                "&owner_screen_name=" + screenName +
                "&screen_name=";
    for (var i = b * batchSize; (i < objects.length) && (i < ((b+1)*batchSize)); ++i) {
      var rowData = objects[i];
      if (rowData.columnToLoad != "" ) {
        if (rowData.columnToLoad[0] == "@") {
          newMember = rowData.columnToLoad.substring(1);
        } else {
          newMember = rowData.columnToLoad;
        }
        query += newMember;
        query += ",";
        dataSheet.getRange(i+2, 2).setValue("Added");
      }
    }
    query.slice(0,-1);
    var result = UrlFetchApp.fetch(URI+query, requestData);
  }
}

function deleteFromList() {
  var dataSheet  = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var dataRange  = dataSheet.getRange(2, 1, dataSheet.getMaxRows(), 1);
  var newMember  = "";
  var screenName = getScreenName();
  var batchSize  = 25;
  
  //get list name to add users to (must already exist)
  var list_id = Browser.inputBox("Enter list name from which users are to be deleted:");
  
  var URI = "https://api.twitter.com/1.1/lists/members/destroy_all.json"

  authorize();

  var requestData = {"method"             : "POST",
                     "MuteHttpExceptions" : "true",
                     "oAuthServiceName"   : "twitter",
                     "oAuthUseToken"      : "always"   };

  objects = getRowsData(dataSheet, dataRange, 1);
  var batchCount = Math.ceil(objects.length / batchSize);
  for (var b = 0; b < batchCount; ++b) {
    var query = "?slug="              + list_id +
                "&owner_screen_name=" + screenName +
                "&screen_name=";
    for (var i = b * batchSize; (i < objects.length) && (i < ((b+1)*batchSize)); ++i) {
      var rowData = objects[i];
      if (rowData.columnToDelete != "" ) {
        if (rowData.columnToDelete[0] == "@") {
          newMember = rowData.columnToDelete.substring(1);
        } else {
          newMember = rowData.columnToDelete;
        }
        query += newMember;
        query += ",";
        dataSheet.getRange(i+2, 2).setValue("Deleted");
      }
    }
    query.slice(0,-1);
    var result = UrlFetchApp.fetch(URI+query, requestData);
  }
}

function downloadTweetsFromList() {
  var dataSheet   = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var newMember   = "";
  var screenName  = getScreenName();
  var batchSize   = "50";
  var cursor      = "-1";
  var startRowNum = 6;
  var rowNum      = startRowNum;
  var colNum      = 1;
  var maxTweets   = 2500;

  var max_id      = dataSheet.getRange(2, 2).getValue();
  if (max_id != "") {
    rowNum = dataSheet.getRange(1, 2).getValue();
  }
  
  var list_id = Browser.inputBox("Enter list name from which tweets are to be downloaded:");
//  var list_id = "big-data-influencers";
  
  var URI         = "https://api.twitter.com/1.1/lists/statuses.json"
  var expandURI   = "http://api.longurl.org/v2/expand";
  
  authorize();
  
  var requestData = {"method"             : "GET",
                     "MuteHttpExceptions" : "true",
                     "oAuthServiceName"   : "twitter",
                     "oAuthUseToken"      : "always"   };

  var query = "?slug=" + list_id +
              "&owner_screen_name=" + screenName +
              "&include_rts=false" +
              "&include_entities=true" +
              "&count=" + batchSize;

  if (max_id != "") {
    query += "&max_id=" + max_id;
  }

  var result = UrlFetchApp.fetch(URI+query, requestData);
  var o = JSON.parse(result.getContentText());
  while (o.length > 0 && rowNum < maxTweets+2) {
    var testVal = ((o.length > 0) && (rowNum < maxTweets+2));
    for (var i = 0; i < o.length; ++i) {
      var i_str   = Math.ceil(i);
      var status   = o[i_str];
      var entities = status.entities;
      var max_id   = status.id_str;
      if (entities.urls.length > 0) {
        for (var u = 0; u < entities.urls.length; ++u) {
          var u_str = Math.ceil(u);
          var eurl  = entities.urls[u_str].expanded_url;
          var surl  = entities.urls[u_str].url;
          rowNum = rowNum + 1;
          dataSheet.getRange(rowNum, 1).setValue(status.id_str);
          dataSheet.getRange(rowNum, 2).setValue(status.user.screen_name);
          dataSheet.getRange(rowNum, 3).setValue(status.text);
          dataSheet.getRange(rowNum, 4).setValue(surl);
          var expandQuery = "?format=json" +
                            "&title=1" +
                            "&url=" + encodeURIComponent(surl);
          var expandData  = {"method"             : "GET",
                             "MuteHttpExceptions" : "true",
                             "User-Agent"         : "BTC Personal Tweep Manager/0.1"};
          try {
            var expandResult = UrlFetchApp.fetch(expandURI + expandQuery, expandData);
            if (expandResult.getResponseCode() === 200) {
              var oResult      = JSON.parse(expandResult.getContentText());
              Logger.log(oResult);
              dataSheet.getRange(rowNum, 5).setValue(oResult["long-url"]);
              dataSheet.getRange(rowNum, 6).setValue(oResult["title"]);
            } else {
              dataSheet.getRange(rowNum, 5),setValue(eurl);
            }
            dataSheet.getRange(1, 2).setValue(rowNum);
            dataSheet.getRange(2, 2).setValue(max_id);
          } catch(e) {
            Logger.log(e);
          }
        } // else if there's no URL, I dont care...
      }
    }
    var query = "?slug=" + list_id +
                "&owner_screen_name=" + screenName +
                "&include_rts=false" +
                "&include_entities=true" +
                "&count=" + batchSize +
                "&max_id=" + max_id;
    result = UrlFetchApp.fetch(URI+query, requestData);
    o = JSON.parse(result.getContentText());
  }
}

function downloadTweetsFromQuery() {
  var dataSheet   = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var newMember   = "";
  var screenName  = getScreenName();
  var batchSize   = "100";
  var cursor      = "-1";
  var startRowNum = 6;
  var rowNum      = startRowNum;
  var colNum      = 1;
  var maxTweets   = 500;

  var max_id      = dataSheet.getRange(2, 2).getValue();
  if (max_id != "") {
    rowNum = dataSheet.getRange(1, 2).getValue();
  }
  
  var tQuery = Browser.inputBox("Enter query for tweets that are to be downloaded:");
//  var tQuery = "from:@dbtrends #DBTA100";
  
  var URI         = "https://api.twitter.com/1.1/search/tweets.json"
  
  authorize();
  
  var requestData = {"method"             : "GET",
                     "MuteHttpExceptions" : "true",
                     "oAuthServiceName"   : "twitter",
                     "oAuthUseToken"      : "always"   };

  var query = "?result_type=recent" +
              "&include_entities=true" +
              "&count=" + batchSize +
              "&q=" + encodeURIComponent(tQuery);

  if (max_id != "") {
    query += "&max_id=" + max_id;
  }

  var result = UrlFetchApp.fetch(URI+query, requestData);
  var o = JSON.parse(result.getContentText());
  while (o.search_metadata.count > 0 && rowNum < maxTweets+2) {
    var testVal = ((o.search_metadata.count > 0) && (rowNum < maxTweets+2));
    for (var i = 0; i < o.search_metadata.count; ++i) {
      var i_str   = Math.ceil(i);
      var status   = o.statuses[i_str];
      var max_id   = o.search_metadata.max_id_str;
      rowNum = rowNum + 1;
      dataSheet.getRange(rowNum, 1).setValue(status.id_str);
      dataSheet.getRange(rowNum, 2).setValue(status.user.screen_name);
      dataSheet.getRange(rowNum, 3).setValue(status.text);
    } // else if there's no URL, I dont care...
    dataSheet.getRange(1, 2).setValue(rowNum);
    dataSheet.getRange(2, 2).setValue(max_id);
//    var query = "?result_type=recent" +
//                "&include_entities=true" +
//                "&count=" + batchSize +
//                "&q=" + encodeURIComponent(tQuery) +
//                "&max_id=" + max_id;
    query  = o.search_metadata.next_results;
    Logger.log(query);
    result = UrlFetchApp.fetch(URI+query, requestData);
    o = JSON.parse(result.getContentText());
  }
}

function getLists() {
  var dataSheet   = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var screenName  = getScreenName();
  var batchSize   = "20";
  var cursor      = "-1";
  var startRowNum = 2;
  var rowNum      = startRowNum;
  var colNum      = 1;

  var URI = "https://api.twitter.com/1.1/lists/ownerships.json"
  var query = "?screen_name=" + screenName +
              "&count="       + batchSize +
              "&cursor="      + cursor;

  authorize();
  
  var requestData = {"method"             : "GET",
                     "MuteHttpExceptions" : "true",
                     "oAuthServiceName"   : "twitter",
                     "oAuthUseToken"      : "always"   };
  do {
    var result = UrlFetchApp.fetch(URI+query, requestData);
    var o = JSON.parse(result.getContentText());
    for (var i = 0; i < o.lists.length; ++i) {
      dataSheet.getRange(rowNum, colNum).setValue(o.lists[i].slug);
      dataSheet.getRange(rowNum++, colNum+1).setValue(o.lists[i].id_str);
    }
    cursor = o.next_cursor_str;
    var query = "?screen_name=" + screenName +
                "&count="       + batchSize
                "&cursor="      + cursor;
  } while (cursor!="0");
}
  
//////////////////////////////////////////////////////////////////////////////////////////
//
// The code below is reused from the 'Reading Spreadsheet data using JavaScript Objects'
// tutorial.
//
//////////////////////////////////////////////////////////////////////////////////////////

// getRowsData iterates row by row in the input range and returns an array of objects.
// Each object contains all the data for a given row, indexed by its normalized column name.
// Arguments:
//   - sheet: the sheet object that contains the data to be processed
//   - range: the exact range of cells where the data is stored
//   - columnHeadersRowIndex: specifies the row number where the column names are stored.
//       This argument is optional and it defaults to the row immediately above range;
// Returns an Array of objects.
function getRowsData(sheet, range, columnHeadersRowIndex) {
  columnHeadersRowIndex = columnHeadersRowIndex || range.getRowIndex() - 1;
  var numColumns = range.getEndColumn() - range.getColumn() + 1;
  var headersRange = sheet.getRange(columnHeadersRowIndex, range.getColumn(), 1, numColumns);
  var headers = headersRange.getValues()[0];
  return getObjects(range.getValues(), normalizeHeaders(headers));
}

// For every row of data in data, generates an object that contains the data. Names of
// object fields are defined in keys.
// Arguments:
//   - data: JavaScript 2d array
//   - keys: Array of Strings that define the property names for the objects to create
function getObjects(data, keys) {
  var objects = [];
  for (var i = 0; i < data.length; ++i) {
    var object = {};
    var hasData = false;
    for (var j = 0; j < data[i].length; ++j) {
      var cellData = data[i][j];
      if (isCellEmpty(cellData)) {
        continue;
      }
      object[keys[j]] = cellData;
      hasData = true;
    }
    if (hasData) {
      objects.push(object);
    }
  }
  return objects;
}

// Returns an Array of normalized Strings.
// Arguments:
//   - headers: Array of Strings to normalize
function normalizeHeaders(headers) {
  var keys = [];
  for (var i = 0; i < headers.length; ++i) {
    var key = normalizeHeader(headers[i]);
    if (key.length > 0) {
      keys.push(key);
    }
  }
  return keys;
}

// Normalizes a string, by removing all alphanumeric characters and using mixed case
// to separate words. The output will always start with a lower case letter.
// This function is designed to produce JavaScript object property names.
// Arguments:
//   - header: string to normalize
// Examples:
//   "First Name" -> "firstName"
//   "Market Cap (millions) -> "marketCapMillions
//   "1 number at the beginning is ignored" -> "numberAtTheBeginningIsIgnored"
function normalizeHeader(header) {
  var key = "";
  var upperCase = false;
  for (var i = 0; i < header.length; ++i) {
    var letter = header[i];
    if (letter == " " && key.length > 0) {
      upperCase = true;
      continue;
    }
    if (!isAlnum(letter)) {
      continue;
    }
    if (key.length == 0 && isDigit(letter)) {
      continue; // first character must be a letter
    }
    if (upperCase) {
      upperCase = false;
      key += letter.toUpperCase();
    } else {
      key += letter.toLowerCase();
    }
  }
  return key;
}

// Returns true if the cell where cellData was read from is empty.
// Arguments:
//   - cellData: string
function isCellEmpty(cellData) {
  return typeof(cellData) == "string" && cellData == "";
}

// Returns true if the character char is alphabetical, false otherwise.
function isAlnum(char) {
  return char >= 'A' && char <= 'Z' ||
    char >= 'a' && char <= 'z' ||
    isDigit(char);
}

// Returns true if the character char is a digit, false otherwise.
function isDigit(char) {
  return char >= '0' && char <= '9';
}

function sendTweet() {
  authorize();
  // URI Encode the Tweet
  var encodedTweet = encodeString(getRowContent());
  var requestData = {
    "method" : "POST",
    "oAuthServiceName" : "twitter",
    "oAuthUseToken" : "always"
  };
  try {
    var result = URLFetchApp.fetch(
      "https://api.twitter.com/1.1/statuses/update.json?status=" + encodedTweet,
      requestData);
  } catch(e) {
    Logger.log(e);
  }
}

function getRowID() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets()[0].getRange(2, 2).getValue();
}

function getRowContent() {
  var content = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0].getRange(getRowID(), 1).getValue();

  increment();

  return content;
}

function increment() {
  var rowVal = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0].getRange(2, 2).getValue() + 1;
  SpreadsheetApp.getActiveSpreadsheet().getSheets()[0].getRange(2, 2).setValue(rowVal);
}

// Thank you +Martin Hawksey - you are awesome

function encodeString(q) {
  
  // Update: 09/06/2013
  
  // Google Apps Script is having issues storing oAuth tokens with the Twitter API 1.1 due to some encoding issues.
  // Henc this workaround to remove all the problematic characters from the status message.
  
  var str = q.replace(/\(/g,'{').replace(/\)/g,'}').replace(/\[/g,'{').replace(/\]/g,'}').replace(/\!/g, '|').replace(/\*/g, 'x').replace(/\'/g, '');
  return encodeURIComponent(str);
}
