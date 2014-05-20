/**
* @author       Craig Thomas
* @date         24/04/2014
* @description  Google Spreadsheets Script to query Salesforce API data
* @URL          https://github.com/tinybigideas/Google-Cloud-Connecter
*/


/**
 * Key of ScriptProperties for Salesforce Username.
 * @type {String}
 * @const
 */
var USERNAME_PROPERTY_NAME = "username";

/**
 * Key of ScriptProperties for Salesforce Password.
 * @type {String}
 * @const
 */
var PASSWORD_PROPERTY_NAME = "password";

/**
 * Key of ScriptProperties for Salesforce Security Token.
 * @type {String}
 * @const
 */
var SECURITY_TOKEN_PROPERTY_NAME = "securityToken";

/**
 * Key of ScriptProperties for Salesforce Session Id.
 * @type {String}
 * @const
 */
var SESSION_ID_PROPERTY_NAME = "sessionId";

/**
 * Key of ScriptProperties for serviceUrl.
 * @type {String}
 * @const
 */
var SERVICE_URL_PROPERTY_NAME = "serviceUrl";


/**
 * Key of ScriptProperties for instance url.
 * @type {String}
 * @const
 */
var INSTANCE_URL_PROPERTY_NAME = "instanceUrl";

/**
 * Key of ScriptProperties for sandbox url.
 * @type {String}
 * @const
 */
var IS_SANDBOX_PROPERTY_NAME = "isSandbox";

/**
 * Key of ScriptProperties for next records url.
 * @type {String}
 * @const
 */
var NEXT_RECORDS_URL_PROPERTY_NAME = "nextRecordsUrl";

var SOBJECT_ATTRIBUTES_PROPERTY_NAME = "sObjectAttributes";

var SANDBOX_SOAP_URL = "https://test.salesforce.com/services/Soap/u/30.0";

var PRODUCTION_SOAP_URL = "https://login.salesforce.com/services/Soap/u/30.0";

/**
 * @return String Username.
 */
function getUsername() {
    var key = ScriptProperties.getProperty(USERNAME_PROPERTY_NAME);
    if (key == null) {
        key = "";
    }
    return key;
};
 
/**
 * @param String Username.
 */
function setUsername(key) {
    ScriptProperties.setProperty(USERNAME_PROPERTY_NAME, key);
};
 
/**
 * @return String Password.
 */
function getPassword() {
    var key = ScriptProperties.getProperty(PASSWORD_PROPERTY_NAME);
    if (key == null) {
        key = "";
    }
    return key;
};
 
/**
 * @param String Password.
 */
function setPassword(key) {
    ScriptProperties.setProperty(PASSWORD_PROPERTY_NAME, key);
};

/**
 * @return String Security Token.
 */
function getSecurityToken() {
    var key = ScriptProperties.getProperty(SECURITY_TOKEN_PROPERTY_NAME);
    if (key == null) {
        key = "";
    }
    return key;
};
 
/**
 * @param String Security Token.
 */
function setSecurityToken(key) {
    ScriptProperties.setProperty(SECURITY_TOKEN_PROPERTY_NAME, key);
}

/**
 * @return String Session Id.
 */
function getSessionId() {
    var key = ScriptProperties.getProperty(SESSION_ID_PROPERTY_NAME);
    if (key == null) {
        key = "";
    }
    return key;
};

/**
 * @param String Session Id.
 */
function setSessionId(key) {
    ScriptProperties.setProperty(SESSION_ID_PROPERTY_NAME, key);
};

/**
 * @return String Instance URL.
 */
function getInstanceUrl() {
    var key = ScriptProperties.getProperty(INSTANCE_URL_PROPERTY_NAME);
    if (key == null) {
        key = "";
    }
    return key;
};

/**
 * @param String Instance URL.
 */
function setInstanceUrl(key) {
    ScriptProperties.setProperty(INSTANCE_URL_PROPERTY_NAME, key);
};

/**
 * @param String use sandbox url.
 */
function setUseSandbox(key) {
    ScriptProperties.setProperty(IS_SANDBOX_PROPERTY_NAME, key);
};

/**
 * @return bool if using sandbox.
 */
function getUseSandbox() {
    var key = ScriptProperties.getProperty(IS_SANDBOX_PROPERTY_NAME);
    if (key == null) {
        key = false;
    }
    return key;
};


/**
 * @param String url for next records url.
 */
function setNextRecordsUrl(key) {
    if(key == undefined) {
        key = "";
    }
    ScriptProperties.setProperty(NEXT_RECORDS_URL_PROPERTY_NAME, key);
};

/**
 * @return String url for next records url (querymore).
 */
function getNextRecordsUrl() {
    var key = ScriptProperties.getProperty(NEXT_RECORDS_URL_PROPERTY_NAME);
    if (key == null || key == undefined) {
        key = "";
    }
    return key;
};

/**
 * @param String Instance URL.
 */
function setInstanceUrl(key) {
    ScriptProperties.setProperty(INSTANCE_URL_PROPERTY_NAME, key);
};

/**
 * @return bool if using sandbox.
 */
function getSfdcSoapEndpoint() {
    var isSandbox = getUseSandbox() == "true" ? true: false;
    if (isSandbox) {
        return SANDBOX_SOAP_URL;
    }
    else {
        return PRODUCTION_SOAP_URL;
    }
};

function getRestEndpoint() {
    var queryEndpoint = ".salesforce.com";
    var endpoint = getInstanceUrl().replace("api-","").match("https://[a-z0-9]*");
    return endpoint + queryEndpoint;
};

function onInstall() {
    onOpen();
};

function onOpen() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var menuEntries = [
        {name: "Settings", functionName: "renderSettingsDialog"},
        {name: "Login with Salesforce", functionName: "login"}
    ];
    ss.addMenu("Salesforce Connector", menuEntries);
};

/** Retrieve config params from the UI and store them. */
function saveConfiguration(e) {
    setUsername(e.parameter.username);
    setPassword(e.parameter.password);
    setSecurityToken(e.parameter.securityToken);
    setUseSandbox(e.parameter.sandbox);
    var app = UiApp.getActiveApplication();
    app.close();
    return app;
};


/**
 * Settings Dialog
 */
function renderSettingsDialog() {
    var doc = SpreadsheetApp.getActiveSpreadsheet();
    var app = UiApp.createApplication().setTitle("Salesforce Configuration");
    app.setStyleAttribute("padding", "10px");
  
    var helpLabel = app.createLabel("Enter your Username, Password, and Security Token");
    helpLabel.setStyleAttribute("text-align", "justify");
 
    var usernameLabel = app.createLabel("Username:");
    var username = app.createTextBox();
    username.setName("username");
    username.setWidth("75%");
    username.setText(getUsername());
  
    var passwordLabel = app.createLabel("Password:");
    var password = app.createPasswordTextBox();
    password.setName("password");
    password.setWidth("75%");
    password.setText(getPassword());
  
    var securityTokenLabel = app.createLabel("Security Token:");
    var securityToken = app.createTextBox();
    securityToken.setName("securityToken");
    securityToken.setWidth("75%");
    securityToken.setText(getSecurityToken());
  
    var sandboxLabel = app.createLabel("Sandbox:");
    var sandbox = app.createCheckBox();
    sandbox.setName("sandbox");
    sandbox.setValue(getUseSandbox() == "true" ? true: false);
  
    var saveHandler = app.createServerClickHandler("saveConfiguration");
    var saveButton = app.createButton("Save Configuration", saveHandler);
  
    var listPanel = app.createGrid(4, 2);
    listPanel.setStyleAttribute("margin-top", "10px")
    listPanel.setWidth("100%");
    listPanel.setWidget(0, 0, usernameLabel);
    listPanel.setWidget(0, 1, username);
    listPanel.setWidget(1, 0, passwordLabel);
    listPanel.setWidget(1, 1, password);
    listPanel.setWidget(2, 0, securityTokenLabel);
    listPanel.setWidget(2, 1, securityToken);
    listPanel.setWidget(3, 0, sandboxLabel);
    listPanel.setWidget(3, 1, sandbox);
  
    // Ensure that all form fields get sent along to the handler
    saveHandler.addCallbackElement(listPanel);
  
    var dialogPanel = app.createFlowPanel();
    dialogPanel.add(helpLabel);
    dialogPanel.add(listPanel);
    dialogPanel.add(saveButton);
    app.add(dialogPanel);
    doc.show(app);

};

/**
 * Login script
 */
function login() {
  
    var message = "<?xml version='1.0' encoding='utf-8'?>" 
    + "<soap:Envelope xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/' " 
    +   "xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://"
    +   "www.w3.org/2001/XMLSchema'>" 
    +  "<soap:Body>" 
    +     "<login xmlns='urn:partner.soap.sforce.com'>" 
    +        "<username>" + getUsername() + "</username>"
    +        "<password>"+ getPassword() + getSecurityToken() + "</password>"
    +     "</login>" 
    +  "</soap:Body>" 
    + "</soap:Envelope>";
  
    var httpheaders = {SOAPAction: "login"};
    var parameters = {
        method : "POST",
        contentType: "text/xml",
        headers: httpheaders,
        payload : message
    };

    try {
        var result = UrlFetchApp.fetch(getSfdcSoapEndpoint(), parameters).getContentText();
        var soapResult = Xml.parse(result, false);

        setSessionId(soapResult.Envelope.Body.loginResponse.result.sessionId.getText());
        setInstanceUrl(soapResult.Envelope.Body.loginResponse.result.serverUrl.getText());
    } 
    catch(e) {
        Browser.msgBox(e);
    }

};

/**
 * Run SOQL Query in spreadsheet
 */
function SOQLQuery(SOQL) {
  var results = fetch(getRestEndpoint() + "/services/data/v30.0/" + "query?q=" + encodeURIComponent(SOQL));
  return renderGridData(Utilities.jsonParse(results));
};

/**
 * Clean data and get records
 */
function renderGridData(object, headers) {
  var data = [];
  var headersArray = [];
  
  // make headers array
  if (headers != undefined && headers.length > 0) {
    if (headers.indexOf(',') != -1) {
      var splitHeaders = headers.split(',');
      for (var header in splitHeaders) {
        headersArray.push(splitHeaders[header].trim());
      }
    }
    else {
      headersArray.push(headers.trim());
    }
  }
  
  
  for (var record in object.records) {
    var values = [];

    if (headersArray.length > 0) {
        for (var header in headersArray) {
          for (var property in object.records[record]) {
            if (object.records[record].hasOwnProperty(property)) {
              if (property != 'attributes') {
                if (Object.prototype.toString.call(object.records[record][property]) === '[object Object]') {
                  for (var subProperty in object.records[record][property]) {
                    if (subProperty != 'attributes') {
                      if (Object.prototype.toString.call(object.records[record][property][subProperty]) === '[object Object]') {
                        for (var subSubProperty in object.records[record][property][subProperty]) {
                          if (subSubProperty != 'attributes') {
                            if (headersArray[header] == subSubProperty) {
                              values.push(object.records[record][property][subProperty][subSubProperty]);
                            }
                          }
                        }
                      }
                      else {
                        if (headersArray[header] == subProperty) {
                          values.push(object.records[record][property][subProperty]);
                        }
                      }
                    }
                  }
                }
                else {
                  if (headersArray[header] == property) {
                    values.push(object.records[record][property]);
                  }
                }
              }
            }
          }
        }                
    }
    else {
      for (var property in object.records[record]) {
        if (object.records[record].hasOwnProperty(property)) {
          if (property != 'attributes') {
            if (Object.prototype.toString.call(object.records[record][property]) === '[object Object]') {
              for (var subProperty in object.records[record][property]) {
                if (subProperty != 'attributes') {
                  if (Object.prototype.toString.call(object.records[record][property][subProperty]) === '[object Object]') {
                    for (var subSubProperty in object.records[record][property][subProperty]) {
                      if (subSubProperty != 'attributes') {
                        values.push(object.records[record][property][subProperty][subSubProperty]);
                      }
                    }
                  }
                  else {
                    values.push(object.records[record][property][subProperty]);
                  }
                }
              }
            }
            else {
              values.push(object.records[record][property]);
            }
          }
        }
      }               
    }
    data.push(values);
  }
  return data;
};

/**
 * Get data from API
 */
function fetch(url) {
    var httpheaders = {Authorization: "OAuth " + getSessionId()};
    var parameters = {headers: httpheaders}; 
    return UrlFetchApp.fetch(url, parameters).getContentText();
};