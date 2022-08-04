const BASE_OKAPI = {   //1 -> YOUR OKAPI ENDPOINTS
    'prod': 'https://lehigh-okapi.folio.indexdata.com',
    'test': 'https://lehigh-test-okapi.folio.indexdata.com'
  };
  
  function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('FOLIO')
      .addItem('Show Sidebar', 'showSidebar')
      .addToUi();
  }
  
  function showSidebar() {  // eslint-disable-line no-unused-vars
    var html = HtmlService.createHtmlOutputFromFile('sidebar')
      .setTitle('Create Holdings and Bound-with Parts')
      .setWidth(500);
    SpreadsheetApp.getUi()
      .showSidebar(html);
  }
  
  function authenticate(config) {
    //AUTHENTICATE
    token = FOLIOAUTHLIBRARY.authenticate(BASE_OKAPI[config.environment]);
  
    getHeaders = {
      "Accept": "application/json",
      "x-okapi-tenant": "lu",
      "x-okapi-token": token
    };
    PropertiesService.getScriptProperties().setProperty("headers", JSON.stringify(getHeaders));
    getOptions = {
      'headers': getHeaders,
      'muteHttpExceptions': true
    }
    PropertiesService.getScriptProperties().setProperty("getOptions", JSON.stringify(getOptions));
  }
  
  function testLoadAndProcessRecords() {
    loadAndProcessRecords({
      'environment': 'test'
    });
  }
  
  function loadAndProcessRecords(config) {
    PropertiesService.getScriptProperties().setProperty("config", JSON.stringify(config));
    authenticate(config);
  
    // let fields = [
    //   ['PO Number', 'poNumber'],
    //   ['Status', 'workflowStatus']
    // ];
    
    let spreadsheet = SpreadsheetApp.getActiveSheet();
    let startRow = parseInt(config.start_row);
    let rowCount = parseInt(config.row_count);
  
    // writeHeaders(spreadsheet, fields);
    for (let row = startRow; row < startRow + rowCount; row++) {
        console.log("row #: " + row);
    //   loadAndCloseOrder(spreadsheet, row);
    }
  }
  
//   function writeHeaders(spreadsheet, fields) {
//     spreadsheet.getRange(1, 1, 1, fields.length).setValues([ fields.map(field => field[0]) ]);
//   }
  
  function loadAndCloseOrder(spreadsheet, row) {
    let poNumber = spreadsheet.getRange(row, 1).getValue();
    let order = loadOrder(poNumber);
    if (order == null) {
      console.log("No order on row " + row);
      return;
    }
    let result = closeOrder(order);
    updateSheet(spreadsheet, row, order, result);
  }
  
//   function loadOrder(poNumber) {
//     let config = JSON.parse(PropertiesService.getScriptProperties().getProperty("config"));
  
//     // execute query 
//     let ordersQuery = BASE_OKAPI[config.environment] + "/orders/composite-orders?query=(poNumber==" + poNumber + ")";
//     console.log("Loading order with query: ", ordersQuery);
//     let getOptions = JSON.parse(PropertiesService.getScriptProperties().getProperty("getOptions"));
//     let ordersResponse = UrlFetchApp.fetch(ordersQuery, getOptions);
  
//     // parse response
//     let ordersResponseText = ordersResponse.getContentText();
//     let orders = JSON.parse(ordersResponseText);
//     let purchaseOrders = orders.purchaseOrders;
    
//     if (purchaseOrders == null) {
//       return null;
//     }
//     return purchaseOrders[0];
//   }
  
//   function closeOrder(order) {
//     // set status
//     order.workflowStatus = "Closed";
  
//     // set reason & note
//     let username = Session.getEffectiveUser().getEmail();
//     order.closeReason = {
//       'reason': 'Cancelled',
//       'note': "Cancelled by " + username + " via Bulk Close Google App Script"
//     };
  
//     let response = putOrder(order);
//     if (response.getResponseCode() == 204) {
//       return "Closed";
//     }
//     else {
//       return "Unexpected response: " + JSON.stringify(response);
//     }
//   }
  
//   function putOrder(order) {
//     let config = JSON.parse(PropertiesService.getScriptProperties().getProperty("config"));
  
//     // execute query 
//     let url = BASE_OKAPI[config.environment] + "/orders/composite-orders/" + order.id;
//     let headers = JSON.parse(PropertiesService.getScriptProperties().getProperty("headers"));
//     let options = {
//       'method': 'put',
//       'contentType': 'application/json',
//       'headers': headers,
//       'payload': JSON.stringify(order),
//       'muteHttpExceptions': true,
//     }
//     console.log("Sending in modified order: ", order);
  
//     // parse response
//     let response = UrlFetchApp.fetch(url, options);
//     let responseContent = response.getContentText()
//     console.log("Got response " + JSON.stringify(response.getContentText()) + " from order " + order.poNumber);
  
//     return response;
//   }
  
//   function updateSheet(spreadsheet, row, order, result) {
//     spreadsheet.getRange(row, 2, 1, 1).setValue(result);
//   }
  