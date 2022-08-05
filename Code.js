const BASE_OKAPI = {   //1 -> YOUR OKAPI ENDPOINTS
    'prod': 'https://lehigh-okapi.folio.indexdata.com',
    'test': 'https://lehigh-test-okapi.folio.indexdata.com'
  };

const ITEM_BARCODE_COLUMN = 'D';
const INSTANCE_UUID_COLUMN = 'K';
const LAST_COLUMN = 'L';
const SOURCE_ID_FOLIO = "f32d531e-df79-46b3-8932-cdd35f7a2264";
  
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
      'environment': 'test',
      'start_row': 2,
      'end_row': 6
    });
}
  
function loadAndProcessRecords(config) {
    PropertiesService.getScriptProperties().setProperty("config", JSON.stringify(config));
    authenticate(config);
  
    let spreadsheet = SpreadsheetApp.getActiveSheet();
    let startRow = parseInt(config.start_row);
    let endRow = parseInt(config.end_row);
  
    // Ensure the first row has a barcode.
    if (getBarcode(spreadsheet, startRow) == null) {
        console.error("Start row must have an item barcode.");
        return;
    }
  
    let currentItem = null;
    let primaryHoldingRecord = null;
    for (let row = startRow; row <= endRow; row++) {
        console.log("Starting on row #" + row);
        let rowBarcode = getBarcode(spreadsheet, row);
        let currentHoldingRecord = null;

        // If the row represents a new set of bound-with parts
        if (rowBarcode != "") {
            // Identify the item and holding record.
            currentItem = getItemRecord(rowBarcode);
            primaryHoldingRecord = getHoldingRecord(currentItem['holdingsRecordId']);
            currentHoldingRecord = primaryHoldingRecord;
        }
        // Otherwise the row should use the prior currentItem
        else {
            // Create a new holding record linked to the new row's instance
            let instanceUuid = getInstanceUuid(spreadsheet, row);
            currentHoldingRecord = cloneHoldingForNewInstance(primaryHoldingRecord, instanceUuid);
            if (currentHoldingRecord == null) {
                console.error("Failed to clone holding record for instanceId " + instanceUuid +  ": " + primaryHoldingRecord)
                updateSheet(spreadsheet, row, false);
                continue;
            }
        }

        // Create a bound-with part
        let result = createBoundWithPart(currentItem, currentHoldingRecord);
        updateSheet(spreadsheet, row, result);
    }
}

function getBarcode(spreadsheet, row) {
    return spreadsheet.getRange(ITEM_BARCODE_COLUMN + row).getValue();
}

function getInstanceUuid(spreadsheet, row) {
    return spreadsheet.getRange(INSTANCE_UUID_COLUMN + row).getValue();
}

function getItemRecord(barcode) {
    let config = JSON.parse(PropertiesService.getScriptProperties().getProperty("config"));
    
    // execute query 
    let itemQuery = BASE_OKAPI[config.environment] + "/inventory/items?query=(barcode==" + barcode + ")";
    console.log("Loading item with query: ", itemQuery);
    let getOptions = JSON.parse(PropertiesService.getScriptProperties().getProperty("getOptions"));
    let response = UrlFetchApp.fetch(itemQuery, getOptions);
    
    // parse response
    let responseText = response.getContentText();
    let responseObject = JSON.parse(responseText);
    let items = responseObject.items;
    
    if (items == null || items.length == 0) {
        console.error("No item for barcode: " + barcode);
        return null;
    }

    return items[0];
}

function getHoldingRecord(id) {
    let config = JSON.parse(PropertiesService.getScriptProperties().getProperty("config"));
    
    // execute query 
    let holdingRecordQuery = BASE_OKAPI[config.environment] + "/holdings-storage/holdings/" + id;
    console.log("Loading holding record with query: ", holdingRecordQuery);
    let getOptions = JSON.parse(PropertiesService.getScriptProperties().getProperty("getOptions"));
    let response = UrlFetchApp.fetch(holdingRecordQuery, getOptions);
    
    // parse response
    let responseText = response.getContentText();
    let responseObject = JSON.parse(responseText);
    let holdingRecord = responseObject;
    return holdingRecord;
}

function cloneHoldingForNewInstance(primaryHoldingRecord, instanceUuid) {
    let holdingRecord = Object.assign({}, primaryHoldingRecord);

    // Set the new instanceId
    holdingRecord.instanceId = instanceUuid;

    // Remove inappropriate fields
    holdingRecord.id = null;
    holdingRecord.hrid = null;
    holdingRecord.formerIds = null;
    holdingRecord.metadata = null;

    // Set the source type to explicitly be FOLIO
    holdingRecord.sourceId = SOURCE_ID_FOLIO;

    // Execute post 
    let config = JSON.parse(PropertiesService.getScriptProperties().getProperty("config"));
    let url = BASE_OKAPI[config.environment] + "/holdings-storage/holdings";
    let headers = JSON.parse(PropertiesService.getScriptProperties().getProperty("headers"));
    let options = {
      'method': 'POST',
      'contentType': 'application/json',
      'headers': headers,
      'payload': JSON.stringify(holdingRecord),
      'muteHttpExceptions': true,
    }
    console.log("Sending in new holding record: ", holdingRecord);

    // Parse response
    let response = UrlFetchApp.fetch(url, options);
    let responseContent = response.getContentText()
    let statusCode = response.getResponseCode();
    console.log("Got response with code " + statusCode + ": " + JSON.stringify(response.getContentText()));
    if (statusCode != 201) {
        console.error("Unexpected status code.")
        return null;
    }

    let createdHoldingRecord = JSON.parse(responseContent);
    return createdHoldingRecord;
}

function createBoundWithPart(item, holdingRecord) {
    let config = JSON.parse(PropertiesService.getScriptProperties().getProperty("config"));

    // execute query 
    let url = BASE_OKAPI[config.environment] + "/inventory-storage/bound-with-parts";
    let headers = JSON.parse(PropertiesService.getScriptProperties().getProperty("headers"));
    let boundWithPart = {
        "holdingsRecordId": holdingRecord.id,
        "itemId": item.id
    };
    let options = {
      'method': 'POST',
      'contentType': 'application/json',
      'headers': headers,
      'payload': JSON.stringify(boundWithPart),
      'muteHttpExceptions': true,
    }
    console.log("Sending in new bound-with part: ", boundWithPart);
  
    // parse response
    let response = UrlFetchApp.fetch(url, options);
    let responseContent = response.getContentText()
    let statusCode = response.getResponseCode();
    console.log("Got response with code " + statusCode + ": " + JSON.stringify(response.getContentText()));

    if (statusCode != 201) {
        console.error("Unexpected status code.")
        return false;
    }
    return true;
}
  
function updateSheet(spreadsheet, row, success) {
    let color = success ? "lightgreen" : "lightcoral";
    let range = spreadsheet.getRange("A" + row + ":" + LAST_COLUMN + row);
    range.setBackground(color);
}