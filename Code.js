/*
  Required FOLIO permissions:
  - inventory.items.collection.get
  - inventory-storage.holdings.item.get
  - inventory-storage.holdings.item.post
  - inventory-storage.bound-with-parts.item.post
*/

const ITEM_BARCODE_COLUMN = 'D';
const INSTANCE_UUID_COLUMN = 'K';
const LAST_COLUMN = 'L';
const SOURCE_ID_FOLIO = "f32d531e-df79-46b3-8932-cdd35f7a2264";
const STATISTICAL_CODE_RETENTION_AGREEMENT = "ba16cd17-fb83-4a14-ab40-23c7ffa5ccb5";

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
  
function testLoadAndProcessRecords() {
    loadAndProcessRecords({
      'environment': 'test',
      'start_row': 2,
      'end_row': 6
    });
}
  
function loadAndProcessRecords(config) {
    PropertiesService.getScriptProperties().setProperty("config", JSON.stringify(config));
    config.username = PropertiesService.getScriptProperties().getProperty("username");
    config.password = Utilities.newBlob(Utilities.base64Decode(
        PropertiesService.getScriptProperties().getProperty("password")))
        .getDataAsString();
    FOLIOAUTHLIBRARY.authenticateAndSetHeaders(config);
  
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
            currentHoldingRecord = cloneHoldingForNewInstance(primaryHoldingRecord, instanceUuid, currentItem);
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
    let itemQuery = FOLIOAUTHLIBRARY.getBaseOkapi(config.environment) + 
        "/inventory/items?query=(barcode==" + barcode + ")";
    console.log("Loading item with query: ", itemQuery);
    let getOptions = FOLIOAUTHLIBRARY.getHttpGetOptions();
    let response = UrlFetchApp.fetch(itemQuery, getOptions);
    if (response.getResponseCode() != 200) {
        throw new Error("Cannot get item records, response: " + response);
    }

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
    let holdingRecordQuery = FOLIOAUTHLIBRARY.getBaseOkapi(config.environment) + 
        "/holdings-storage/holdings/" + id;
    console.log("Loading holding record with query: ", holdingRecordQuery);
    let getOptions = FOLIOAUTHLIBRARY.getHttpGetOptions();
    let response = UrlFetchApp.fetch(holdingRecordQuery, getOptions);
    if (response.getResponseCode() != 200) {
        throw new Error("Cannot get holding record, response: " + response);
    }
    
    // parse response
    let responseText = response.getContentText();
    let responseObject = JSON.parse(responseText);
    let holdingRecord = responseObject;
    return holdingRecord;
}

function cloneHoldingForNewInstance(primaryHoldingRecord, instanceUuid, itemRecord) {
    let holdingRecord = Object.assign({}, primaryHoldingRecord);

    // Set the new instanceId
    holdingRecord.instanceId = instanceUuid;

    // Remove inappropriate fields
    holdingRecord.id = null;
    holdingRecord.hrid = null;
    holdingRecord.formerIds = null;
    holdingRecord.metadata = null;
    holdingRecord.retentionPolicy = null;
    clearRetentionAgreements(holdingRecord);

    // Set the source type to explicitly be FOLIO
    holdingRecord.sourceId = SOURCE_ID_FOLIO;

    // Add holding note referencing the item barcode
    let config = JSON.parse(PropertiesService.getScriptProperties().getProperty("config"));
    let noteType = (config.environment == "prod") 
        ? "479353a3-15df-4deb-b03e-a45289196d01" 
        : "28694388-1395-4373-b620-d3269dfcfc70";
    let noteText = itemRecord.barcode;
    let note = {
        holdingsNoteTypeId: noteType,
        note: noteText,
        staffOnly: true
    };
    holdingRecord.notes = [...holdingRecord.notes];
    holdingRecord.notes.push(note);

    // Execute post 
    let url = FOLIOAUTHLIBRARY.getBaseOkapi(config.environment) + "/holdings-storage/holdings";
    let headers = FOLIOAUTHLIBRARY.getHttpGetHeaders();
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
    if (response.getResponseCode() != 201) {
        throw new Error("Cannot create holding record, response: " + response);
    }

    let createdHoldingRecord = JSON.parse(responseContent);
    return createdHoldingRecord;
}

function clearRetentionAgreements(holdingRecord) {
    let statisticalCodeIds = [];
    for (const statisticalCodeId of holdingRecord.statisticalCodeIds) {
        if (statisticalCodeId != STATISTICAL_CODE_RETENTION_AGREEMENT) {
            statisticalCodeIds.push(statisticalCodeId);
        }
    }
    holdingRecord.statisticalCodeIds = statisticalCodeIds;
}

function createBoundWithPart(item, holdingRecord) {
    let config = JSON.parse(PropertiesService.getScriptProperties().getProperty("config"));

    // execute query 
    let url = FOLIOAUTHLIBRARY.getBaseOkapi(config.environment) + "/inventory-storage/bound-with-parts";
    let headers = FOLIOAUTHLIBRARY.getHttpGetHeaders();
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

    if (response.getResponseCode() != 201) {
        throw new Error("Cannot create bound-with part, response: " + response);
    }
    return true;
}
  
function updateSheet(spreadsheet, row, success) {
    let color = success ? "lightgreen" : "lightcoral";
    let range = spreadsheet.getRange("A" + row + ":" + LAST_COLUMN + row);
    range.setBackground(color);
}