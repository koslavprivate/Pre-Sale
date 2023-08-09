const API_URL = "https://orgd59dbe8d.api.crm4.dynamics.com/api/data/v9.2/";

function updateToken() {
  const token = generateToken();
  Logger.log(token);
  PropertiesService.getScriptProperties().setProperty('Token', token);
}

function generateToken() {
    var payload = {
      grant_type: "client_credentials",
      client_id: "13df98dc-9dfc-4b2b-8255-dca0e171876f",
      client_secret: "rV98Q~Z6Rvm-.vRc0wqA_CCSX23SiByzxUm7Kahr",
      resource: "https://orgd59dbe8d.crm4.dynamics.com"
    };

    const payloadString = Object.keys(payload).map(function(key) {
      return encodeURIComponent(key) + "=" + encodeURIComponent(payload[key]);
    }).join("&");


    const headers = {
      "Content-Type": "application/x-www-form-urlencoded"
    };

    const response = UrlFetchApp.fetch("https://login.microsoftonline.com/e3a8922e-40b0-4804-8eb0-971f15f73877/oauth2/token", {
        method: "POST",
        headers: headers,
        payload: payloadString
    });

    const responseData = JSON.parse(response.getContentText());
    return responseData.access_token;
}

function readData() {   
   const token = PropertiesService.getScriptProperties().getProperty('Token');
   const DATA_ENDPOINT = "cr254_producationapplications";
   const response = UrlFetchApp.fetch(API_URL + DATA_ENDPOINT, {
       headers: {
           "Authorization": "Bearer " + token
       }
   });
   const content = JSON.parse(response.getContentText());
   console.log(content);
   return content;
}

function updateData() {
   const HEADERS = ['Name', 'Field 1', 'Field 2', 'Field 3', 'Id', '__PowerAppsId__']
   const spreadsheet = SpreadsheetApp.getActive();
   const sheet = spreadsheet.getSheetByName('Аркуш1');
   sheet.clear();
   sheet.appendRow(HEADERS);

  const headerRange = sheet.getRange(1, 1, 1, HEADERS.length - 1);
  headerRange.setBackground('yellow');

   const content = readData();
   const users = content.value; // Access the users using content.value

   if (Array.isArray(users)) {
       users.forEach(user => {
           const name = user.cr254_name; // Access the 'cr254_name' property for the name
           const field1 = user.cr254_field1; // Access the 'cr254_field1' property for Field 1
           const field2 = user.cr254_field2; // Access the 'cr254_field2' property for Field 2
           const field3 = user.cr254_field3; // Access the 'cr254_field3' property for Field 3
           const powerAppsId = user.cr254_producationapplicationid;
           sheet.appendRow([name, field1, field2, field3, powerAppsId]);
       });
   } else {
       console.log("No users found.");
   }
}

function cellsAreEmpty(row, col, sheet) {
  if (!sheet.getRange(row, col).isBlank()) {
    return false;
  }

  let resultColumns; 
  if (col === 1) {
    resultColumns = [2, 3, 4];
  } else if (col === 2) {
    resultColumns = [1, 3, 4];
  } else if (col === 3) {
    resultColumns = [1, 2, 4];
  } else if (col === 4) {
    resultColumns = [1, 3, 2];
  } else {
    return false;
  }
  
  for(let i = 0; i < resultColumns.length; i++) {
    if(!sheet.getRange(row, resultColumns[i]).isBlank()) {
      return false;
    }
  } 
  return true;
}

function onEditTrigger(e) {
  const spreadsheet = SpreadsheetApp.getActive();
  const sheet = spreadsheet.getSheetByName('Аркуш1');
  const oldValue = e.oldValue;
  const range = e.range;
  const row = range.getRow();
  const col = range.getColumn();
  const powerAppsIdValue = sheet.getRange(row, 5).getValue();
  
  if (cellsAreEmpty(row, col, sheet)) {
    deleteRecord(powerAppsIdValue);
    updateData();
  }
  else if (range.getNumRows() > 1 || range.getNumColumns() > 1) {
    updateData();
  }
  else if (col === 5) {
    deleteRecord(oldValue);
    updateData();
  }
  else if(col < 5 && sheet.getRange(row, 5).isBlank()) {
    createRecord(row, col, e.value);
  } 
  else if (col < 5) {
    updateProperty(powerAppsIdValue, col, e.value);
  }
  else {
    updateData();
  }
}

function createRecord(row, col, value) {
    const token = PropertiesService.getScriptProperties().getProperty('Token');
    let property = PropertyName(col);
    if (property === "cr254_field2") {
      value = parseFloat(value);
    }

    const DATA_ENDPOINT = "cr254_producationapplications";
    const headers = {
      "Authorization": "Bearer " + token,
      "Content-Type": "application/json",
      "OData-MaxVersion": "4.0",
      "OData-Version": "4.0",
      "Accept": "application/json",
      "Prefer": "return=representation"
    };

    const body = {
      [property] : value
    };

    const response = UrlFetchApp.fetch(API_URL + DATA_ENDPOINT, {
        method: "POST",
        headers: headers,
        payload: JSON.stringify(body)
    });
    const responseData = JSON.parse(response.getContentText());
    const idResult = responseData.cr254_producationapplicationid;
    setCellValue(row, 5, idResult);
  }

  function updateProperty(id, col, val) {
    const token = PropertiesService.getScriptProperties().getProperty('Token');
    let property = PropertyName(col);
    if (property === "cr254_field2") {
      val = parseFloat(val).toFixed(1);
  }
  
  const DATA_ENDPOINT = `cr254_producationapplications(${id})/${property}`;
  const headers = {
    "Authorization": "Bearer " + token,
    "Content-Type": "application/json",
    "OData-MaxVersion": "4.0",
    "OData-Version": "4.0"
  };

  const body = {
    value: val
  }

  const options = {
    method: "PUT",
    headers: headers,
    payload: JSON.stringify(body)
  };

  const response = UrlFetchApp.fetch(API_URL + DATA_ENDPOINT, options);
}

function updateProperty(id, col, val) {
  const token = PropertiesService.getScriptProperties().getProperty('Token');
  let property = PropertyName(col);
  
  if (val === undefined) {
    val = null;
  }
  if (property === "cr254_field2") {
    val = parseFloat(val);
  }
  
  const DATA_ENDPOINT = `cr254_producationapplications(${id})/${property}`;
  const headers = {
    "Authorization": "Bearer " + token,
    "Content-Type": "application/json",
    "OData-MaxVersion": "4.0",
    "OData-Version": "4.0"
  };

  const body = {
    value: val
  }

  const options = {
    method: "PUT",
    headers: headers,
    payload: JSON.stringify(body)
  };

  const response = UrlFetchApp.fetch(API_URL + DATA_ENDPOINT, options);
}

function deleteRecord(id) {
  const token = PropertiesService.getScriptProperties().getProperty('Token');
  const DATA_ENDPOINT = `cr254_producationapplications(${id})`;
  const headers = {
    "Authorization": "Bearer " + token,
    "Content-Type": "application/json",
    "OData-MaxVersion": "4.0",
    "OData-Version": "4.0"
  };
  
  const response = UrlFetchApp.fetch(API_URL + DATA_ENDPOINT, {
       method: "DELETE",
       headers: headers,
   });
}

function PropertyName(col) {
  if (col === 1) {
    return "cr254_name";
  }
  else if (col === 2) {
    return "cr254_field1";
  }
  else if (col === 3) {
    return "cr254_field2";
    val = parseFloat(val);
  }
  else if (col === 4) {
    return "cr254_field3";
  }
}

function setCellValue(row, col, value) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName('Аркуш1'); // Замените на имя вашего листа
  const cell = sheet.getRange(row, col);
  cell.setValue(value);
}