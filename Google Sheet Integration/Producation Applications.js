let dataGoogleSheets;
const PERSONAL_ACCESS_TOKEN = "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6Ii1LSTNROW5OUjdiUm9meG1lWm9YcWJIWkdldyIsImtpZCI6Ii1LSTNROW5OUjdiUm9meG1lWm9YcWJIWkdldyJ9.eyJhdWQiOiJodHRwczovL29yZ2Q1OWRiZThkLmNybTQuZHluYW1pY3MuY29tIiwiaXNzIjoiaHR0cHM6Ly9zdHMud2luZG93cy5uZXQvZTNhODkyMmUtNDBiMC00ODA0LThlYjAtOTcxZjE1ZjczODc3LyIsImlhdCI6MTY5MTA2NjkyOCwibmJmIjoxNjkxMDY2OTI4LCJleHAiOjE2OTEwNzA4MjgsImFpbyI6IkUyRmdZSGg5dGUyUm0zV3ppWm5NVnNhY0IwSENBQT09IiwiYXBwaWQiOiIxM2RmOThkYy05ZGZjLTRiMmItODI1NS1kY2EwZTE3MTg3NmYiLCJhcHBpZGFjciI6IjEiLCJpZHAiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC9lM2E4OTIyZS00MGIwLTQ4MDQtOGViMC05NzFmMTVmNzM4NzcvIiwib2lkIjoiNWIxZDIzOTEtOGQxYy00ZGVjLWEwNGUtMGY4OTY5MDE1MTMyIiwicmgiOiIwLkFhOEFMcEtvNDdCQUJFaU9zSmNmRmZjNGR3Y0FBQUFBQUFBQXdBQUFBQUFBQUFDdkFBQS4iLCJzdWIiOiI1YjFkMjM5MS04ZDFjLTRkZWMtYTA0ZS0wZjg5NjkwMTUxMzIiLCJ0aWQiOiJlM2E4OTIyZS00MGIwLTQ4MDQtOGViMC05NzFmMTVmNzM4NzciLCJ1dGkiOiJDQ3JPMXpodm8waUZ5TF9odXBZQ0FBIiwidmVyIjoiMS4wIiwieG1zX2NhZSI6IjEifQ.Tj32TngQO-Yx5wF5eNSXUOolBviI0n3cwx6IGR7IPdKHWxeVgEYG-hVlJe1QiXvI5w77KCI2xBFDCPFhAev6BlEsE9DBbqLdqGS8RYa34c3LVfrNInBnKSihIyThaEDOTjsX6k5zTG2mY361HU5g9EidrI31nxEqAjh7s2ZTVwVH71GWK22MB8tP8qvS5LmxcpAnrPE5KhJzt3B3jD0WUjpCyS4CIETXY96p35iu9odatsFtf5WemlTqxkYz7pHqdcde7y0Rr28C3KM_72HJymq9yPnFDdS7ZBmZmcrT2zN1mUUVkhyzcLvn6MvXiv9IaM2S0duNvKae3MiV8Vu90g";
const API_URL = "https://orgd59dbe8d.api.crm4.dynamics.com/api/data/v9.2/";


function getDataGoogleSheets() {
  const spreadsheet = SpreadsheetApp.getActive();
  const sheet = spreadsheet.getSheetByName('Аркуш1');
  const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
  dataGoogleSheets = dataRange.getValues();
  console.log(dataGoogleSheets);
}

function readData() {   
   const DATA_ENDPOINT = "cr254_producationapplications";
   const response = UrlFetchApp.fetch(API_URL + DATA_ENDPOINT, {
       headers: {
           "Authorization": "Bearer " + PERSONAL_ACCESS_TOKEN
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
    let property = PropertyName(col);
    if (property === "cr254_field2") {
      value = parseFloat(value);
    }

    const DATA_ENDPOINT = "cr254_producationapplications";
    const headers = {
      "Authorization": "Bearer " + PERSONAL_ACCESS_TOKEN,
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
    let property = PropertyName(col);
    if (property === "cr254_field2") {
      val = parseFloat(val).toFixed(1);
  }
  
  const DATA_ENDPOINT = `cr254_producationapplications(${id})/${property}`;
  const headers = {
    "Authorization": "Bearer " + PERSONAL_ACCESS_TOKEN,
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
  let property = PropertyName(col);
  
  if (val === undefined) {
    val = null;
  }
  if (property === "cr254_field2") {
    val = parseFloat(val);
  }
  
  const DATA_ENDPOINT = `cr254_producationapplications(${id})/${property}`;
  const headers = {
    "Authorization": "Bearer " + PERSONAL_ACCESS_TOKEN,
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
  const DATA_ENDPOINT = `cr254_producationapplications(${id})`;
  const headers = {
    "Authorization": "Bearer " + PERSONAL_ACCESS_TOKEN,
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




















function getAccessToken() {
  eval(UrlFetchApp.fetch('https://alcdn.msauth.net/browser/2.14.2/js/msal-browser.min.js').getContentText());
  const msalConfig = {
    auth: {
      clientId: '13df98dc-9dfc-4b2b-8255-dca0e171876f', // Замените на ваш client_id из Azure AD
      authority: 'https://login.microsoftonline.com/e3a8922e-40b0-4804-8eb0-971f15f73877', // Замените на ваш tenant_id из Azure AD
      redirectUri: 'https://orgd59dbe8d.crm4.dynamics.com' // Замените на URL вашего приложения
    }
  };

  // Создание экземпляра объекта MSAL
  const msalInstance = new msal.PublicClientApplication(msalConfig);

  // Запрашиваемый область (scopes) и другие параметры для запроса токена доступа
  const request = {
    scopes: ['user.read'], // Замените на необходимые области (scopes) вашего приложения
    prompt: 'select_account' // Дополнительные параметры, если необходимо
  };

  // Выполняем процесс авторизации и получаем токен доступа
  msalInstance.loginPopup(request).then(response => {
    // Получаем токен доступа из ответа
    const accessToken = response.accessToken;

    // Используйте полученный токен доступа для доступа к защищенным ресурсам на сервере
    console.log('Access token:', accessToken);
  }).catch(error => {
    // Обработка ошибок, если произошла ошибка авторизации
    console.log('An error occurred during login:', error);
});
}