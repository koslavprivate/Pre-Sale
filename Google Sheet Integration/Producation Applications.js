let dataGoogleSheets;
const PERSONAL_ACCESS_TOKEN = "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6Ii1LSTNROW5OUjdiUm9meG1lWm9YcWJIWkdldyIsImtpZCI6Ii1LSTNROW5OUjdiUm9meG1lWm9YcWJIWkdldyJ9.eyJhdWQiOiJodHRwczovL29yZ2Q1OWRiZThkLmNybTQuZHluYW1pY3MuY29tIiwiaXNzIjoiaHR0cHM6Ly9zdHMud2luZG93cy5uZXQvZTNhODkyMmUtNDBiMC00ODA0LThlYjAtOTcxZjE1ZjczODc3LyIsImlhdCI6MTY5MDk1NjEzNiwibmJmIjoxNjkwOTU2MTM2LCJleHAiOjE2OTA5NjAwMzYsImFpbyI6IkUyRmdZR0E1R3IxZHlTbFYzaWJ0NExlTVd4T2ZBUUE9IiwiYXBwaWQiOiIxM2RmOThkYy05ZGZjLTRiMmItODI1NS1kY2EwZTE3MTg3NmYiLCJhcHBpZGFjciI6IjEiLCJpZHAiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC9lM2E4OTIyZS00MGIwLTQ4MDQtOGViMC05NzFmMTVmNzM4NzcvIiwib2lkIjoiNWIxZDIzOTEtOGQxYy00ZGVjLWEwNGUtMGY4OTY5MDE1MTMyIiwicmgiOiIwLkFhOEFMcEtvNDdCQUJFaU9zSmNmRmZjNGR3Y0FBQUFBQUFBQXdBQUFBQUFBQUFDdkFBQS4iLCJzdWIiOiI1YjFkMjM5MS04ZDFjLTRkZWMtYTA0ZS0wZjg5NjkwMTUxMzIiLCJ0aWQiOiJlM2E4OTIyZS00MGIwLTQ4MDQtOGViMC05NzFmMTVmNzM4NzciLCJ1dGkiOiJqQi11ak9qSkFFeVp5c296QXFnT0FBIiwidmVyIjoiMS4wIiwieG1zX2NhZSI6IjEifQ.qxAofTKgkiKRP_YrxORG9rmKdgm5WCZZMfDWRL4tj5fxXQ-JqFcfoPJVxn-7ocnNVhE0inoG5dBbgh_B6C-CRaE12PKredK0xsW1dsWMpKR8Hc0JP2Ru_uQHPpWi6eszS-pSrkesodUnfFq9qFif5JSjb6Kvh1PoLg7VfYIOnkQwbdjkL12fnBlEXAUTZrTqggj7iWwJsa1NMokJUUl1z62o0dyIdCKT1bZGuHLcPnJhMFoHxqG9tOphALZKbeWIOF6mG9ajYWPiyGrnpPpmy8Vjnt9bzVB2nQ25wntYg2H2FZj6VzXMl6vHha3wAKzJOvkcz2mF-QXfvqhqcnCPIA";
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

//not work
function createRecord() {;
  const DATA_ENDPOINT = "accounts";

  const headers = {
    "Authorization": "Bearer " + PERSONAL_ACCESS_TOKEN,
    "Content-Type": "application/json",
    "OData-MaxVersion": "4.0",
    "OData-Version": "4.0",
    "Accept": "application/json",
    "Prefer": "return=representation"
  };

  const body = {
    "cr254_name": "121231233"
  };

  const response = UrlFetchApp.fetch(API_URL + DATA_ENDPOINT, {
       method: "POST",
       headers: headers,
       body: body
   });
  const responseData = JSON.parse(response.getContentText());
  
  // Получаем созданный идентификатор записи
  const accountID = responseData.accountid;
  // Выводим информацию о созданной записи
  console.log("Created account with ID:", accountID);
  console.log("Name:", responseData.name);
  console.log("Description:", responseData.description);
  console.log("Field 1:", responseData.field1);
  console.log("Field 2:", responseData.field2);
  console.log("Field 3:", responseData.field3);
}



function updateData() {
  
   const HEADERS = ['Name', 'Field 1', 'Field 2', 'Field 3', 'PowerAppsId']
   const spreadsheet = SpreadsheetApp.getActive();
   const sheet = spreadsheet.getSheetByName('Аркуш1');
   sheet.clear();
   sheet.appendRow(HEADERS);

  const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
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

function onEdit(e) {
  const spreadsheet = SpreadsheetApp.getActive();
  const sheet = spreadsheet.getSheetByName('Аркуш1');
  const oldValue = e.oldValue;
  const range = e.range;
  const row = range.getRow();
  const col = range.getColumn();
  const powerAppsIdValue = sheet.getRange(row, 5).getValue();
  
  if (col === 5) {
    PropertiesService.getScriptProperties().setProperty('recordsToDelete', oldValue);
  } else {
    range.setValue(PropertiesService.getScriptProperties().getProperty('recordsToDelete'))
  }
}

function processDeletions() {
   deleteRecord(PropertiesService.getScriptProperties().getProperty('recordsToDelete'));
   updateData();
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


















































// not work
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