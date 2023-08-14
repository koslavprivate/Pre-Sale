using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using PreSale.Properties;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static Google.Apis.Sheets.v4.SpreadsheetsResource;
using static Google.Apis.Sheets.v4.SpreadsheetsResource.ValuesResource;

namespace PreSale
{
    public class GoogleSheetsHelper
    {
        private readonly string SHEET_NAME;
        private readonly string ReadRange;
        private readonly string SpreadsheetId;
        private readonly string GoogleCredentialsFileName;
        private static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
        SpreadsheetsResource.ValuesResource _googleSheetValues;

        public GoogleSheetsHelper()
        {
            SHEET_NAME = "Аркуш1";
            ReadRange = $"{SHEET_NAME}!E:E";
            SpreadsheetId = "1LQBhkysYphFHl1mi3o1psEntbkY49f1mmIyR0m-SHJI";
            GoogleCredentialsFileName = "client_secrets.json";
            _googleSheetValues = GetSheetsService().Spreadsheets.Values;
        }

        public SheetsService GetSheetsService()
        {
            using (var stream = new FileStream(GoogleCredentialsFileName, FileMode.Open, FileAccess.Read))
            {
                var serviceInitializer = new BaseClientService.Initializer
                {
                    HttpClientInitializer = GoogleCredential.FromStream(stream).CreateScoped(Scopes)
                };
                return new SheetsService(serviceInitializer);
            }
        }

        public AppendValuesResponse Create(cr254_ProducationApplication application)
        {
            var range = $"{SHEET_NAME}!A:D";
            var valueRange = new ValueRange
            {
                Values = ProducationApplicationMapper.MapToRangeData(application)
            };
            var appendRequest = _googleSheetValues.Append(valueRange, SpreadsheetId, range);
            appendRequest.ValueInputOption = AppendRequest.ValueInputOptionEnum.USERENTERED;
            return appendRequest.Execute();
        }

        public UpdateValuesResponse Update(Guid id, cr254_ProducationApplication application)
        {
            var response = _googleSheetValues.Get(SpreadsheetId, ReadRange).Execute();
            var values = response.Values;
            var numOfRow = 1;
            var range = "";

            foreach (var row in values.Skip(1))
            {
                numOfRow++;
                if (row[0].ToString() == id.ToString())
                {
                    range = $"{SHEET_NAME}!A{numOfRow}:E{numOfRow}";
                }
            }

            var valueRange = new ValueRange
            {
                Values = ProducationApplicationMapper.MapToRangeData(application)
            };

            var updateRequest = _googleSheetValues.Update(valueRange, SpreadsheetId, range);
            updateRequest.ValueInputOption = UpdateRequest.ValueInputOptionEnum.USERENTERED;
            return updateRequest.Execute();
        }

        public ClearValuesResponse Delete(Guid id)
        {
            var response = _googleSheetValues.Get(SpreadsheetId, ReadRange).Execute();
            var values = response.Values;
            var numOfRow = 1;
            var range = "";

            foreach (var row in values.Skip(1))
            {
                numOfRow++;
                if (row[0].ToString() == id.ToString())
                {
                    range = $"{SHEET_NAME}!A{numOfRow}:E{numOfRow}";
                }
            }

            var requestBody = new ClearValuesRequest() ;
            var deleteRequest = _googleSheetValues.Clear(requestBody, SpreadsheetId, range);
            return deleteRequest.Execute();
        }
    }
}
