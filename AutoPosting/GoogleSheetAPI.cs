using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;

namespace AutoPosting
{
    class GoogleSheetAPI
    {
        private static string[] m_Scopes = { SheetsService.Scope.Spreadsheets };

        private SheetsService m_Service;

        private GoogleCredential m_Credential;

        private String m_SpreadsheetId;

        private string m_SheetRange;

        public GoogleSheetAPI(string i_SpreadsheetId, string i_SheetRange)
        {
            m_SpreadsheetId = i_SpreadsheetId;
            m_SheetRange = i_SheetRange;
            setCredentials();
            setService();
        }

        //Sets google credentials to open GoogleAPI service. 
        private void setCredentials()
        {
            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                m_Credential = GoogleCredential.FromStream(stream)
                    .CreateScoped(m_Scopes);
            }

        }

        //Sets a servise conection to GoogleAPI.
        private void setService()
        {
            m_Service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = m_Credential,
                ApplicationName = "Google Sheets API .NET Quickstart",
            });
        }

        //Returns the desire sheet's range.
        public ValueRange GetSheet()
        {
            SpreadsheetsResource.ValuesResource.GetRequest request =
 m_Service.Spreadsheets.Values.Get(m_SpreadsheetId, m_SheetRange);
            ValueRange response = request.Execute();
            
            return response;
        }

        //Deletes the confession row to set the next confession to the front. 
        public void deleteRow()
        {
            var request = new Request
            {
                DeleteDimension = new DeleteDimensionRequest
                {
                    Range = new DimensionRange
                    {
                        SheetId = 14066518,
                        Dimension = "ROWS",
                        StartIndex = 2,
                        EndIndex = 3
                    }
                }
            };
            var re = new BatchUpdateSpreadsheetRequest();
            re.Requests = new List<Request>() { request };

            var b = m_Service.Spreadsheets.BatchUpdate(re, m_SpreadsheetId).Execute();
            
        }

        public void UpdateCell(string i_Range, List<Object> i_ObjList)
        {
            
            ValueRange valueRange = new ValueRange();
            valueRange.Values = new List<IList<Object>>() { i_ObjList };

            var updateReq = m_Service.Spreadsheets.Values.Update(valueRange, m_SpreadsheetId, i_Range);
            updateReq.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            updateReq.Execute();

        }
    }
}
