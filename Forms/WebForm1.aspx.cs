using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Threading;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace GetDataFromGoogleSheet.Forms
{
    public partial class WebForm1 : System.Web.UI.Page
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/sheets.googleapis.com-dotnet-quickstart.json
        static readonly string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly };
        static readonly string ApplicationName = "Application Name";

        [Obsolete]
        protected void Page_Load(object sender, EventArgs e)
        {
            UserCredential credential;

            using (var stream = new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
            }

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            // Define request parameters.
            String spreadsheetId = "spreadsheetId";
            String range = "Class Data!A2:E";
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    service.Spreadsheets.Values.Get(spreadsheetId, range);

            // Import Data From a sample spreadsheet To Sqlserver :
            ValueRange response = request.Execute();
            IList<IList<Object>> values = response.Values;
            if (values != null && values.Count > 0)
            {
                using (SqlConnection con = new SqlConnection("Connection String"))
                {
                    foreach (var row in values)
                    {
                        // Insert columns A and E, which correspond to indices 0 and 4.
                        SqlCommand cmd = new SqlCommand("INSERT INTO [dbo].[Asset]([Asset],[Asset Name],[Model],[Vendor],[Description]) " +
                            "VALUES(" + Convert.ToInt32(row[0]) + "," + row[1] + "," + row[2] + "," + row[3] + "," + row[4] + ")", con);
                        cmd.ExecuteNonQuery();
                    }
                }
            }
        }
    }
}