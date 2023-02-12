using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;

namespace GoogleSheetAPI
{
    internal class Program
    {
        static string[] Scoped = { SheetsService.Scope.Spreadsheets };
        static string applicationName = "Desktop client 1";
        static string spreadsheetID = "1XAaPYpH_mNH2Xfe5tPfV7CexN-4VOZmj77QyPIwzBIw"; // write you spreadsheetId

        static void Main(string[] args)
        {
            GoogleCredential credential;

            // download this file from your google workspace (Sevice Account)
            using (var stream = new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream).CreateScoped(Scoped);
            }

            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = applicationName
            });


            // Expample for updation the sheet

            var range = "Sheet1!A1:B2";  // {sheetsName! From:To}

            // read this https://developers.google.com/sheets/api/reference/rest/v4/ValueInputOption
            SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum valueInputOption = (SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum)1;

            ValueRange values = new ValueRange();
            List<object> firstRow = new List<object> { "Name: ", "Danil" };
            List<object> secondRow = new List<object> { "Age: ", 20 };

            values.Values = new List<IList<object>> { firstRow, secondRow };

            //create request
            var request = service.Spreadsheets.Values.Update(values, spreadsheetID, range);
            request.ValueInputOption = valueInputOption;

            var response = request.Execute();
            Console.WriteLine(response.UpdatedCells); // number of updated cells


            // Example for getting the sheet

            var req = service.Spreadsheets.Values.Get(spreadsheetID, range);

            var res = req.Execute();

            if (res.Values != null && res.Values.Count > 0)
            {
                foreach (var item in res.Values)
                {
                    foreach (var cell in item)
                    {
                        Console.Write(cell);
                    }
                    Console.WriteLine();
                }
            }
            else
                Console.WriteLine("No data available");
        }
    }
}