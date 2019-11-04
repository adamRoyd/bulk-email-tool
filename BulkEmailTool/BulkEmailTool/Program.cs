using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;

namespace BulkEmailTool
{
    class Program
    {
        static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static readonly string ApplicationName = "Current Legislators";
        static readonly string SpreadsheetId = "12_gF3sRVk57oD2z8_Yb7mro-DSwAder-lobjC3S71t4";
        static readonly string sheet = "Submissions";
        static SheetsService service;

        static void Main(string[] args)
        {
            GoogleCredential credential;
            using (var stream = new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream)
                    .CreateScoped(Scopes);
            }

            // Create Google Sheets API service.
            service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            ReadEntries();

            Console.ReadKey();
        }

        static void ReadEntries()
        {
            var range = $"{sheet}!F:I";
            SpreadsheetsResource.ValuesResource.GetRequest request =
                service.Spreadsheets.Values.Get(SpreadsheetId, range);

            var response = request.Execute();
            IList<IList<object>> values = response.Values;
            if (values != null && values.Count > 0)
            {
                var promoters = GetPromoters(values);

                foreach (var promoter in promoters)
                {
                    Console.WriteLine("{0} | {1} | {2}", promoter.Name, promoter.Type.ToString(), promoter.EmailAddress);
                }
            }
            else
            {
                Console.WriteLine("No data found.");
            }
        }

        private static IEnumerable<Promoter> GetPromoters(IList<IList<object>> values)
        {
            var filteredRows = values.Where(entry => entry.Count == 4 && entry[0].ToString() == "Email").ToList();

            IList<Promoter> promoters = new List<Promoter>();

            foreach (var row in filteredRows)
            {
                Enum.TryParse(row[3].ToString(), out PlatformType type);

                promoters.Add(new Promoter
                {
                    Name = row[1].ToString(),
                    EmailAddress = row[2].ToString(),
                    Type = type
                });
            }

            return promoters;
        }
    }
}
