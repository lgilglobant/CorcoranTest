using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using Swashbuckle.Swagger.Annotations;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Reflection;
using System.Web;
using Newtonsoft.Json;

namespace CorcoranTest.Controllers
{
    public class ValuesController : ApiController
    {
        static string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly };
        static string ApplicationName = "Cliente CorcoranTest";

        public class DTO
        {
            public string President { get; set; }
            public DateTime? Birthday { get; set; }
            public string Birthplace { get; set; }
            public DateTime? Deathday { get; set; }
            public string Deathplace { get; set; }
        }


        public IEnumerable<DTO> ReadAll()
        {
            var result = new List<DTO>();
            UserCredential credential;
            string newPath = HttpContext.Current.Server.MapPath("~/App_Data/client_secret.json");
            string absolute = Path.GetFullPath(newPath);

            using (var stream =
                new FileStream(absolute, FileMode.Open, FileAccess.Read))
            {
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    
                    "user",
                    CancellationToken.None).Result;
            }

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            // Define request parameters.
            String spreadsheetId = "1i2qbKeasPptIrY1PkFVjbHSrLtKEPIIwES6m2l2Mdd8";
            String range = "A2:E";
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    service.Spreadsheets.Values.Get(spreadsheetId, range);

            // Prints the names and majors of students in a sample spreadsheet:
            // https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
            ValueRange response = request.Execute();
            IList<IList<Object>> values = response.Values;
            if (values != null && values.Count > 0)
            {
                //Console.WriteLine("Name, Major");
                foreach (var row in values)
                {
                    var dto = new DTO()
                    {                        
                        
                        President = row[0].ToString(),
                        Birthplace = row[2].ToString()
                    };
                    DateTime bd; 
                    DateTime.TryParse(row[1].ToString(), out bd);
                    dto.Birthday = bd;

                    if (row.Count > 3)
                    {
                        DateTime dd;
                        DateTime.TryParse(row[3].ToString(), out dd);
                        dto.Deathday = dd;
                        dto.Deathplace = row[4].ToString();

                    }


                    result.Add(dto);

                    // Print columns A and E, which correspond to indices 0 and 4.
                    //Console.WriteLine("{0}, {1}", row[0], row[4]);
                }
            }
            else
            {
                //Console.WriteLine("No data found.");
            }

            return result;
        }

        // GET api/values
        [SwaggerOperation("GetAll")]
        [Route("api/values/GetAll")]
        [HttpGet]
        public IEnumerable<DTO> Get()
        {
            return ReadAll();
        }

        public enum SortOrder
        {
            ASC,
            DES
        }

        [SwaggerOperation("GetAllSorted")]
        [Route("api/values/GetAllSorted")]
        [HttpGet]
        [SwaggerResponse(HttpStatusCode.OK)]
        public IEnumerable<DTO> Get([FromUri] SortOrder birthDaySort = SortOrder.ASC, [FromUri]SortOrder deathDaySort = SortOrder.ASC)
        {
            var result = ReadAll();
            IOrderedEnumerable<DTO> bdOrder;
            if(birthDaySort == SortOrder.ASC)
            {
                bdOrder = result.OrderBy(x => x.Birthday);
            }
            else
            {
                bdOrder = result.OrderByDescending(x => x.Birthday);
            }

            if (birthDaySort == SortOrder.ASC)
            {
                bdOrder = bdOrder.ThenBy(x => ( x.Deathday ?? DateTime.MaxValue));
            }
            else
            {
                bdOrder = bdOrder.ThenBy(x => (x.Deathday ?? DateTime.MinValue));
            }


            return result.ToList();
        }

        [SwaggerOperation("GetFilterByName")]
        [Route("api/values/GetFilterByName")]

        [HttpGet]
        [SwaggerResponse(HttpStatusCode.OK)]
        public IEnumerable<DTO> Get([FromUri] string name)
        {            
            return ReadAll().Where(x=>x.President == name).ToList();
        }

    }
}
