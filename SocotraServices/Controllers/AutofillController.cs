using log4net;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json.Serialization;
using OfficeOpenXml;
using System.Diagnostics;
using System.Net.Http.Headers;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace SocotraServices.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class AutofillController : ControllerBase
    {
        private readonly IConfiguration _configuration;
        private readonly ILog _logger;
        private readonly IHttpClientFactory _httpClientFactory;


        public AutofillController(IConfiguration configuration, ILog logger, IHttpClientFactory httpClientFactory)
        {
            _httpClientFactory = httpClientFactory;
            _configuration = configuration;
            _logger = logger;
        }


        [HttpPost]
        [Route("HHautofill")]
        public async Task<IActionResult> postAutofillAsync(JsonObject requestBody)
        {
            try
            {

                // Access nested properties within "updates" object

                var updates = requestBody["updates"];


                // Access properties within "updateExposures" array

                var updateExposures = updates["updateExposures"];


                var firstExposure = updateExposures[0];
                string exposureLocator = firstExposure["exposureLocator"]?.ToString();
                string exposureName = firstExposure["exposureName"]?.ToString();

                // Access nested properties within "fieldValues" object

                var fieldValues = firstExposure["fieldValues"];

                string city = fieldValues?["City_Town"][0].ToString();
                string address = fieldValues?["address_1"][0].ToString();
                string zipCode = fieldValues?["ZIP_Code"][0].ToString();
                string state = fieldValues?["State"][0].ToString();

                // Genrating Hazardhub Api url

                string apiUrl = $"{_configuration["HazardHubApiBaseUrl"]}?address={address}&city={city}&state={state}&zip={zipCode}";

                // Set up HttpClient

                using (HttpClient httpClient = new HttpClient())
                {
                    string apikey = _configuration["api_key"];

                    // Set authorization header
                    httpClient.DefaultRequestHeaders.Add("Authorization", "Bearer " + apikey);

                    // Send GET request
                    HttpResponseMessage apiResponse = await httpClient.GetAsync(apiUrl);

                    // Check if the response is successful
                    if (apiResponse.IsSuccessStatusCode)
                    {
                        // Read the response content as a string
                        var json = await apiResponse.Content.ReadAsStringAsync();
                        dynamic data = JObject.Parse(json);


                        // Access specific properties

                        string nearestFireStationTitle = data.nearest_fire_station.title;
                        string crimeRate = data.crime.score;
                        string earthquakeScore = data.earthquake.score;


                        // Creating response body
                        var response = new
                        {
                            updateExposures = new[]
                            {
                                new
                                    {
                                        exposureLocator = exposureLocator,
                                        exposureName = exposureName,
                                        fieldValues = new
                                        {
                                            nearest_fire_station = nearestFireStationTitle,
                                            crime_rate = crimeRate,
                                            earthquake_score= earthquakeScore,
                                        }
                                    }
                             }
                        };


                        return Ok(response);
                    }
                    else
                    {
                        // Handle the case where the API request was not successful
                        return StatusCode((int)apiResponse.StatusCode, new { error = "API request failed" });
                    }
                }
            }
            catch (Exception ex)
            {
                // 5. Handle exceptions
                return StatusCode(500, new { error = "An error occurred while processing the request" });
            }
        }

        [HttpPost]
        [Route("excelAutofill")]
        public async Task<IActionResult> AddExposures(JsonObject requestBody)
        {
            try
            {

                var mediaLocator = requestBody["updates"]["fieldValues"]["excel_sheet"][0].ToString();

                var httpClient = _httpClientFactory.CreateClient();

                // Creating Authorization token
                var authRequest = new
                {
                    username = "alice.lee",
                    password = _configuration["tenant_password"],
                    hostName = _configuration["hostname"],
                };

                var authResponse = await httpClient.PostAsJsonAsync(_configuration["socotraBaseUrl"] + "/account/authenticate", authRequest);

                if (!authResponse.IsSuccessStatusCode)
                {
                    return StatusCode(500, "Failed to authenticate with Socotra API");
                }

                var authResult = await authResponse.Content.ReadAsStringAsync();
                dynamic authToken = JObject.Parse(authResult);

                string authorizationToken = authToken.authorizationToken.ToString();

                httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", authorizationToken);

                // Fetch media from mediaLocator
                var mediaResponse = await httpClient.GetAsync($"{_configuration["socotraBaseUrl"]}/media/{mediaLocator}");

                if (!mediaResponse.IsSuccessStatusCode)
                {
                    return StatusCode(500, "Failed to fetch policy data from Socotra API");
                }


                // Fetching neccessary data
                var mediaData = await mediaResponse.Content.ReadAsStringAsync();
                dynamic media_data = JObject.Parse(mediaData);




                string excelFileUrl = media_data.url.ToString();

                // Send an HTTP request to download the file
                HttpResponseMessage excel_response = await httpClient.GetAsync(excelFileUrl);

                if (excel_response.IsSuccessStatusCode)
                {
                    Stream contentStream = await excel_response.Content.ReadAsStreamAsync();

                    List<object> expList = new List<object>();
                    List<string> exposure_delete_list = new List<string>();


                    using (var package = new ExcelPackage(contentStream))
                    {

                        var worksheet = package.Workbook.Worksheets["Worksheet"]; // Replace "Worksheet" with your sheet name.

                        var sheet1 = package.Workbook.Worksheets["Sheet1"];

                        int rowCount = worksheet.Dimension.Rows;

                        List<string> loc_nos = new List<string>();
                        List<string> excel_loc_nos = new List<string>();

                        try
                        {
                            var policyLocator = requestBody["policyLocator"].ToString();

                            var policyResponse = await httpClient.GetAsync($"{_configuration["socotraBaseUrl"]}/policy/{policyLocator}");

                            if (!policyResponse.IsSuccessStatusCode)
                            {
                                return StatusCode(500, "Failed to fetch policy data from Socotra API");
                            }


                            // Fetching neccessary data
                            var policyData = await policyResponse.Content.ReadAsStringAsync();
                            dynamic policy_data = JObject.Parse(policyData);

                            var exposure_count = policy_data.exposures.Count;

                            for (int i = 0; i < exposure_count; i++)
                            {
                                loc_nos.Add(policy_data.exposures[i].characteristics[0]?.fieldValues?.location_number[0].ToString());
                            }
                            for (int y = 5; y <= rowCount; y++)
                            {
                                if (worksheet.Cells[package.Workbook.Names[sheet1.Cells[16, 6]?.Value.ToString() + $"{y - 4}"].Address].Value != null)
                                {
                                    excel_loc_nos.Add(worksheet.Cells[package.Workbook.Names[sheet1.Cells[16, 6]?.Value.ToString() + $"{y - 4}"].Address].Value.ToString());

                                }
                            }

                            List<string> ui_loc = loc_nos.Except(excel_loc_nos).ToList();

                            Console.WriteLine(ui_loc);



                            for (int j = 0; j < exposure_count; j++)
                            {
                                var loc_no = policy_data.exposures[j].characteristics[0]?.fieldValues?.location_number[0].ToString();

                                for (int i = 0; i < ui_loc.Count; i++)
                                {
                                    if (ui_loc[i] == loc_no)
                                    {

                                        exposure_delete_list.Add(policy_data.exposures[j].locator.ToString());
                                    }
                                }
                            }

                        }
                        catch (Exception ex)
                        {
                            Console.Error.WriteLine("Error: " + ex.Message);
                            return StatusCode(500, "An error occurred while processing the request.  " + ex.Message);
                        }



                        var stopwatch = Stopwatch.StartNew();

                        



                        for (int row = 5; row <= rowCount; row++)
                        {
                            if (worksheet.Cells[package.Workbook.Names[sheet1.Cells[16, 6]?.Value.ToString() + $"{row - 4}"].Address].Value != null)
                            {
                                if (!loc_nos.Contains(worksheet.Cells[package.Workbook.Names[sheet1.Cells[16, 6]?.Value.ToString() + $"{row - 4}"].Address].Value.ToString()))
                                {


                                    List<object> perilList = new List<object>();

                                    var table = sheet1.Tables["perils"];

                                    if (table != null)
                                    {

                                        // Access table data
                                        var tableData = table.WorkSheet.Cells[table.Address.Start.Row + 1, table.Address.Start.Column, table.Address.End.Row, table.Address.End.Column];

                                        // Loop through the table data
                                        foreach (var cell in tableData)
                                        {
                                            // Access cell values
                                            string cellValue = cell.Text;


                                            var peril_table = sheet1.Tables[cellValue];

                                            if (peril_table != null)
                                            {
                                                // Access table properties

                                                int peril_table_row_count = peril_table.Address.Rows;


                                                // Access table data
                                                var perilTableData = peril_table.WorkSheet.Cells[peril_table.Address.Start.Row + 1, peril_table.Address.Start.Column, peril_table.Address.End.Row, peril_table.Address.End.Column];

                                                Dictionary<string, string> peril = new Dictionary<string, string>();

                                                if (Convert.ToInt32(worksheet.Cells[package.Workbook.Names[perilTableData[peril_table.Address.Start.Row + 2, 3].Text + $"{row - 4}"].Address].Value.ToString()) != 0)
                                                {
                                                    // Loop through the table data
                                                    for (int row2 = perilTableData.Start.Row; row2 <= peril_table.Address.End.Row; row2++)
                                                    {

                                                        peril[perilTableData[row2, 2].Text] = worksheet.Cells[package.Workbook.Names[perilTableData[row2, 3].Text + $"{row - 4}"].Address].Value.ToString();


                                                    }

                                                    var peril_object = new
                                                    {
                                                        name = perilTableData[peril_table.Address.Start.Row + 1, 3].Text,
                                                        fieldValues = peril
                                                    };

                                                    perilList.Add(peril_object);
                                                }
                                            }
                                            else
                                            {
                                                // Table not found
                                                Console.WriteLine("Table not found in the worksheet.");
                                            }
                                        }
                                    }
                                    else
                                    {
                                        // Table not found
                                        Console.WriteLine("Table not found in the worksheet.");
                                    }

                                    object[] perilArray = perilList.ToArray();

                                    var exposure_table = sheet1.Tables["exposure1"];

                                    if (exposure_table != null)
                                    {

                                        // Access table data
                                        var exposureTableData = exposure_table.WorkSheet.Cells[exposure_table.Address.Start.Row + 1, exposure_table.Address.Start.Column, exposure_table.Address.End.Row, exposure_table.Address.End.Column];

                                        Dictionary<string, string> exposure_fields = new Dictionary<string, string>();


                                        // Loop through the table data
                                        for (int row2 = exposureTableData.Start.Row; row2 <= exposure_table.Address.End.Row; row2++)
                                        {

                                            exposure_fields[exposureTableData[row2, 5].Text] = worksheet.Cells[package.Workbook.Names[exposureTableData[row2, 6].Text + $"{row - 4}"].Address].Value.ToString();


                                        }

                                        var exposureFieldValues = exposure_fields;


                                        var exposure = new
                                        {
                                            exposureName = sheet1.Cells[3, 6]?.Value.ToString(),
                                            fieldValues = exposureFieldValues,
                                            perils = perilArray

                                        };

                                        expList.Add(exposure);


                                    }
                                    else
                                    {
                                        // Table not found
                                        Console.WriteLine("Table not found");
                                    }
                                }
                            }
                        }

                        stopwatch.Stop();
                        // Calculate and log the elapsed time
                        var elapsedMilliseconds = stopwatch.ElapsedMilliseconds;
                        _logger.Info($"Execution time of the for loop: {elapsedMilliseconds} ms");


                        object[] exposureArray = expList.ToArray();

                        var policy_table = sheet1.Tables["policy"];

                        if (policy_table != null)
                        {

                            // Access table data
                            var policyTableData = policy_table.WorkSheet.Cells[policy_table.Address.Start.Row + 1, policy_table.Address.Start.Column, policy_table.Address.End.Row, policy_table.Address.End.Column];

                            Dictionary<string, string> policy_fields = new Dictionary<string, string>();


                            // Loop through the table data
                            for (int row2 = policyTableData.Start.Row; row2 <= policy_table.Address.End.Row; row2++)
                            {

                                policy_fields[policyTableData[row2, 5].Text] = policyTableData[row2, 6].Value.ToString();


                            }

                            var fieldValues = policy_fields;

                            string[] remove_exposure_array = exposure_delete_list.ToArray();

                            var response = new
                            {
                                fieldValues = fieldValues,
                                addExposures = exposureArray,
                                removeExposures = remove_exposure_array
                            };

                            var settings = new JsonSerializerSettings
                            {
                                ContractResolver = new DefaultContractResolver
                                {
                                    NamingStrategy = new DefaultNamingStrategy()
                                }
                            };

                            var jsonResponse = JsonConvert.SerializeObject(response, settings);

                            return new ContentResult
                            {
                                Content = jsonResponse,
                                ContentType = "application/json",
                                StatusCode = 200
                            };
                        }
                        else
                        {
                            // Table not found


                            return StatusCode(500, new
                            {
                                error = "policy table not found"
                            });
                        }



                    }
                }
                else
                {
                    return StatusCode(500, new
                    {
                        error = "excel not found"
                    });
                }
            }
            catch (Exception ex)
            {
                return StatusCode(500, new
                {
                    error = "An error occurred while processing the request" + ex.Message
                });
            }
        }

        [HttpPost]
        [Route("excelAutofill_new")]
        public async Task<IActionResult> AddExposures_new(JsonObject requestBody)
        {
            try
            {

                var mediaLocator = requestBody["updates"]["fieldValues"]["excel_sheet"][0].ToString();

                var httpClient = _httpClientFactory.CreateClient();

                // Creating Authorization token
                var authRequest = new
                {
                    username = "alice.lee",
                    password = _configuration["tenant_password"],
                    hostName = _configuration["hostname"],
                };

                var authResponse = await httpClient.PostAsJsonAsync(_configuration["socotraBaseUrl"] + "/account/authenticate", authRequest);

                if (!authResponse.IsSuccessStatusCode)
                {
                    return StatusCode(500, "Failed to authenticate with Socotra API");
                }

                var authResult = await authResponse.Content.ReadAsStringAsync();
                dynamic authToken = JObject.Parse(authResult);

                string authorizationToken = authToken.authorizationToken.ToString();

                httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", authorizationToken);

                // Fetch media from mediaLocator
                var mediaResponse = await httpClient.GetAsync($"{_configuration["socotraBaseUrl"]}/media/{mediaLocator}");

                if (!mediaResponse.IsSuccessStatusCode)
                {
                    return StatusCode(500, "Failed to fetch policy data from Socotra API");
                }


                // Fetching neccessary data
                var mediaData = await mediaResponse.Content.ReadAsStringAsync();
                dynamic media_data = JObject.Parse(mediaData);




                string excelFileUrl = media_data.url.ToString();

                // Send an HTTP request to download the file
                HttpResponseMessage excel_response = await httpClient.GetAsync(excelFileUrl);

                if (excel_response.IsSuccessStatusCode)
                {
                    Stream contentStream = await excel_response.Content.ReadAsStreamAsync();

                    List<object> expList = new List<object>();
                    List<string> exposure_delete_list = new List<string>();


                    using (var package = new ExcelPackage(contentStream))
                    {

                        var worksheet = package.Workbook.Worksheets["Worksheet"]; // Replace "Worksheet" with your sheet name.

                        var sheet1 = package.Workbook.Worksheets["Sheet1"];

                        var exposure_table = sheet1.Tables["exposure1"];

                        var table = sheet1.Tables["perils"];

                        int rowCount = worksheet.Dimension.Rows;

                        List<string> policy_loc_nos = new List<string>();

                        List<string> excel_loc_nos = new List<string>();

                        try
                        {
                            var policyLocator = requestBody["policyLocator"].ToString();

                            var policyResponse = await httpClient.GetAsync($"{_configuration["socotraBaseUrl"]}/policy/{policyLocator}");

                            if (!policyResponse.IsSuccessStatusCode)
                            {
                                return StatusCode(500, "Failed to fetch policy data from Socotra API");
                            }


                            // Fetching neccessary data
                            var policyData = await policyResponse.Content.ReadAsStringAsync();
                            dynamic policy_data = JObject.Parse(policyData);

                            var exposure_count = policy_data.exposures.Count;


                            for (int i = 0; i < exposure_count; i++)
                            {
                                //policy_loc_nos.Add(policy_data.exposures[i].characteristics[0]?.fieldValues?.location_number[0].ToString());
                                exposure_delete_list.Add(policy_data.exposures[i].locator.ToString());
                            }

                            /*
                            for (int y = 5; y <= rowCount; y++)
                            {
                                if (worksheet.Cells[package.Workbook.Names[sheet1.Cells[16, 6]?.Value.ToString() + $"{y - 4}"].Address].Value != null)
                                {
                                    excel_loc_nos.Add(worksheet.Cells[package.Workbook.Names[sheet1.Cells[16, 6]?.Value.ToString() + $"{y - 4}"].Address].Value.ToString());

                                }
                            }

                            List<string> ui_loc = policy_loc_nos.Except(excel_loc_nos).ToList();

                            Console.WriteLine(ui_loc);

                            
                            
                            for (int j = 0; j < exposure_count; j++)
                            {
                                var loc_no = policy_data.exposures[j].characteristics[0]?.fieldValues?.location_number[0].ToString();

                                for(int i = 0; i < ui_loc.Count; i++)
                                {
                                    if (ui_loc[i] == loc_no)
                                    {

                                        exposure_delete_list.Add(policy_data.exposures[j].locator.ToString());
                                    }
                                }
                            }

                            */


                            /*
                            for (int i = 0; i < ui_loc.Count; i++)
                            {
                                for (int j = 0; j < exposure_count; j++)
                                {
                                    var peril_count = policy_data.exposures[j].perils.Count;

                                    int empty_row=0;

                                    for (int row = 5; row <= rowCount; row++)
                                    {
                                        if (worksheet.Cells[package.Workbook.Names[sheet1.Cells[16, 6]?.Value.ToString() + $"{row - 4}"].Address].Value == null)
                                        {    // Loop through the table data
                                            empty_row = row;
                                            break;
                                        }
                                    }
                                        if (policy_data.exposures[j].characteristics[0]?.fieldValues?.location_number[0].ToString() == ui_loc[i])
                                        {
                                            if (exposure_table != null)
                                            {

                                                // Access table data
                                                var exposureTableData = exposure_table.WorkSheet.Cells[exposure_table.Address.Start.Row + 1, exposure_table.Address.Start.Column, exposure_table.Address.End.Row, exposure_table.Address.End.Column];


                                                if (worksheet.Cells[package.Workbook.Names[sheet1.Cells[16, 6]?.Value.ToString() + $"{empty_row - 4}"].Address].Value == null)
                                                {    // Loop through the table data
                                                    for (int row2 = exposureTableData.Start.Row; row2 <= exposure_table.Address.End.Row; row2++)
                                                    {

                                                        worksheet.Cells[package.Workbook.Names[exposureTableData[row2, 6].Text + $"{empty_row - 4}"].Address].Value = policy_data.exposures[j].characteristics[0]?.fieldValues?[exposureTableData[row2, 5].Text];


                                                    }
                                                }
                                            }
                                        }
                                        else
                                        {
                                            // Table not found
                                            Console.WriteLine("Table not found");
                                        }


                                        if (table != null)
                                        {

                                            // Access table data
                                            var tableData = table.WorkSheet.Cells[table.Address.Start.Row + 1, table.Address.Start.Column, table.Address.End.Row, table.Address.End.Column];

                                            // Loop through the table data
                                            foreach (var cell in tableData)
                                            {
                                                // Access cell values
                                                string cellValue = cell.Text;


                                                var peril_table = sheet1.Tables[cellValue];

                                                if (peril_table != null)
                                                {
                                                    // Access table properties


                                                    // Access table data
                                                    var perilTableData = peril_table.WorkSheet.Cells[peril_table.Address.Start.Row + 1, peril_table.Address.Start.Column, peril_table.Address.End.Row, peril_table.Address.End.Column];

                                                    var peril_name = perilTableData[peril_table.Address.Start.Row + 1, 3].Text;

                                                    for (int m = 0; m < peril_count; m++)
                                                    {
                                                        if (peril_name == policy_data.exposures[j].perils[m].name.ToString())
                                                        {
                                                            // Loop through the table data
                                                            for (int row2 = perilTableData.Start.Row+1; row2 <= peril_table.Address.End.Row; row2++)
                                                            {

                                                                worksheet.Cells[package.Workbook.Names[perilTableData[row2, 3].Text + $"{empty_row - 4}"].Address].Value = policy_data.exposures[j].perils[m].characteristics[0]?.fieldValues?[perilTableData[row2, 2].Text];


                                                            }
                                                        }
                                                    }

                                                }
                                                else
                                                {
                                                    // Table not found
                                                    Console.WriteLine("Table not found in the worksheet.");
                                                }
                                            }
                                        }
                                        else
                                        {
                                            // Table not found
                                            Console.WriteLine("Table not found in the worksheet.");
                                        }


                                }

                            }
                            package.Save();
                            */
                        }
                        catch (Exception ex)
                        {
                            Console.Error.WriteLine("Error: " + ex.Message);
                            return StatusCode(500, "1.An error occurred while processing the request.  " + ex.Message);
                        }



                        var stopwatch = Stopwatch.StartNew();





                        for (int row = 5; row <= rowCount; row++)
                        {

                            Console.WriteLine(rowCount);
                            if (worksheet.Cells[package.Workbook.Names[sheet1.Cells[16, 6]?.Value.ToString() + $"{row - 4}"].Address].Value != null)
                            {
                                /*
                                Console.WriteLine(row);
                                if (!policy_loc_nos.Contains(worksheet.Cells[package.Workbook.Names[sheet1.Cells[16, 6]?.Value.ToString() + $"{row - 4}"].Address].Value.ToString()))
                                {
                            */

                                List<object> perilList = new List<object>();



                                if (table != null)
                                {

                                    // Access table data
                                    var tableData = table.WorkSheet.Cells[table.Address.Start.Row + 1, table.Address.Start.Column, table.Address.End.Row, table.Address.End.Column];

                                    // Loop through the table data
                                    foreach (var cell in tableData)
                                    {
                                        // Access cell values
                                        string cellValue = cell.Text;


                                        var peril_table = sheet1.Tables[cellValue];

                                        if (peril_table != null)
                                        {
                                            // Access table properties

                                            int peril_table_row_count = peril_table.Address.Rows;


                                            // Access table data
                                            var perilTableData = peril_table.WorkSheet.Cells[peril_table.Address.Start.Row + 1, peril_table.Address.Start.Column, peril_table.Address.End.Row, peril_table.Address.End.Column];

                                            Dictionary<string, string> peril = new Dictionary<string, string>();

                                            if (Convert.ToInt32(worksheet.Cells[package.Workbook.Names[perilTableData[peril_table.Address.Start.Row + 2, 3].Text + $"{row - 4}"].Address].Value.ToString()) != 0)
                                            {
                                                // Loop through the table data
                                                for (int row2 = perilTableData.Start.Row; row2 <= peril_table.Address.End.Row; row2++)
                                                {

                                                    peril[perilTableData[row2, 2].Text] = worksheet.Cells[package.Workbook.Names[perilTableData[row2, 3].Text + $"{row - 4}"].Address].Value.ToString();


                                                }

                                                var peril_object = new
                                                {
                                                    name = perilTableData[peril_table.Address.Start.Row + 1, 3].Text,
                                                    fieldValues = peril
                                                };

                                                perilList.Add(peril_object);
                                            }
                                        }
                                        else
                                        {
                                            // Table not found
                                            Console.WriteLine("Table not found in the worksheet.");
                                        }
                                    }
                                }
                                else
                                {
                                    // Table not found
                                    Console.WriteLine("Table not found in the worksheet.");
                                }

                                object[] perilArray = perilList.ToArray();



                                if (exposure_table != null)
                                {

                                    // Access table data
                                    var exposureTableData = exposure_table.WorkSheet.Cells[exposure_table.Address.Start.Row + 1, exposure_table.Address.Start.Column, exposure_table.Address.End.Row, exposure_table.Address.End.Column];

                                    Dictionary<string, string> exposure_fields = new Dictionary<string, string>();


                                    // Loop through the table data
                                    for (int row2 = exposureTableData.Start.Row; row2 <= exposure_table.Address.End.Row; row2++)
                                    {

                                        exposure_fields[exposureTableData[row2, 5].Text] = worksheet.Cells[package.Workbook.Names[exposureTableData[row2, 6].Text + $"{row - 4}"].Address].Value.ToString();


                                    }

                                    var exposureFieldValues = exposure_fields;


                                    var exposure = new
                                    {
                                        exposureName = sheet1.Cells[3, 6]?.Value.ToString(),
                                        fieldValues = exposureFieldValues,
                                        perils = perilArray

                                    };

                                    expList.Add(exposure);


                                }
                                else
                                {
                                    // Table not found
                                    Console.WriteLine("Table not found");
                                }

                            }
                        }

                        stopwatch.Stop();
                        // Calculate and log the elapsed time
                        var elapsedMilliseconds = stopwatch.ElapsedMilliseconds;
                        _logger.Info($"Execution time of the for loop: {elapsedMilliseconds} ms");


                        object[] exposureArray = expList.ToArray();

                        var policy_table = sheet1.Tables["policy"];

                        if (policy_table != null)
                        {

                            // Access table data
                            var policyTableData = policy_table.WorkSheet.Cells[policy_table.Address.Start.Row + 1, policy_table.Address.Start.Column, policy_table.Address.End.Row, policy_table.Address.End.Column];

                            Dictionary<string, string> policy_fields = new Dictionary<string, string>();


                            // Loop through the table data
                            for (int row2 = policyTableData.Start.Row; row2 <= policy_table.Address.End.Row; row2++)
                            {

                                policy_fields[policyTableData[row2, 5].Text] = policyTableData[row2, 6].Value.ToString();


                            }

                            var fieldValues = policy_fields;

                            string[] remove_exposure_array = exposure_delete_list.ToArray();

                            var response = new
                            {
                                fieldValues = fieldValues,
                                addExposures = exposureArray,
                                removeExposures = remove_exposure_array
                            };

                            var settings = new JsonSerializerSettings
                            {
                                ContractResolver = new DefaultContractResolver
                                {
                                    NamingStrategy = new DefaultNamingStrategy()
                                }
                            };

                            var jsonResponse = JsonConvert.SerializeObject(response, settings);

                            return new ContentResult
                            {
                                Content = jsonResponse,
                                ContentType = "application/json",
                                StatusCode = 200
                            };
                        }
                        else
                        {
                            // Table not found


                            return StatusCode(500, new
                            {
                                error = "policy table not found"
                            });
                        }



                    }
                }
                else
                {
                    return StatusCode(500, new
                    {
                        error = "excel not found"
                    });
                }
            }
            catch (Exception ex)
            {
                return StatusCode(500, new
                {
                    error = "2.An error occurred while processing the request" + ex.Message
                });
            }
        }

        [HttpPost]
        [Route("excelAutofill_new2")]
        public async Task<IActionResult> AddExposures_new2(JsonObject requestBody)
        {
            try
            {

                var mediaLocator = requestBody["updates"]["fieldValues"]["excel_sheet"][0].ToString();

                var httpClient = _httpClientFactory.CreateClient();

                // Creating Authorization token
                var authRequest = new
                {
                    username = "alice.lee",
                    password = _configuration["tenant_password"],
                    hostName = _configuration["hostname"],
                };

                var authResponse = await httpClient.PostAsJsonAsync(_configuration["socotraBaseUrl"] + "/account/authenticate", authRequest);

                if (!authResponse.IsSuccessStatusCode)
                {
                    return StatusCode(500, "Failed to authenticate with Socotra API");
                }

                var authResult = await authResponse.Content.ReadAsStringAsync();
                dynamic authToken = JObject.Parse(authResult);

                string authorizationToken = authToken.authorizationToken.ToString();

                httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", authorizationToken);

                // Fetch media from mediaLocator
                var mediaResponse = await httpClient.GetAsync($"{_configuration["socotraBaseUrl"]}/media/{mediaLocator}");

                if (!mediaResponse.IsSuccessStatusCode)
                {
                    return StatusCode(500, "Failed to fetch policy data from Socotra API");
                }


                // Fetching neccessary data
                var mediaData = await mediaResponse.Content.ReadAsStringAsync();
                dynamic media_data = JObject.Parse(mediaData);




                string excelFileUrl = media_data.url.ToString();

                // Send an HTTP request to download the file
                HttpResponseMessage excel_response = await httpClient.GetAsync(excelFileUrl);

                if (excel_response.IsSuccessStatusCode)
                {
                    Stream contentStream = await excel_response.Content.ReadAsStreamAsync();

                    List<object> expList = new List<object>();
                    List<string> exposure_delete_list = new List<string>();


                    using (var package = new ExcelPackage(contentStream))
                    {

                        var worksheet = package.Workbook.Worksheets["Worksheet"]; // Replace "Worksheet" with your sheet name.

                        var sheet1 = package.Workbook.Worksheets["Sheet1"];

                        var exposure_table = sheet1.Tables["exposure1"];

                        var table = sheet1.Tables["perils"];

                        int rowCount = worksheet.Dimension.Rows;

                        List<string> policy_loc_nos = new List<string>();

                        List<string> excel_loc_nos = new List<string>();

                        try
                        {
                            var policyLocator = requestBody["policyLocator"].ToString();

                            var policyResponse = await httpClient.GetAsync($"{_configuration["socotraBaseUrl"]}/policy/{policyLocator}");

                            if (!policyResponse.IsSuccessStatusCode)
                            {
                                return StatusCode(500, "Failed to fetch policy data from Socotra API");
                            }


                            // Fetching neccessary data
                            var policyData = await policyResponse.Content.ReadAsStringAsync();
                            dynamic policy_data = JObject.Parse(policyData);

                            var exposure_count = policy_data.exposures.Count;


                            for (int i = 0; i < exposure_count; i++)
                            {
                                policy_loc_nos.Add(policy_data.exposures[i].characteristics[0]?.fieldValues?.location_number[0].ToString());
                                exposure_delete_list.Add(policy_data.exposures[i].locator.ToString());
                            }


                            for (int y = 5; y <= rowCount; y++)
                            {
                                if (worksheet.Cells[package.Workbook.Names[sheet1.Cells[16, 6]?.Value.ToString() + $"{y - 4}"].Address].Value != null)
                                {


                                    excel_loc_nos.Add(worksheet.Cells[package.Workbook.Names[sheet1.Cells[16, 6]?.Value.ToString() + $"{y - 4}"].Address].Value.ToString());

                                }
                            }

                            List<string> ui_loc = policy_loc_nos.Except(excel_loc_nos).ToList();

                            Console.WriteLine(ui_loc);




                            // Adding the excel by adding those exposures that are in UI but not in excel.

                            for (int i = 0; i < ui_loc.Count; i++)
                            {
                                for (int j = 0; j < exposure_count; j++)
                                {
                                    var peril_count = policy_data.exposures[j].perils.Count;

                                    int empty_row = 0;

                                    for (int row = 5; row <= rowCount; row++)
                                    {
                                        if (worksheet.Cells[package.Workbook.Names[sheet1.Cells[16, 6]?.Value.ToString() + $"{row - 4}"].Address].Value == null)
                                        {    // Loop through the table data
                                            empty_row = row;
                                            break;
                                        }
                                    }
                                    if (policy_data.exposures[j].characteristics[0]?.fieldValues?.location_number[0].ToString() == ui_loc[i])
                                    {
                                        if (exposure_table != null)
                                        {

                                            // Access table data
                                            var exposureTableData = exposure_table.WorkSheet.Cells[exposure_table.Address.Start.Row + 1, exposure_table.Address.Start.Column, exposure_table.Address.End.Row, exposure_table.Address.End.Column];


                                            if (worksheet.Cells[package.Workbook.Names[sheet1.Cells[16, 6]?.Value.ToString() + $"{empty_row - 4}"].Address].Value == null)
                                            {    // Loop through the table data
                                                for (int row2 = exposureTableData.Start.Row; row2 <= exposure_table.Address.End.Row; row2++)
                                                {
                                                    var existingFormat = worksheet.Cells[package.Workbook.Names[exposureTableData[row2, 6].Text + $"{empty_row - 4}"].Address].Style.Numberformat.Format;

                                                    worksheet.Cells[package.Workbook.Names[exposureTableData[row2, 6].Text + $"{empty_row - 4}"].Address].Value = policy_data.exposures[j].characteristics[0]?.fieldValues?[exposureTableData[row2, 5].Text][0].ToString();

                                                    worksheet.Cells[package.Workbook.Names[exposureTableData[row2, 6].Text + $"{empty_row - 4}"].Address].Style.Numberformat.Format = existingFormat;
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        // Table not found
                                        Console.WriteLine("Table not found");
                                    }


                                    if (table != null)
                                    {

                                        // Access table data
                                        var tableData = table.WorkSheet.Cells[table.Address.Start.Row + 1, table.Address.Start.Column, table.Address.End.Row, table.Address.End.Column];

                                        // Loop through the table data
                                        foreach (var cell in tableData)
                                        {
                                            // Access cell values
                                            string cellValue = cell.Text;


                                            var peril_table = sheet1.Tables[cellValue];

                                            if (peril_table != null)
                                            {
                                                // Access table properties


                                                // Access table data
                                                var perilTableData = peril_table.WorkSheet.Cells[peril_table.Address.Start.Row + 1, peril_table.Address.Start.Column, peril_table.Address.End.Row, peril_table.Address.End.Column];

                                                var peril_name = perilTableData[peril_table.Address.Start.Row + 1, 3].Text;

                                                for (int m = 0; m < peril_count; m++)
                                                {
                                                    if (peril_name == policy_data.exposures[j].perils[m].name.ToString())
                                                    {
                                                        // Loop through the table data
                                                        for (int row2 = perilTableData.Start.Row + 1; row2 <= peril_table.Address.End.Row; row2++)
                                                        {
                                                            var existingFormat = worksheet.Cells[package.Workbook.Names[perilTableData[row2, 3].Text + $"{empty_row - 4}"].Address].Style.Numberformat.Format;

                                                            worksheet.Cells[package.Workbook.Names[perilTableData[row2, 3].Text + $"{empty_row - 4}"].Address].Value = policy_data.exposures[j].perils[m].characteristics[0]?.fieldValues?[perilTableData[row2, 2].Text][0].ToString();

                                                            worksheet.Cells[package.Workbook.Names[perilTableData[row2, 3].Text + $"{empty_row - 4}"].Address].Style.Numberformat.Format = existingFormat;
                                                        }
                                                    }
                                                }

                                            }
                                            else
                                            {
                                                // Table not found
                                                Console.WriteLine("Peril Table not found in the worksheet.");
                                            }
                                        }

                                    }
                                    else
                                    {
                                        // Table not found
                                        Console.WriteLine("Exposure Table not found in the worksheet.");
                                    }
                                }
                            }
                            package.Save();

                        }
                        catch (Exception ex)
                        {
                            Console.Error.WriteLine("Error: " + ex.Message);
                            return StatusCode(500, "1.An error occurred while processing the request. : " + ex.Message);
                        }



                        var stopwatch = Stopwatch.StartNew();



                        // creating exposure object from excel rows.

                        for (int row = 5; row <= rowCount; row++)
                        {

                            Console.WriteLine(rowCount);
                            if (worksheet.Cells[package.Workbook.Names[sheet1.Cells[16, 6]?.Value.ToString() + $"{row - 4}"].Address].Value != null)
                            {
                                /*
                                Console.WriteLine(row);
                                if (!policy_loc_nos.Contains(worksheet.Cells[package.Workbook.Names[sheet1.Cells[16, 6]?.Value.ToString() + $"{row - 4}"].Address].Value.ToString()))
                                {
                            */

                                List<object> perilList = new List<object>();



                                if (table != null)
                                {

                                    // Access table data
                                    var tableData = table.WorkSheet.Cells[table.Address.Start.Row + 1, table.Address.Start.Column, table.Address.End.Row, table.Address.End.Column];

                                    // Loop through the table data
                                    foreach (var cell in tableData)
                                    {
                                        // Access cell values
                                        string cellValue = cell.Text;


                                        var peril_table = sheet1.Tables[cellValue];

                                        if (peril_table != null)
                                        {
                                            // Access table properties

                                            int peril_table_row_count = peril_table.Address.Rows;


                                            // Access table data
                                            var perilTableData = peril_table.WorkSheet.Cells[peril_table.Address.Start.Row + 1, peril_table.Address.Start.Column, peril_table.Address.End.Row, peril_table.Address.End.Column];

                                            Dictionary<string, string> peril = new Dictionary<string, string>();

                                            Console.WriteLine(worksheet.Cells[package.Workbook.Names[perilTableData[peril_table.Address.Start.Row + 2, 3].Text + $"{row - 4}"].Address].Value.ToString());

                                            if (Convert.ToInt32(worksheet.Cells[package.Workbook.Names[perilTableData[peril_table.Address.Start.Row + 2, 3].Text + $"{row - 4}"].Address].Value.ToString()) != 0)
                                            {
                                                // Loop through the table data
                                                for (int row2 = perilTableData.Start.Row; row2 <= peril_table.Address.End.Row; row2++)
                                                {

                                                    peril[perilTableData[row2, 2].Text] = worksheet.Cells[package.Workbook.Names[perilTableData[row2, 3].Text + $"{row - 4}"].Address].Value.ToString();


                                                }

                                                var peril_object = new
                                                {
                                                    name = perilTableData[peril_table.Address.Start.Row + 1, 3].Text,
                                                    fieldValues = peril
                                                };

                                                perilList.Add(peril_object);
                                            }


                                        }
                                        else
                                        {
                                            // Table not found
                                            Console.WriteLine("Table not found in the worksheet.");
                                        }
                                    }
                                }
                                else
                                {
                                    // Table not found
                                    Console.WriteLine("Table not found in the worksheet.");
                                }

                                object[] perilArray = perilList.ToArray();



                                if (exposure_table != null)
                                {

                                    // Access table data
                                    var exposureTableData = exposure_table.WorkSheet.Cells[exposure_table.Address.Start.Row + 1, exposure_table.Address.Start.Column, exposure_table.Address.End.Row, exposure_table.Address.End.Column];

                                    Dictionary<string, string> exposure_fields = new Dictionary<string, string>();


                                    // Loop through the table data
                                    for (int row2 = exposureTableData.Start.Row; row2 <= exposure_table.Address.End.Row; row2++)
                                    {

                                        exposure_fields[exposureTableData[row2, 5].Text] = worksheet.Cells[package.Workbook.Names[exposureTableData[row2, 6].Text + $"{row - 4}"].Address].Value.ToString();


                                    }

                                    var exposureFieldValues = exposure_fields;


                                    var exposure = new
                                    {
                                        exposureName = sheet1.Cells[3, 6]?.Value.ToString(),
                                        fieldValues = exposureFieldValues,
                                        perils = perilArray

                                    };

                                    expList.Add(exposure);


                                }
                                else
                                {
                                    // Table not found
                                    Console.WriteLine("Table not found");
                                }

                            }
                        }

                        stopwatch.Stop();
                        // Calculate and log the elapsed time
                        var elapsedMilliseconds = stopwatch.ElapsedMilliseconds;
                        _logger.Info($"Execution time of the for loop: {elapsedMilliseconds} ms");


                        object[] exposureArray = expList.ToArray();

                        var policy_table = sheet1.Tables["policy"];

                        if (policy_table != null)
                        {

                            // Access table data
                            var policyTableData = policy_table.WorkSheet.Cells[policy_table.Address.Start.Row + 1, policy_table.Address.Start.Column, policy_table.Address.End.Row, policy_table.Address.End.Column];

                            Dictionary<string, string> policy_fields = new Dictionary<string, string>();


                            // Loop through the table data
                            for (int row2 = policyTableData.Start.Row; row2 <= policy_table.Address.End.Row; row2++)
                            {

                                policy_fields[policyTableData[row2, 5].Text] = policyTableData[row2, 6].Value.ToString();


                            }

                            var fieldValues = policy_fields;

                            string[] remove_exposure_array = exposure_delete_list.ToArray();

                            var response = new
                            {
                                fieldValues = fieldValues,
                                addExposures = exposureArray,
                                removeExposures = remove_exposure_array
                            };

                            var settings = new JsonSerializerSettings
                            {
                                ContractResolver = new DefaultContractResolver
                                {
                                    NamingStrategy = new DefaultNamingStrategy()
                                }
                            };

                            var jsonResponse = JsonConvert.SerializeObject(response, settings);

                            return new ContentResult
                            {
                                Content = jsonResponse,
                                ContentType = "application/json",
                                StatusCode = 200
                            };
                        }
                        else
                        {
                            // Table not found


                            return StatusCode(500, new
                            {
                                error = "policy table not found"
                            });
                        }




                    }
                }
                else
                {
                    return StatusCode(500, new
                    {
                        error = "excel not found"
                    });
                }
            }
            catch (Exception ex)
            {
                return StatusCode(500, new
                {
                    error = "2.An error occurred while processing the request.",
                    message = ex.Message,
                    stackTrace = ex.StackTrace
                });
            }
        }

        [HttpPost]
        [Route("renewalInflation")]
        public async Task<IActionResult> renewalInflation(JsonObject requestBody)
        {

            try
            {
                var policyLocator = requestBody["policyLocator"].ToString();

                var httpClient = _httpClientFactory.CreateClient();

                // Creating Authorization token
                var authRequest = new
                {
                    username = "alice.lee",
                    password = _configuration["tenant_password"],
                    hostName = _configuration["hostname"],
                };

                var authResponse = await httpClient.PostAsJsonAsync(_configuration["socotraBaseUrl"] + "/account/authenticate", authRequest);

                if (!authResponse.IsSuccessStatusCode)
                {
                    return StatusCode(500, "Failed to authenticate with Socotra API");
                }

                var authResult = await authResponse.Content.ReadAsStringAsync();
                dynamic authToken = JObject.Parse(authResult);

                string authorizationToken = authToken.authorizationToken.ToString();

                httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", authorizationToken);

                // Fetch policy from policyLocator
                var policyResponse = await httpClient.GetAsync($"{_configuration["socotraBaseUrl"]}/policy/{policyLocator}");

                if (!policyResponse.IsSuccessStatusCode)
                {
                    return StatusCode(500, "Failed to fetch policy data from Socotra API");
                }


                // Fetching neccessary data
                var policyData = await policyResponse.Content.ReadAsStringAsync();
                var policy_data = (JsonObject)JsonObject.Parse(policyData);

                int exposureCount = 0;

                var exposureArray = policy_data["exposures"];

                if (exposureArray is JsonArray arr)
                {
                    exposureCount = arr.Count;
                }

                List<object> expList = new List<object>();

                for (int i = 0; i < exposureCount; i++)
                {
                    if (exposureArray[i]["name"].ToString() == "Commercial Building")
                    {
                        var exposureLocator = exposureArray[i]["locator"].ToString();

                        var fireAndPerils = exposureArray[i]["perils"][1]?["characteristics"][0]["fieldValues"];

                        var Fire_Limit_Building = Convert.ToDouble(fireAndPerils["Fire_Limit_Building"][0].ToString());
                        var Fire_Limit_Contents = Convert.ToDouble(fireAndPerils["Fire_Limit_Contents"][0].ToString());
                        var Stock_Limit = Convert.ToDouble(fireAndPerils["Stock_Limit"][0].ToString());
                        var Machinery_Or_Equipment_Breakdown_Limit = Convert.ToDouble(fireAndPerils["Machinery_Or_Equipment_Breakdown_Limit"][0].ToString());
                        var Electronic_Breakdown_Limit = Convert.ToDouble(fireAndPerils["Electronic_Breakdown_Limit"][0].ToString());
                        var Business_Interruption_or_Income_Limit = Convert.ToDouble(fireAndPerils["Business_Interruption_or_Income_Limit"][0].ToString());
                        var Personal_Property_of_others_Limit = Convert.ToDouble(fireAndPerils["Personal_Property_of_others_Limit"][0].ToString());

                        var TotalLimit = Fire_Limit_Building + Fire_Limit_Contents + Stock_Limit + Machinery_Or_Equipment_Breakdown_Limit + Electronic_Breakdown_Limit + Business_Interruption_or_Income_Limit + Personal_Property_of_others_Limit;

                        if (TotalLimit < 25000000)
                        {
                            Fire_Limit_Building = Fire_Limit_Building * 1.05;
                            Fire_Limit_Contents = Fire_Limit_Contents * 1.05;
                            Stock_Limit = Stock_Limit * 1.05;
                            Machinery_Or_Equipment_Breakdown_Limit = Machinery_Or_Equipment_Breakdown_Limit * 1.05;
                            Electronic_Breakdown_Limit = Electronic_Breakdown_Limit * 1.05;
                            Business_Interruption_or_Income_Limit = Business_Interruption_or_Income_Limit * 1.05;
                            Personal_Property_of_others_Limit = Personal_Property_of_others_Limit * 1.05;
                        }

                        var updatedPeril = new
                        {
                            perilLocator = exposureArray[i]["perils"][1]["locator"],
                            fieldValues = new
                            {
                                Fire_Limit_Building = Fire_Limit_Building,
                                Fire_Limit_Contents = Fire_Limit_Contents,
                                Stock_Limit = Stock_Limit,
                                Machinery_Or_Equipment_Breakdown_Limit = Machinery_Or_Equipment_Breakdown_Limit,
                                Electronic_Breakdown_Limit = Electronic_Breakdown_Limit,
                                Business_Interruption_or_Income_Limit = Business_Interruption_or_Income_Limit,
                                Personal_Property_of_others_Limit = Personal_Property_of_others_Limit,
                            }
                        };

                        var exposure = new
                        {
                            exposureName = "Commercial Building",
                            exposureLocator = exposureLocator,
                            perils = updatedPeril

                        };

                        expList.Add(exposure);
                    }
                }

                object[] exposureArray2 = expList.ToArray();

                var response = new
                {
                    updateExposures = exposureArray2
                };


                var jsonOptions = new JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true, // This corresponds to DefaultNamingStrategy in Newtonsoft.Json
                    WriteIndented = false // You can set this to true for pretty-printed JSON
                };

                var jsonResponse = System.Text.Json.JsonSerializer.Serialize(response, jsonOptions);


                return new ContentResult
                {
                    Content = jsonResponse,
                    ContentType = "application/json",
                    StatusCode = 200
                };

                return Ok(response);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("Error: " + ex.Message);
                return StatusCode(500, "An error occurred while processing the request.  " + ex.Message);
            }
        }
    }
}
