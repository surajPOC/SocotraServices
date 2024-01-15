using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using System.Net.Http.Headers;
using System.Text.Json.Nodes;

namespace SocotraServices.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExternalRaterController : ControllerBase
    {
        private readonly IConfiguration _configuration;
        private readonly IHttpClientFactory _httpClientFactory;


        public ExternalRaterController(IConfiguration configuration, IHttpClientFactory httpClientFactory)
        {
            _httpClientFactory = httpClientFactory;
            _configuration = configuration;

        }

        [HttpPost]
        [Route("excelRater")]
        public async Task<IActionResult> excelRater(JsonObject requestBody)
        {
            try
            {
                var policyExposurePerils = requestBody["policyExposurePerils"];

                var perilPremium = new Dictionary<string, object>();

                using (var package = new ExcelPackage(new FileInfo("C:\\Users\\SURAJ KUMAR\\Desktop\\SocotraServices\\main_excelRater.xlsx")))
                {
                    var worksheet = package.Workbook.Worksheets["Worksheet"]; // Replace "Worksheet" with your sheet name.

                    int rowCount = 5;

                    int perilsCount = 0;

                    if (requestBody.ContainsKey("policyExposurePerils") && requestBody["policyExposurePerils"] is JsonArray arr)
                    {
                        perilsCount = arr.Count;
                        // Now, 'length' contains the number of elements in the array
                    }

                    for (int i = 0; i < perilsCount; i = i + 5)
                    {


                        perilPremium[policyExposurePerils[i]["perilCharacteristicsLocator"].ToString()] = new { yearlyPremium = worksheet.Cells[rowCount, 43]?.Value.ToString() };
                        perilPremium[policyExposurePerils[i + 1]["perilCharacteristicsLocator"].ToString()] = new { yearlyPremium = worksheet.Cells[rowCount, 40]?.Value.ToString() };
                        perilPremium[policyExposurePerils[i + 2]["perilCharacteristicsLocator"].ToString()] = new { yearlyPremium = worksheet.Cells[rowCount, 46]?.Value.ToString() };
                        perilPremium[policyExposurePerils[i + 3]["perilCharacteristicsLocator"].ToString()] = new { yearlyPremium = worksheet.Cells[rowCount, 49]?.Value.ToString() };

                        int? value1 = Convert.ToInt32(worksheet.Cells[rowCount, 51]?.Value);
                        int? value2 = Convert.ToInt32(worksheet.Cells[rowCount, 53]?.Value);

                        perilPremium[policyExposurePerils[i + 4]["perilCharacteristicsLocator"].ToString()] = new
                        {
                            yearlyPremium = ((value1 ?? 0) + (value2 ?? 0)).ToString()
                        };


                        rowCount++;
                    }
                    var response = new
                    {
                        pricedPerilCharacteristics = perilPremium,
                    };

                    return Ok(response);
                }
            }
            catch (Exception ex)
            {
                // 5. Handle exceptions
                return StatusCode(500, new { error = "An error occurred while processing the request" });
            }

        }

        [HttpPost]
        [Route("excelRater_new")]
        public async Task<IActionResult> excelRater_new(JsonObject requestBody)
        {
            try
            {
                var policyExposurePerils = requestBody["policyExposurePerils"];

                var perilPremium = new Dictionary<string, object>();

                var mediaLocator = requestBody["policy"]["characteristics"][0]["fieldValues"]["excel_sheet"][0].ToString();

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

                        int perilsCount = 0;

                        if (requestBody.ContainsKey("policyExposurePerils") && requestBody["policyExposurePerils"] is JsonArray arr)
                        {
                            perilsCount = arr.Count;
                            // Now, 'length' contains the number of elements in the array
                        }


                        int j = 0;

                        while (j < perilsCount)
                        {
                            for (int row = 5; row <= rowCount; row++)
                            {
                                if (worksheet.Cells[package.Workbook.Names[sheet1.Cells[16, 6]?.Value.ToString() + $"{row - 4}"].Address].Value != null)
                                {


                                    var BI_premiun = worksheet.Cells[row, 43]?.Value.ToString();
                                    var BBB_premiun = worksheet.Cells[row, 40]?.Value.ToString();
                                    var Earthquake_premiun = worksheet.Cells[row, 46]?.Value.ToString();
                                    var Flood_premiun = worksheet.Cells[row, 49]?.Value.ToString();

                                    int? value1 = Convert.ToInt32(worksheet.Cells[row, 51]?.Value);
                                    int? value2 = Convert.ToInt32(worksheet.Cells[row, 53]?.Value);

                                    var Additional_premiun = ((value1 ?? 0) + (value2 ?? 0)).ToString();

                                    List<string> coverages = new List<string>();

                                    if (BI_premiun != "0")
                                    {
                                        coverages.Add(BI_premiun);
                                    }
                                    if (BBB_premiun != "0")
                                    {
                                        coverages.Add(BBB_premiun);
                                    }
                                    if (Earthquake_premiun != "0")
                                    {
                                        coverages.Add(Earthquake_premiun);
                                    }
                                    if (Flood_premiun != "0")
                                    {
                                        coverages.Add(Flood_premiun);
                                    }
                                    if (BBB_premiun != "0")
                                    {
                                        coverages.Add(Additional_premiun);
                                    }

                                    int count = coverages.Count;

                                    for (int i = 0; i < count; i++)
                                    {
                                        perilPremium[policyExposurePerils[j]["perilCharacteristicsLocator"].ToString()] = new { yearlyPremium = coverages[i] };
                                        j++;
                                    }

                                }
                            }
                        }

                        var response = new
                        {
                            pricedPerilCharacteristics = perilPremium,
                        };

                        return Ok(response);
                    }
                }
                else
                {
                    return StatusCode(500, new { error = "Excel not found." });
                }
            }
            catch (Exception ex)
            {
                // 5. Handle exceptions
                return StatusCode(500, new { error = "An error occurred while processing the request. " + ex.Message });
            }

        }
    }
}
