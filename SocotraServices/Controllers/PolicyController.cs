using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Text.Json.Nodes;

namespace SocotraServices.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class PolicyController : ControllerBase
    {
        private readonly IHttpClientFactory _httpClientFactory;

        private readonly IConfiguration _configuration;

        public PolicyController(IHttpClientFactory httpClientFactory, IConfiguration configuration)
        {
            _httpClientFactory = httpClientFactory;
            _configuration = configuration;
        }

        [HttpPost]
        [Route("createPolicyholder")]
        public async Task<IActionResult> createPolicyholder([FromBody] JsonObject requestBody)
        {
            try
            {

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

                // Create PolicyHolder
                var createPolicyholderResponse = await httpClient.PostAsJsonAsync(_configuration["socotraBaseUrl"] + "/policyholder/create", requestBody);

                if (!createPolicyholderResponse.IsSuccessStatusCode)
                {
                    string response = await createPolicyholderResponse.Content.ReadAsStringAsync();
                    JToken parsedJson = JToken.Parse(response);
                    return StatusCode(500, "Failed to create policyholder from Socotra API \n  API Response: " + parsedJson);
                }


                // Fetching neccessary data
                var policyholderData = await createPolicyholderResponse.Content.ReadAsStringAsync();
                var json = JsonConvert.DeserializeObject(policyholderData);
                var beautifiedPolicyholderResponseJson = JsonConvert.SerializeObject(json, Formatting.Indented);



                return Ok(beautifiedPolicyholderResponseJson);


            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("Error: " + ex.Message);
                return StatusCode(500, "An error occurred while processing the request.  " + ex.Message);
            }
        }


        [HttpPost]
        [Route("createPolicy")]
        public async Task<IActionResult> createPolicy([FromBody] JsonObject requestBody)
        {
            try
            {

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

                // Create Policy
                var createPolicyResponse = await httpClient.PostAsJsonAsync(_configuration["socotraBaseUrl"] + "/policy?summarizeQuotes=true", requestBody);

                if (!createPolicyResponse.IsSuccessStatusCode)
                {
                    string response = await createPolicyResponse.Content.ReadAsStringAsync();
                    JToken parsedJson = JToken.Parse(response);
                    return StatusCode(500, "Failed to create policy from Socotra API \n  Message: " + parsedJson);
                }


                // Fetching neccessary data
                var policyData = await createPolicyResponse.Content.ReadAsStringAsync();
                //dynamic policy_data = JsonObject.Parse(policyData);
                var json = JsonConvert.DeserializeObject(policyData);
                var beautifiedPolicyJson = JsonConvert.SerializeObject(json, Formatting.Indented);



                return Ok(beautifiedPolicyJson);


            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("Error: " + ex.Message);
                return StatusCode(500, "An error occurred while processing the request.  " + ex.Message);
            }
        }

        [HttpPost]
        [Route("PricePolicy")]
        public async Task<IActionResult> PricePolicy(string policyLocator)
        {
            try
            {

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

                // Price Policy
                var pricePolicyResponse = await httpClient.PostAsJsonAsync(_configuration["socotraBaseUrl"] + "/policy/" + policyLocator + "/price", "");

                if (!pricePolicyResponse.IsSuccessStatusCode)
                {
                    string response = await pricePolicyResponse.Content.ReadAsStringAsync();
                    JToken parsedJson = JToken.Parse(response);
                    return StatusCode(500, "Failed to price policy from Socotra API.   \n Message :  " + parsedJson);
                }


                // Fetching neccessary data
                var priceData = await pricePolicyResponse.Content.ReadAsStringAsync();
                //dynamic policy_data = JsonObject.Parse(policyData);
                var json = JsonConvert.DeserializeObject(priceData);
                var beautifiedPriceDataJson = JsonConvert.SerializeObject(json, Formatting.Indented);



                return Ok(beautifiedPriceDataJson);


            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("Error: " + ex.Message);
                return StatusCode(500, "An error occurred while processing the request.  " + ex.Message);
            }
        }

        [HttpPost]
        [Route("LockAndPriceQuote")]
        public async Task<IActionResult> LockAndPriceQuote(string quoteLocator)
        {
            try
            {

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

                // Lock and Price a quote
                var lockAndPriceQuoteResponse = await httpClient.PatchAsync(_configuration["socotraBaseUrl"] + "/quotes/" + quoteLocator + "/quote", null);

                if (!lockAndPriceQuoteResponse.IsSuccessStatusCode)
                {
                    string response = await lockAndPriceQuoteResponse.Content.ReadAsStringAsync();
                    JToken parsedJson = JToken.Parse(response);
                    return StatusCode(500, "Failed to Lock And price quote from Socotra API.   \n Message :  " + parsedJson);
                }


                // Fetching neccessary data
                var LockAndPriceData = await lockAndPriceQuoteResponse.Content.ReadAsStringAsync();
                //dynamic policy_data = JsonObject.Parse(policyData);
                var json = JsonConvert.DeserializeObject(LockAndPriceData);
                var beautifiedLockAndPriceDataJson = JsonConvert.SerializeObject(json, Formatting.Indented);


                return Ok(beautifiedLockAndPriceDataJson);


            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("Error: " + ex.Message);
                return StatusCode(500, "An error occurred while processing the request.  " + ex.Message);
            }
        }

        [HttpPost]
        [Route("AcceptQuote")]
        public async Task<IActionResult> AcceptQuote(string quoteLocator)
        {
            try
            {

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

                // Acccept a quote
                var acceptQuoteResponse = await httpClient.PatchAsync(_configuration["socotraBaseUrl"] + "/quotes/" + quoteLocator + "/accept", null);

                if (!acceptQuoteResponse.IsSuccessStatusCode)
                {
                    string response = await acceptQuoteResponse.Content.ReadAsStringAsync();
                    JToken parsedJson = JToken.Parse(response);
                    return StatusCode(500, "Failed to accept quote from Socotra API.   \n Message :  " + parsedJson);
                }


                // Fetching neccessary data
                var AcceptQuoteData = await acceptQuoteResponse.Content.ReadAsStringAsync();
                //dynamic policy_data = JsonObject.Parse(policyData);
                var json = JsonConvert.DeserializeObject(AcceptQuoteData);
                var beautifiedAcceptQuoteDataJson = JsonConvert.SerializeObject(json, Formatting.Indented);



                return Ok(beautifiedAcceptQuoteDataJson);


            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("Error: " + ex.Message);
                return StatusCode(500, "An error occurred while processing the request.  " + ex.Message);
            }
        }

        [HttpPost]
        [Route("IssuePolicy")]
        public async Task<IActionResult> IssuePolicy(string policyLocator)
        {
            try
            {

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

                // Issue Policy
                var issuePolicyResponse = await httpClient.PostAsJsonAsync(_configuration["socotraBaseUrl"] + "/policy/" + policyLocator + "/issue", "");

                if (!issuePolicyResponse.IsSuccessStatusCode)
                {
                    string response = await issuePolicyResponse.Content.ReadAsStringAsync();
                    JToken parsedJson = JToken.Parse(response);
                    return StatusCode(500, "Failed to issue policy from Socotra API.   \n Message :  " + parsedJson);
                }


                // Fetching neccessary data
                var issuePolicyData = await issuePolicyResponse.Content.ReadAsStringAsync();
                //dynamic policy_data = JsonObject.Parse(policyData);
                var json = JsonConvert.DeserializeObject(issuePolicyData);
                var beautifiedIssuePolicyDataJson = JsonConvert.SerializeObject(json, Formatting.Indented);


                return Ok(beautifiedIssuePolicyDataJson);


            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("Error: " + ex.Message);
                return StatusCode(500, "An error occurred while processing the request.  " + ex.Message);
            }
        }

        [HttpGet]
        [Route("FetchQuote")]
        public async Task<IActionResult> FetchQuote(string quoteLocator)
        {
            try
            {

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

                // Fetch Quote
                var fetchQuoteResponse = await httpClient.GetAsync(_configuration["socotraBaseUrl"] + "/quotes/" + quoteLocator);

                if (!fetchQuoteResponse.IsSuccessStatusCode)
                {
                    string response = await fetchQuoteResponse.Content.ReadAsStringAsync();
                    JToken parsedJson = JToken.Parse(response);
                    return StatusCode(500, "Failed to fetch quote from Socotra API.   \n Message :  " + parsedJson);
                }


                // Fetching neccessary data
                var fetchQuoteData = await fetchQuoteResponse.Content.ReadAsStringAsync();
                //dynamic policy_data = JsonObject.Parse(policyData);
                var json = JsonConvert.DeserializeObject(fetchQuoteData);
                var beautifiedFetchQuoteDataJson = JsonConvert.SerializeObject(json, Formatting.Indented);


                return Ok(beautifiedFetchQuoteDataJson);


            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("Error: " + ex.Message);
                return StatusCode(500, "An error occurred while processing the request.  " + ex.Message);
            }
        }



        [HttpGet]
        [Route("downloadPdf")]
        public void downloadPdf(string formUrl, string filePath)
        {
            using (var client = new HttpClient())
            {
                using (var s = client.GetStreamAsync(formUrl))
                {
                    //var uniqueFileName = $"{Guid.NewGuid()}.pdf";
                    using (var fs = new FileStream(filePath, System.IO.FileMode.OpenOrCreate))
                    {
                        s.Result.CopyTo(fs);
                    }
                }
            }

        }
    }
}
