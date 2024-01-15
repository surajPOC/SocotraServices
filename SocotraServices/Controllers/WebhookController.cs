using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json.Linq;
using System.Net.Http.Headers;
using System.Net.Mail;
using System.Text.Json.Nodes;

namespace SocotraServices.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class WebhookController : ControllerBase
    {
        private readonly IHttpClientFactory _httpClientFactory;

        private readonly IConfiguration _configuration;

        public WebhookController(IHttpClientFactory httpClientFactory, IConfiguration configuration)
        {
            _httpClientFactory = httpClientFactory;
            _configuration = configuration;
        }


        [HttpPost]
        [Route("quotationMail")]
        public async Task<IActionResult> SendQuotationSchedule([FromBody] JsonObject requestBody)
        {
            try
            {
                var policyLocator = requestBody["data"]["policyLocator"].ToString();

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
                dynamic policy_data = JObject.Parse(policyData);

                var email = policy_data.characteristics?[0]?.fieldValues?.email[0].ToString();
                var docUrl = policy_data.documents[0]?.url.ToString();
                var documentType = policy_data.documents[0]?.displayName.ToString();


                var documentBytes = await httpClient.GetByteArrayAsync(docUrl);

                // Save the document as a temporary file (you may want to change this)
                var tempFilePath = Path.Combine("E:\\SocotraServices\\Docs\\", $"{policyLocator}_{documentType}.pdf");
                await System.IO.File.WriteAllBytesAsync(tempFilePath, documentBytes);

                Console.WriteLine(tempFilePath);

                SendEmailWithAttachment(tempFilePath, email, documentType + "-" + policyLocator);

                return Ok("Email sent");

               
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("Error: " + ex.Message);
                return StatusCode(500, "An error occurred while processing the request.  "+ ex.Message);
            }
        }

        // Function for sending email
        private void SendEmailWithAttachment(string attachmentFilePath, string email, string subject)
        {
            try
            {
                using (SmtpClient smtpClient = new SmtpClient("smtp.gmail.com"))
                {
                    smtpClient.Port = 587;
                    smtpClient.Credentials = new System.Net.NetworkCredential(_configuration["senderMail"], _configuration["senderPassword"]);
                    smtpClient.EnableSsl = true;
                    

                    using (MailMessage mail = new MailMessage())
                    {
                        mail.From = new MailAddress(_configuration["senderMail"]);
                        mail.To.Add(email);
                        mail.Subject = subject;
                        mail.Body = "Please find the attached Quotation Schedule.";

                        // Attach the PDF to the email
                        mail.Attachments.Add(new Attachment(attachmentFilePath));

                        smtpClient.Send(mail);

                        Console.WriteLine("Email sent Successfully");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error sending email: {ex.Message}");
            }
        }


    }
}
