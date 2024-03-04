using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Routing.Constraints;
using Newtonsoft.Json.Linq;
using PSachiv_dotnet.Services;
using System.Net.Http.Headers;

namespace PSachiv_dotnet.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class CandidateMasterController : Controller
    {
        private readonly AccessTokenService _accessTokenService;

        public CandidateMasterController(AccessTokenService accessTokenService)
        {
            _accessTokenService = accessTokenService;
        }
        public class EntryRequestModel
        {
            public List<List<object>> values { get; set; }
        }

        [HttpGet("UsedRange")] // get all the response values
        public async Task<IActionResult> GetUsedRange()
        {
            try
            {
                var accessToken = await _accessTokenService.GetAccessTokenAsync();
                using var httpClient = new HttpClient();
                httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                var response = await httpClient.GetAsync("https://graph.microsoft.com/v1.0/sites/7bhrxr.sharepoint.com,28983962-2b27-4b16-976c-24ebb19788d6,9b8c48bb-f5a5-4e40-92d6-8978f10efaad/drives/b!YjmYKCcrFkuXbCTrsZeI1rtIjJul9UBOktaJePEO-q2B8AEr-kgpQJgsYIVKte_z/items/01MP4UW3TCJE43KAMG75CLESVWC7CDFZNJ/workbook/worksheets('Sheet1')/usedRange");

                if (response.IsSuccessStatusCode)
                {
                    var jsonResponse = await response.Content.ReadAsStringAsync();
                    return Ok(jsonResponse);
                }
                else
                {
                    var errorResponse = await response.Content.ReadAsStringAsync();
                    return StatusCode((int)response.StatusCode, errorResponse);
                }
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message);
            }
        }

        [HttpGet("GetAllCandidateEntries")] // Get all candidate entries 
        public async Task<IActionResult> GetEntries()
        {
            try
            {
                var accessToken = await _accessTokenService.GetAccessTokenAsync();
                using var httpClient = new HttpClient();
                httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                var response = await httpClient.GetAsync("https://graph.microsoft.com/v1.0/sites/7bhrxr.sharepoint.com,28983962-2b27-4b16-976c-24ebb19788d6,9b8c48bb-f5a5-4e40-92d6-8978f10efaad/drives/b!YjmYKCcrFkuXbCTrsZeI1rtIjJul9UBOktaJePEO-q2B8AEr-kgpQJgsYIVKte_z/items/01MP4UW3TCJE43KAMG75CLESVWC7CDFZNJ/workbook/worksheets('Sheet1')/usedRange");
                if (response.IsSuccessStatusCode)
                {
                    var jsonResponse = await response.Content.ReadAsStringAsync();
                    var jsonObject = JObject.Parse(jsonResponse);

                    // Get the array of values from the response
                    var valuesArray = jsonObject["values"] as JArray;

                    // Use LINQ to select all entries
                    var allEntries = valuesArray
                        .Where((entry, index) => index > 0 && entry.Count() > 0 && !string.IsNullOrEmpty(entry[0]?.ToString())) // Skip the first entry, exclude empty entries, and filter out entries with empty reqId
                        .Select(entry => new Candidate
                        {
                            candidate_id = entry[0].ToString(),
                            candidate_Name = entry[1].ToString(),
                            requirement_ID = entry[2].ToString(),
                            recruiter = entry[3].ToString()
                        })
                        .ToList();

                    return Ok(allEntries);
                }
                else
                {
                    var errorResponse = await response.Content.ReadAsStringAsync();
                    return StatusCode((int)response.StatusCode, errorResponse);
                }
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message);
            }
        }
        [HttpGet("GetEntriesByCandidateId")] // get values based on candidate id
        public async Task<IActionResult> GetEntriesByRequirementId(string candidateId, string requirementid)
        {
            try
            {
                var accessToken = await _accessTokenService.GetAccessTokenAsync();
                using var httpClient = new HttpClient();
                httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                var response = await httpClient.GetAsync("https://graph.microsoft.com/v1.0/sites/7bhrxr.sharepoint.com,28983962-2b27-4b16-976c-24ebb19788d6,9b8c48bb-f5a5-4e40-92d6-8978f10efaad/drives/b!YjmYKCcrFkuXbCTrsZeI1rtIjJul9UBOktaJePEO-q2B8AEr-kgpQJgsYIVKte_z/items/01MP4UW3TCJE43KAMG75CLESVWC7CDFZNJ/workbook/worksheets('Sheet1')/usedRange");
                if (response.IsSuccessStatusCode)
                {
                    var jsonResponse = await response.Content.ReadAsStringAsync();
                    var jsonObject = JObject.Parse(jsonResponse);

                    // Get the array of values from the response
                    var valuesArray = jsonObject["values"] as JArray;

                    // Use LINQ to find entries matching the provided candidate ID
                    var matchingEntries = valuesArray
                        .Where(entry => entry.Count() > 0 && (entry[0].ToString() == candidateId && entry[2].ToString() == requirementid))
                        .Select(entry => new Candidate
                        {
                            candidate_id = entry[0].ToString(),
                            candidate_Name = entry[1].ToString(),
                            requirement_ID = entry[2].ToString(),
                            recruiter = entry[3].ToString()
                        })
                        .FirstOrDefault(); // Assuming you only expect one matching entry

                    if (matchingEntries != null)
                    {
                        return Ok(matchingEntries);
                    }
                    else
                    {
                        return NotFound(); // Return 404 if no matching entry is found
                    }
                }
                else
                {
                    var errorResponse = await response.Content.ReadAsStringAsync();
                    return StatusCode((int)response.StatusCode, errorResponse);
                }
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message);
            }
        }
    }
    public class Candidate
    {
        [System.Text.Json.Serialization.JsonIgnore]
        public string candidate_id { get; set; }
        public string candidate_Name { get; set; }
        public string requirement_ID { get; set; }
        public string recruiter { get; set; }
        public string tech_Skill { get; set; }
        public string education { get; set; }
        public string educational_Details_in_Percentage { get; set; }
        public string Type_of_source { get; set; }
        public string Company_Experience { get; set; }
        public string Years_of_experience { get; set; }
        public string Communication { get; set; }
        public string CTC_Details { get; set; }
        public string CTC_Breakup { get; set; }
        public string Reason_for_Change { get; set; }
        public string Last_working_date { get; set; }
        public string Any_Other_offer { get; set; }
        public string Contact_Number { get; set; }
        public string Candiate_Email_Id { get; set; }
        public string Offers_at_hand_Value { get; set; }
        public string Offer_status { get; set; }
        public string Candidate_reply_offer { get; set; }
        public string Onboarding_status { get; set; }
        public string Overall_status { get; set; }
        public string Shortlisted_date { get; set; }
        public string Created_by { get; set; }
        public string Created_on { get; set; }
        public string Modified_by { get; set; }
        public string Modified_on { get; set; }
    }
}
