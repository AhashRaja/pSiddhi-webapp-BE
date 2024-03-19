using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Routing.Constraints;
using Microsoft.Graph.Models.Security;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using PSachiv_dotnet.Services;
using System.Net.Http.Headers;
using System.Text;

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

        [HttpGet("GetAllEntriesByCandidateId")] // get all values based on candidate id
        public async Task<IActionResult> GetAllEntriesByCandidateId(string candidateId, string requirementid)
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
                        .Select(entry => new
                        {
                            candidate_id = entry[0].ToString(),
                            candidate_Name = entry[1].ToString(),
                            requirement_ID = entry[2].ToString(),
                            recruiter = entry[3].ToString(),
                            tech_Skill = entry[4].ToString(),
                            education = entry[5].ToString(),
                            educational_Details_in_Percentage = entry[6].ToString(),
                            Type_of_source = entry[7].ToString(),
                            Company_Experience = entry[8].ToString(),
                            Years_of_experience = entry[9].ToString(),
                            Communication = entry[10].ToString(),
                            CTC_Details = entry[11].ToString(),
                            CTC_Breakup = entry[12].ToString(),
                            Reason_for_Change = entry[13].ToString(),
                            Last_working_date = DateTime.FromOADate(double.Parse(entry[14].ToString())).ToString("M/d/yyyy"),
                            Any_Other_offer = entry[15].ToString(),
                            Contact_Number = entry[16].ToString(),
                            Offers_at_hand_Value = entry[17].ToString(),
                            Offer_status = entry[18].ToString(),
                            Candidate_reply_offer = entry[19].ToString(),
                            Onboarding_status = entry[20].ToString(),
                            Overall_status = entry[21].ToString(),
                            Shortlisted_date = entry[22].ToString(),
                            Created_by = DateTime.FromOADate(double.Parse(entry[23].ToString())).ToString("M/d/yyyy"),
                            Created_on = entry[24].ToString(),
                            Modified_by = DateTime.FromOADate(double.Parse(entry[25].ToString())).ToString("M/d/yyyy"),
                            Modified_on = entry[26].ToString(),
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
        [HttpGet("GetSpecificEntriesByCandidateId")] // get specific values based on candidate id
        public async Task<IActionResult> GetSpecificEntriesByCandidateId(string candidateId, string requirementid)
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
                        .Select(entry => new
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
        [HttpPost("AddEntry")]
        public async Task<IActionResult> AddEntry([FromBody] Candidate can)
        {
            try
            {
                var accessToken = await _accessTokenService.GetAccessTokenAsync();
                using var httpClient = new HttpClient();
                httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                // Fetch the latest candidate ID and increment it
                var latestCandidateId = await GetLatestCandidateId();
                var newCandidateId = latestCandidateId;


                var requestBody = new EntryRequestModel
                {
                    values = new List<List<object>>
                    {
                        new List<object>
                        {
                            newCandidateId.ToString(),
                            can.candidate_Name,
                            can.requirement_ID,
                            can.recruiter,
                            can.tech_Skill,
                            can.education,
                            can.educational_Details_in_Percentage,
                            can.Type_of_source,
                            can.Company_Experience,
                            can.Years_of_experience,
                            can.Communication,
                            can.CTC_Details,
                            can.CTC_Breakup,
                            can.Reason_for_Change,
                            can.Last_working_date,
                            can.Any_Other_offer,
                            can.Contact_Number,
                            can.Candidate_Email_Id,
                            can.Offers_at_hand_Value,
                            can.Offer_status,
                            can.Candidate_reply_offer,
                            can.Onboarding_status,
                            can.Overall_status,
                            can.Shortlisted_date,
                            can.Created_by,
                            can.Created_on,
                            can.Modified_by,
                            can.Modified_on

                        }
                     }
                };
                var jsonBody = JsonConvert.SerializeObject(requestBody);


                var tableId = "{62549E24-B12A-41B6-A044-8E3993A7F7B0}";
                // Construct the URL for adding rows to the table
                var url = $"https://graph.microsoft.com/v1.0/sites/7bhrxr.sharepoint.com,28983962-2b27-4b16-976c-24ebb19788d6,9b8c48bb-f5a5-4e40-92d6-8978f10efaad/drive/items/01MP4UW3TCJE43KAMG75CLESVWC7CDFZNJ/workbook/tables/" + tableId + "/rows";

                // Send the POST request to add the entry
                var response = await httpClient.PostAsync(url, new StringContent(jsonBody, Encoding.UTF8, "application/json"));

                if (response.IsSuccessStatusCode)
                {
                    return Ok("Entry added successfully");
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

        [HttpPost("AddSpecificEntry")]
        public async Task<IActionResult> AddSpecificEntry([FromBody] Candidate can)
        {
            try
            {
                var accessToken = await _accessTokenService.GetAccessTokenAsync();
                using var httpClient = new HttpClient();
                httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                // Fetch the latest candidate ID and increment it
                var latestCandidateId = await GetLatestCandidateId();
                var newCandidateId = latestCandidateId;


                var requestBody = new EntryRequestModel
                {
                    values = new List<List<object>>
                    {
                        new List<object>
                        {
                            newCandidateId.ToString(),
                            can.candidate_Name,
                            "",
                            can.recruiter,
                            can.tech_Skill,
                            can.education,
                            can.educational_Details_in_Percentage,
                            can.Type_of_source,
                            can.Company_Experience,
                            can.Years_of_experience,
                            can.Communication,
                            can.CTC_Details,
                            can.CTC_Breakup,
                            can.Reason_for_Change,
                            can.Last_working_date,
                            can.Any_Other_offer,
                            can.Contact_Number,
                            "",
                            "",
                            "",
                            "",
                            "",
                            "",
                            "",
                            "",
                            "",
                            "",
                            ""
                        }
                     }
                };
                var jsonBody = JsonConvert.SerializeObject(requestBody);

                var tableId = "{62549E24-B12A-41B6-A044-8E3993A7F7B0}";
                // Construct the URL for adding rows to the table
                var url = $"https://graph.microsoft.com/v1.0/sites/7bhrxr.sharepoint.com,28983962-2b27-4b16-976c-24ebb19788d6,9b8c48bb-f5a5-4e40-92d6-8978f10efaad/drive/items/01MP4UW3TCJE43KAMG75CLESVWC7CDFZNJ/workbook/tables/" + tableId + "/rows";

                // Send the POST request to add the entry
                var response = await httpClient.PostAsync(url, new StringContent(jsonBody, Encoding.UTF8, "application/json"));

                if (response.IsSuccessStatusCode)
                {
                    return Ok("Entry added successfully");
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
        private async Task<string> GetLatestCandidateId()
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

                    // Find the index of the last row without null entries
                    var lastRowIndex = valuesArray
                        .Cast<JArray>()
                        .Select((row, index) => new { Row = row, Index = index })
                        .LastOrDefault(item => item.Row != null && item.Row.HasValues && item.Row.Any(v => !string.IsNullOrEmpty(v.ToString())))
                        ?.Index ?? -1;

                    if (lastRowIndex != -1)
                    {
                        // Retrieve the latest candidate ID from the first column of the last row
                        var latestCandidateId = valuesArray[lastRowIndex][0].ToString();
                        // Extract numeric portion from the candidate ID
                        var numericPortion = latestCandidateId.Substring(3);
                        // Parse the numeric portion to integer
                        if (int.TryParse(numericPortion, out int latestNumericId))
                        {
                            // Increment the numeric ID by 1
                            latestNumericId++;
                            // Format the numeric ID back to "CANXXX" format
                            var formattedLatestId = $"CAN{latestNumericId:D3}";
                            return formattedLatestId;
                        }
                        else
                        {
                            // If parsing fails, return an error
                            return "Error: Unable to parse requirement ID";
                        }
                    }
                    else
                    {
                        // Handle case where no data is available
                        return "CAN001"; // Assuming default starting ID is CAN001
                    }
                }
                else
                {
                    var errorResponse = await response.Content.ReadAsStringAsync();
                    // Handle error response
                    return "Error: " + errorResponse; // Or return a default value, or throw an exception
                }
            }
            catch (Exception ex)
            {
                // Handle exceptions
                return "Exception: " + ex.Message; // Or return a default value, or rethrow the exception
            }
        }
    }
    public class Candidate
    {
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
        public string Candidate_Email_Id { get; set; }
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
