using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json.Linq;
using PSachiv_dotnet.Services;
using System.Net.Http.Headers;

namespace PSachiv_dotnet.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class CandidateApprovalController : Controller
    {
        private readonly AccessTokenService _accessTokenService;
        public CandidateApprovalController(AccessTokenService accessTokenService)
        {
            _accessTokenService = accessTokenService;
        }
        [HttpGet("GetAllEntriesByCandidateApprovalId")] // get all values based on candidate approval id
        public async Task<IActionResult> GetAllEntriesByCandidateApprovalId(string candidate_approval_id)
        {
            try
            {
                var accessToken = await _accessTokenService.GetAccessTokenAsync();
                using var httpClient = new HttpClient();
                httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                var response = await httpClient.GetAsync("https://graph.microsoft.com/v1.0/sites/7bhrxr.sharepoint.com,28983962-2b27-4b16-976c-24ebb19788d6,9b8c48bb-f5a5-4e40-92d6-8978f10efaad/drives/b!YjmYKCcrFkuXbCTrsZeI1rtIjJul9UBOktaJePEO-q2B8AEr-kgpQJgsYIVKte_z/items/01MP4UW3T2XN6J3LDBEZAYOGKWT2UDTM2Q/workbook/worksheets('Sheet1')/usedRange");
                if (response.IsSuccessStatusCode)
                {
                    var jsonResponse = await response.Content.ReadAsStringAsync();
                    var jsonObject = JObject.Parse(jsonResponse);

                    // Get the array of values from the response
                    var valuesArray = jsonObject["values"] as JArray;

                    // Use LINQ to find entries matching the provided requirement ID
                    var matchingEntries = valuesArray
                        .Where(entry => entry.Count() > 0 && entry[0].ToString() == candidate_approval_id)
                        .Select(entry => new
                        {
                            candidate_approval_id = entry[0].ToString(),
                            candidate_Id = entry[1].ToString(),
                            approval_by = entry[2].ToString(),
                            approval_status = entry[3].ToString(),
                            approval_feedback = entry[4].ToString(),
                            deliberation_status = entry[5].ToString(),
                            deliberation_feedback = entry[6].ToString(),
                            created_by = entry[7].ToString(),
                            created_on = DateTime.FromOADate(double.Parse(entry[8].ToString())).ToString("M/d/yyyy"), // Convert serial number to date
                            modified_by = entry[9].ToString(),
                            modified_on = DateTime.FromOADate(double.Parse(entry[10].ToString())).ToString("M/d/yyyy") // Convert serial number to date

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
}
