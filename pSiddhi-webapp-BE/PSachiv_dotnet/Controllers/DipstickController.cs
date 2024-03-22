using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json.Linq;
using PSachiv_dotnet.Services;
using System.Net.Http.Headers;

namespace PSachiv_dotnet.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class DipstickController : Controller
    {
        private readonly AccessTokenService _accessTokenService;
        public DipstickController(AccessTokenService accessTokenService)
        {
            _accessTokenService = accessTokenService;
        }
        [HttpGet("GetAllEntriesByDipstickId")] // get all values based on dipstick id
        public async Task<IActionResult> GetAllEntriesByDipstickId(string dipstick_id)
        {
            try
            {
                var accessToken = await _accessTokenService.GetAccessTokenAsync();
                using var httpClient = new HttpClient();
                httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                var response = await httpClient.GetAsync("https://graph.microsoft.com/v1.0/sites/7bhrxr.sharepoint.com,28983962-2b27-4b16-976c-24ebb19788d6,9b8c48bb-f5a5-4e40-92d6-8978f10efaad/drives/b!YjmYKCcrFkuXbCTrsZeI1rtIjJul9UBOktaJePEO-q2B8AEr-kgpQJgsYIVKte_z/items/01MP4UW3W3WITBG7KQYBDIFLOJG5NZPTSJ/workbook/worksheets('Sheet1')/usedRange");
                if (response.IsSuccessStatusCode)
                {
                    var jsonResponse = await response.Content.ReadAsStringAsync();
                    var jsonObject = JObject.Parse(jsonResponse);

                    // Get the array of values from the response
                    var valuesArray = jsonObject["values"] as JArray;

                    // Use LINQ to find entries matching the provided requirement ID
                    var matchingEntries = valuesArray
                        .Where(entry => entry.Count() > 0 && entry[0].ToString() == dipstick_id)
                        .Select(entry => new
                        {
                            dipstick_id = entry[0].ToString(),
                            years_of_experience_needed = entry[1].ToString(),
                            criticality = entry[2].ToString(),
                            closure_range = entry[3].ToString(),
                            category = entry[4].ToString(),
                            direct_client_interaction = entry[5].ToString(),
                            prior_team_handling_experience = entry[6].ToString(),
                            Round = entry[7].ToString(),
                            Screening = entry[8].ToString(),
                            customer_facing_role = entry[9].ToString(),
                            communication_grade = entry[10].ToString(),
                            client_Interview_necessary = entry[11].ToString(),
                            internal_benchmark = entry[12].ToString(),
                            individual_or_teamplayer = entry[13].ToString(),
                            Preferred_method_of_Sourcing = entry[14].ToString(),
                            Target_Organizations = entry[15].ToString(),
                            Summary = entry[16].ToString(),
                            created_by = entry[17].ToString(),
                            created_on = DateTime.FromOADate(double.Parse(entry[18].ToString())).ToString("M/d/yyyy"), // Convert serial number to date
                            modified_by = entry[19].ToString(),
                            modified_on = DateTime.FromOADate(double.Parse(entry[20].ToString())).ToString("M/d/yyyy") // Convert serial number to date
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
