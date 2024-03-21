using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using PSachiv_dotnet.Services;
using System.Text.Json.Serialization;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;

namespace PSachiv_dotnet.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class RequirementController : Controller
    {
        private readonly AccessTokenService _accessTokenService;

        public RequirementController(AccessTokenService accessTokenService)
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

                var response = await httpClient.GetAsync("https://graph.microsoft.com/v1.0/sites/7bhrxr.sharepoint.com,28983962-2b27-4b16-976c-24ebb19788d6,9b8c48bb-f5a5-4e40-92d6-8978f10efaad/drives/b!YjmYKCcrFkuXbCTrsZeI1rtIjJul9UBOktaJePEO-q2B8AEr-kgpQJgsYIVKte_z/items/01MP4UW3UXIP62MK4R6NB2V2WP6DNQTEYI/workbook/worksheets('Sheet1')/usedRange");

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

        [HttpGet("GetAllRequirementEntries")] // Get all requirement entries 
        public async Task<IActionResult> GetEntries()
        {
            try
            {
                var accessToken = await _accessTokenService.GetAccessTokenAsync();
                using var httpClient = new HttpClient();
                httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                var response = await httpClient.GetAsync("https://graph.microsoft.com/v1.0/sites/7bhrxr.sharepoint.com,28983962-2b27-4b16-976c-24ebb19788d6,9b8c48bb-f5a5-4e40-92d6-8978f10efaad/drives/b!YjmYKCcrFkuXbCTrsZeI1rtIjJul9UBOktaJePEO-q2B8AEr-kgpQJgsYIVKte_z/items/01MP4UW3UXIP62MK4R6NB2V2WP6DNQTEYI/workbook/worksheets('Sheet1')/usedRange");
                if (response.IsSuccessStatusCode)
                {
                    var jsonResponse = await response.Content.ReadAsStringAsync();
                    var jsonObject = JObject.Parse(jsonResponse);

                    // Get the array of values from the response
                    var valuesArray = jsonObject["values"] as JArray;

                    // Use LINQ to select all entries
                    var allEntries = valuesArray
                        .Where((entry, index) => index > 0 && entry.Count() > 0 && !string.IsNullOrEmpty(entry[0]?.ToString())) // Skip the first entry, exclude empty entries, and filter out entries with empty reqId
                        .Select(entry => new Req
                        {
                            reqId = entry[0].ToString(),
                            reqName = entry[1].ToString(),
                            position_type = entry[4].ToString(),
                            status = entry[17].ToString()
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

        [HttpGet("GetAllEntriesByRequirementId")] // get all values based on requirement id
        public async Task<IActionResult> GetAllEntriesByRequirementId(string requirementId)
        {
            try
            {
                var accessToken = await _accessTokenService.GetAccessTokenAsync();
                using var httpClient = new HttpClient();
                httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                var response = await httpClient.GetAsync("https://graph.microsoft.com/v1.0/sites/7bhrxr.sharepoint.com,28983962-2b27-4b16-976c-24ebb19788d6,9b8c48bb-f5a5-4e40-92d6-8978f10efaad/drives/b!YjmYKCcrFkuXbCTrsZeI1rtIjJul9UBOktaJePEO-q2B8AEr-kgpQJgsYIVKte_z/items/01MP4UW3UXIP62MK4R6NB2V2WP6DNQTEYI/workbook/worksheets('Sheet1')/usedRange");
                if (response.IsSuccessStatusCode)
                {
                    var jsonResponse = await response.Content.ReadAsStringAsync();
                    var jsonObject = JObject.Parse(jsonResponse);

                    // Get the array of values from the response
                    var valuesArray = jsonObject["values"] as JArray;

                    // Use LINQ to find entries matching the provided requirement ID
                    var matchingEntries = valuesArray
                        .Where(entry => entry.Count() > 0 && entry[0].ToString() == requirementId)
                        .Select(entry => new Req
                        {
                            reqId = entry[0].ToString(),
                            dipId = entry[1].ToString(),
                            reqName = entry[2].ToString(),
                            band = entry[3].ToString(),
                            level = entry[4].ToString(),
                            position_type = entry[5].ToString(),
                            number_of_openings = entry[6].ToString(),
                            account = entry[7].ToString(),
                            coe = entry[8].ToString(),
                            coe_manager = entry[9].ToString(),
                            criticality = entry[10].ToString(),
                            years_of_experience_needed = entry[11].ToString(),
                            expected_date_of_closure = DateTime.FromOADate(double.Parse(entry[12].ToString())).ToString("M/d/yyyy"),
                            requirement_Type = entry[13].ToString(),
                            oc_stage1_approval_status = entry[14].ToString(),
                            oc_stage2_approval_status = entry[15].ToString(),
                            created_by = entry[16].ToString(),
                            created_on = DateTime.FromOADate(double.Parse(entry[17].ToString())).ToString("M/d/yyyy"),
                            modified_by = entry[18].ToString(),
                            modified_on = DateTime.FromOADate(double.Parse(entry[19].ToString())).ToString("M/d/yyyy"),
                            status = entry[20].ToString(),

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

        [HttpGet("GetSpecificEntriesByRequirementId")] // get values based on requirement id
        public async Task<IActionResult> GetSpecificEntriesByRequirementId(string requirementId)
        {
            try
            {
                var accessToken = await _accessTokenService.GetAccessTokenAsync();
                using var httpClient = new HttpClient();
                httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                var response = await httpClient.GetAsync("https://graph.microsoft.com/v1.0/sites/7bhrxr.sharepoint.com,28983962-2b27-4b16-976c-24ebb19788d6,9b8c48bb-f5a5-4e40-92d6-8978f10efaad/drives/b!YjmYKCcrFkuXbCTrsZeI1rtIjJul9UBOktaJePEO-q2B8AEr-kgpQJgsYIVKte_z/items/01MP4UW3UXIP62MK4R6NB2V2WP6DNQTEYI/workbook/worksheets('Sheet1')/usedRange");
                if (response.IsSuccessStatusCode)
                {
                    var jsonResponse = await response.Content.ReadAsStringAsync();
                    var jsonObject = JObject.Parse(jsonResponse);

                    // Get the array of values from the response
                    var valuesArray = jsonObject["values"] as JArray;

                    // Use LINQ to find entries matching the provided requirement ID
                    var matchingEntries = valuesArray
                        .Where(entry => entry.Count() > 0 && entry[0].ToString() == requirementId)
                        .Select(entry => new
                        {
                            reqId = entry[0].ToString(),
                            reqName = entry[2].ToString(),
                            band = entry[3].ToString(),
                            level = entry[4].ToString(),
                            position_type = entry[5].ToString(),
                            number_of_openings = entry[6].ToString(),
                            account = entry[7].ToString(),
                            coe = entry[8].ToString(),
                            coe_manager = entry[9].ToString(),
                            criticality = entry[10].ToString(),
                            years_of_experience_needed = entry[11].ToString(),
                            status = entry[17].ToString()
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
        public async Task<IActionResult> AddEntry([FromBody] Req req)
        {
            try
            {
                var accessToken = await _accessTokenService.GetAccessTokenAsync();
                using var httpClient = new HttpClient();
                httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                // Fetch the latest requirement ID and increment it
                var latestRequirementId = await GetLatestRequirementId();
                var newRequirementId = latestRequirementId;


                var requestBody = new EntryRequestModel
                {
                    values = new List<List<object>>
                    {
                        new List<object>
                        {
                            newRequirementId.ToString(),
                            req.dipId,
                            req.reqName,
                            req.band,
                            req.level,
                            req.position_type,
                            req.number_of_openings,
                            req.account,
                            req.coe,
                            req.coe_manager,
                            req.criticality,
                            req.years_of_experience_needed,
                            req.expected_date_of_closure,
                            req.requirement_Type,
                            req.oc_stage1_approval_status,
                            req.oc_stage2_approval_status,
                            req.created_by,
                            req.created_on,
                            req.modified_by,
                            req.modified_on,
                            req.status

                        }
                     }
                };
                var jsonBody = JsonConvert.SerializeObject(requestBody);


                /*EntryRequestModel entryRequestModel = new EntryRequestModel();
                entryRequestModel.values[0][0] = newRequirementId.ToString();
                entryRequestModel.values[0][1] = req.reqName.ToString();

                // Add the new requirement ID to the entry values
                entry.values.Insert(0, new List<string> { newRequirementId });

                // Construct the request body from the modified entry model
                var requestBody = new { Values = entry.values };

                // Remove the empty array at the beginning of the values array
                requestBody.values.RemoveAll(x => x.Count == 0);

                // Insert the newRequirementId at the beginning of the inner array
                //requestBody.values[1].Insert(0, newRequirementId);

                var jsonBody = JsonConvert.SerializeObject(requestBody);

                // Fetch the table ID dynamically based on the requirement ID
                //var tableId = await GetTableIdByRequirementId(newRequirementId, httpClient);*/


                var tableId = "{3F5DC472-92FD-4BB2-A036-DB8AE3FD89C2}";
                // Construct the URL for adding rows to the table
                var url = $"https://graph.microsoft.com/v1.0/sites/7bhrxr.sharepoint.com,28983962-2b27-4b16-976c-24ebb19788d6,9b8c48bb-f5a5-4e40-92d6-8978f10efaad/drive/items/01MP4UW3UXIP62MK4R6NB2V2WP6DNQTEYI/workbook/tables/" + tableId + "/rows";

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
        public async Task<IActionResult> AddSpecificEntry([FromBody] Req req)
        {
            try
            {
                var accessToken = await _accessTokenService.GetAccessTokenAsync();
                using var httpClient = new HttpClient();
                httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                // Fetch the latest requirement ID and increment it
                var latestRequirementId = await GetLatestRequirementId();
                var newRequirementId = latestRequirementId;


                var requestBody = new EntryRequestModel
                {
                    values = new List<List<object>>
                    {
                        new List<object>
                        {
                            newRequirementId.ToString(),
                            "",
                            req.reqName,
                            req.band,
                            req.level,
                            req.position_type,
                            req.number_of_openings,
                            req.account,
                            req.coe,
                            req.coe_manager,
                            req.criticality,
                            req.years_of_experience_needed,
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

                var tableId = "{3F5DC472-92FD-4BB2-A036-DB8AE3FD89C2}";
                // Construct the URL for adding rows to the table
                var url = $"https://graph.microsoft.com/v1.0/sites/7bhrxr.sharepoint.com,28983962-2b27-4b16-976c-24ebb19788d6,9b8c48bb-f5a5-4e40-92d6-8978f10efaad/drive/items/01MP4UW3UXIP62MK4R6NB2V2WP6DNQTEYI/workbook/tables/" + tableId + "/rows";

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
        private async Task<string> GetLatestRequirementId()
        {
            try
            {
                var accessToken = await _accessTokenService.GetAccessTokenAsync();
                using var httpClient = new HttpClient();
                httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                var response = await httpClient.GetAsync("https://graph.microsoft.com/v1.0/sites/7bhrxr.sharepoint.com,28983962-2b27-4b16-976c-24ebb19788d6,9b8c48bb-f5a5-4e40-92d6-8978f10efaad/drives/b!YjmYKCcrFkuXbCTrsZeI1rtIjJul9UBOktaJePEO-q2B8AEr-kgpQJgsYIVKte_z/items/01MP4UW3UXIP62MK4R6NB2V2WP6DNQTEYI/workbook/worksheets('Sheet1')/usedRange");
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
                        // Retrieve the latest requirement ID from the first column of the last row
                        var latestRequirementId = valuesArray[lastRowIndex][0].ToString();
                        // Extract numeric portion from the requirement ID
                        var numericPortion = latestRequirementId.Substring(3);
                        // Parse the numeric portion to integer
                        if (int.TryParse(numericPortion, out int latestNumericId))
                        {
                            // Increment the numeric ID by 1
                            latestNumericId++;
                            // Format the numeric ID back to "REQXXX" format
                            var formattedLatestId = $"REQ{latestNumericId:D3}";
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
                        return "REQ001"; // Assuming default starting ID is REQ001
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

        [HttpPut("UpdateSpecificEntry/{rowId}")] // Define the route to include the row ID
        public async Task<IActionResult> UpdateSpecificEntry(string rowId, [FromBody] Req updatedReq)
        {
            try
            {
                var accessToken = await _accessTokenService.GetAccessTokenAsync();
                using var httpClient = new HttpClient();
                httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                var tableId = "{3F5DC472-92FD-4BB2-A036-DB8AE3FD89C2}";
                // Construct the URL for adding rows to the table
                var url = $"https://graph.microsoft.com/v1.0/sites/7bhrxr.sharepoint.com,28983962-2b27-4b16-976c-24ebb19788d6,9b8c48bb-f5a5-4e40-92d6-8978f10efaad/drive/items/01MP4UW3UXIP62MK4R6NB2V2WP6DNQTEYI/workbook/tables/" + tableId + "/rows/"+ rowId;

                // Fetch the existing values
                var response = await httpClient.GetAsync(url);
                if (!response.IsSuccessStatusCode)
                {
                    var errorResponse = await response.Content.ReadAsStringAsync();
                    return StatusCode((int)response.StatusCode, errorResponse);
                }

                var jsonResponse = await response.Content.ReadAsStringAsync();
                var jsonObject = JObject.Parse(jsonResponse);
                var valuesArray = jsonObject["values"] as JArray;

                // Use LINQ to find the entry matching the provided row ID
                var matchingEntry = valuesArray
                    .Where(entry => entry.Count() > 0 && entry[0].ToString() == rowId) // Assuming the row ID is in the first column
                    .FirstOrDefault(); // Assuming you only expect one matching entry

                if (matchingEntry != null)
                {
                    // Update the reqName field of the matching entry
                    matchingEntry[2] = updatedReq.reqName; // Assuming reqName is in the third column (index 2)

                    // Construct the updated response with only reqId and updated reqName
                    var updatedResponse = new
                    {
                        reqId = matchingEntry[0].ToString(),
                        reqName = matchingEntry[2].ToString() // Updated reqName
                    };

                    return Ok(updatedResponse);
                }
                else
                {
                    return NotFound(); // Return 404 if no matching entry is found
                }

            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message);
            }
        }



    }

    public class Req
    {
        public string reqId { get; set; }
        public string dipId { get; set; }
        public string reqName { get; set; }//
        public string band { get; set; }//
        public string level { get; set; }//
        public string position_type { get; set; }//
        public string number_of_openings { get; set; }//
        public string account { get; set; }//
        public string coe { get; set; }//
        public string coe_manager { get; set; }//
        public string criticality { get; set; }//
        public string years_of_experience_needed { get; set; }//
        public string? expected_date_of_closure { get; set; }
        public string? requirement_Type { get; set; }
        public string? oc_stage1_approval_status { get; set; }
        public string? oc_stage2_approval_status { get; set; }
        public string? created_by { get; set; }
        public string? created_on { get; set; }
        public string? modified_by { get; set; }
        public string? modified_on { get; set; }
        public string? status { get; set; }
    }
}


/* var matchingEntries = new List<JToken>();
 Req req = new Req();
 // Find entries matching the provided requirement ID
 foreach (var entry in valuesArray)
 {
     if (entry.Count() > 0 && entry[0].ToString() == requirementId.ToString())
     {
         //matchingEntries.Add(entry);
         req.reqId = entry[0].ToString();
         req.reqName = entry[1].ToString();
         req.Position_type = entry[4].ToString();
         req.status = entry[17].ToString();
     }
 }                
 return Ok(req);*/




/* // Construct the request body from the entry model
 var requestBody = new { values = entry.Values };

 // Serialize the request body to JSON
 var jsonBody = JsonConvert.SerializeObject(requestBody);

 // Replace {tableId} with the actual ID of the table where you want to add rows
 var tableId = "{59B3CEF6-5295-44F5-B4EF-39E6A09E7F83}";

 // Construct the URL for adding rows to the table
 var url = $"https://graph.microsoft.com/v1.0/sites/7bhrxr.sharepoint.com,28983962-2b27-4b16-976c-24ebb19788d6,9b8c48bb-f5a5-4e40-92d6-8978f10efaad/drive/items/01MP4UW3QBXSO774XI6RD2WPIM54NWRAPN/workbook/tables/{tableId}/rows";
*/


/*private async Task<string> GetTableIdByRequirementId(double requirementId, HttpClient httpClient)
        {
            //var response = await httpClient.GetAsync("https://graph.microsoft.com/v1.0/tables");

            var tableId = "{59B3CEF6-5295-44F5-B4EF-39E6A09E7F83}";
            var url = "https://graph.microsoft.com/v1.0/sites/7bhrxr.sharepoint.com,28983962-2b27-4b16-976c-24ebb19788d6,9b8c48bb-f5a5-4e40-92d6-8978f10efaad/drive/items/01MP4UW3UXIP62MK4R6NB2V2WP6DNQTEYI/workbook/tables/" + tableId + "/rows";

            var response = await httpClient.GetAsync(url);
            if (response.IsSuccessStatusCode)
            {
                var jsonResponse = await response.Content.ReadAsStringAsync();
                var jsonObject = JObject.Parse(jsonResponse);

                // Assuming the response contains an array of tables
                var tablesArray = jsonObject["values"] as JArray;

                // Find the table ID based on the requirement ID
                foreach (var table in tablesArray)
                {
                    var tableId = table["@odata.id"].ToString();
                    var valuesArray = table["values"] as JArray;
                    var requirementIds = valuesArray.Select(v => v[0].ToObject<int>());
                    if (requirementIds.Contains(requirementId))
                    {
                        // Return the table ID if the requirement ID is found
                        return tableId;
                    }
                }
            }

            // Return null if the table ID is not found
            return null;
        }*/