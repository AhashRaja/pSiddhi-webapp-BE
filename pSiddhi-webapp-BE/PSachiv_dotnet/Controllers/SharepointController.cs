using Microsoft.AspNetCore.Mvc;
using System.Text;
using Microsoft.Extensions.Logging;
using System;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Authorization;

namespace PSachiv_dotnet.Controllers
{
    [Authorize]
    [ApiController]
    [Route("api/sharepoint")]
    public class SharepointController : Controller
    {
        private readonly ILogger<SharepointController> _logger;
        private readonly HttpClient _httpClient;

        public SharepointController(ILogger<SharepointController> logger, IHttpClientFactory httpClientFactory)
        {
            _logger = logger;
            _httpClient = httpClientFactory.CreateClient("SharePointClient");
        }

        [HttpGet("get-items")]
        public async Task<IActionResult> GetItems()
        {
            try
            {
                HttpResponseMessage response = await _httpClient.GetAsync("/_api/web/lists/getbytitle('SampleList')/items");

                if (response.IsSuccessStatusCode)
                {
                    string content = await response.Content.ReadAsStringAsync();
                    // Parse and process the response content as needed
                    return Ok(content);
                }
                else
                {
                    return StatusCode((int)response.StatusCode, response.ReasonPhrase);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error retrieving items from SharePoint");
                return StatusCode(500, "Internal Server Error");
            }
        }

        /* [HttpPost("add-item")]
         public async Task<IActionResult> AddItem([FromBody] YourModel model)
         {
             try
             {
                 string json = Newtonsoft.Json.JsonConvert.SerializeObject(model);
                 HttpContent content = new StringContent(json, Encoding.UTF8, "application/json");
                 HttpResponseMessage response = await _httpClient.PostAsync("/_api/web/lists/getbytitle('SampleList')/items", content);

                 if (response.IsSuccessStatusCode)
                 {
                     string createdItemId = response.Headers.GetValues("Location").FirstOrDefault()?.Split('/').Last();
                     // Return the ID of the created item
                     return Ok(createdItemId);
                 }
                 else
                 {
                     return StatusCode((int)response.StatusCode, response.ReasonPhrase);
                 }
             }
             catch (Exception ex)
             {
                 _logger.LogError(ex, "Error adding item to SharePoint");
                 return StatusCode(500, "Internal Server Error");
             }
         }

         [HttpPut("update-item/{itemId}")]
         public async Task<IActionResult> UpdateItem(int itemId, [FromBody] YourModel model)
         {
             try
             {
                 string json = Newtonsoft.Json.JsonConvert.SerializeObject(model);
                 HttpContent content = new StringContent(json, Encoding.UTF8, "application/json");
                 HttpResponseMessage response = await _httpClient.PutAsync($"/_api/web/lists/getbytitle('SampleList')/items({itemId})", content);

                 if (response.IsSuccessStatusCode)
                 {
                     return Ok();
                 }
                 else
                 {
                     return StatusCode((int)response.StatusCode, response.ReasonPhrase);
                 }
             }
             catch (Exception ex)
             {
                 _logger.LogError(ex, $"Error updating item with ID {itemId} in SharePoint");
                 return StatusCode(500, "Internal Server Error");
             }
         }

         [HttpDelete("delete-item/{itemId}")]
         public async Task<IActionResult> DeleteItem(int itemId)
         {
             try
             {
                 HttpResponseMessage response = await _httpClient.DeleteAsync($"/_api/web/lists/getbytitle('SampleList')/items({itemId})");

                 if (response.IsSuccessStatusCode)
                 {
                     return Ok();
                 }
                 else
                 {
                     return StatusCode((int)response.StatusCode, response.ReasonPhrase);
                 }
             }
             catch (Exception ex)
             {
                 _logger.LogError(ex, $"Error deleting item with ID {itemId} from SharePoint");
                 return StatusCode(500, "Internal Server Error");
             }
         }*/
    }
}
