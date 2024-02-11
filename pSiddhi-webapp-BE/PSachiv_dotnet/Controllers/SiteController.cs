using Microsoft.AspNetCore.Mvc;
using PSachiv_dotnet.Models;
using PSachiv_dotnet.Services;
using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using Newtonsoft.Json;
using System.Text.Json;
using System.Threading.Tasks;

namespace PSachiv_dotnet.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class SiteController : ControllerBase
    {
        private readonly AccessTokenService _accessTokenService;

        public SiteController(AccessTokenService accessTokenService)
        {
            _accessTokenService = accessTokenService;
        }

        [HttpGet("Sites")]
        public async Task<IActionResult> GetSites()
        {
            try
            {
                // Get access token using the AccessTokenService
                var accessToken = await _accessTokenService.GetAccessTokenAsync();

                // Create HttpClient instance
                using var httpClient = new HttpClient();

                // Set authorization header
                httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                // Make request to the Microsoft Graph API
                var response = await httpClient.GetAsync("https://graph.microsoft.com/v1.0/sites?search=NewSite");

                // Check if request was successful
                if (response.IsSuccessStatusCode)
                {
                    // Read and return the response content
                    var jsonResponse = await response.Content.ReadAsStringAsync();
                    return Ok(jsonResponse);
                }
                else
                {
                    // Return error message if request failed
                    var errorResponse = await response.Content.ReadAsStringAsync();
                    return StatusCode((int)response.StatusCode, errorResponse);
                }
            }
            catch (Exception ex)
            {
                // Return 500 Internal Server Error if an exception occurs
                return StatusCode(500, ex.Message);
            }
        }

        [HttpGet("Sites/{siteId}")]
        public async Task<IActionResult> GetSite(string siteId)
        {
            try
            {
                // Get access token using the AccessTokenService
                var accessToken = await _accessTokenService.GetAccessTokenAsync();

                // Create HttpClient instance
                using var httpClient = new HttpClient();

                // Set authorization header
                httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                // Construct the request URL with the siteId
                var requestUrl = $"https://graph.microsoft.com/v1.0/sites/{siteId}/drive/root:/Psiddhi_HR_Recruitment:/children";

                // Make request to the Microsoft Graph API
                var response = await httpClient.GetAsync(requestUrl);

                // Check if request was successful
                if (response.IsSuccessStatusCode)
                {
                    // Read and return the response content
                    var jsonResponse = await response.Content.ReadAsStringAsync();
                    return Ok(jsonResponse);
                }
                else
                {
                    // Return error message if request failed
                    var errorResponse = await response.Content.ReadAsStringAsync();
                    return StatusCode((int)response.StatusCode, errorResponse);
                }
            }
            catch (Exception ex)
            {
                // Return 500 Internal Server Error if an exception occurs
                return StatusCode(500, ex.Message);
            }
        }

        [HttpPost("AddTable")]
        public async Task<IActionResult> AddTableToWorkbook()
        {
            try
            {
                // Get access token using the AccessTokenService
                var accessToken = await _accessTokenService.GetAccessTokenAsync();

                // Create HttpClient instance
                using var httpClient = new HttpClient();

                // Set authorization header
                httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                // Construct the request URL
                var requestUrl = "https://graph.microsoft.com/v1.0/sites/7bhrxr.sharepoint.com,28983962-2b27-4b16-976c-24ebb19788d6,9b8c48bb-f5a5-4e40-92d6-8978f10efaad/drive/items/01MP4UW3QBXSO774XI6RD2WPIM54NWRAPN/workbook/tables/add";

                // Create request content
                var requestData = new
                {
                    address = "Sheet1!A1:A10",
                    hasHeaders = true
                };
                var jsonContent = JsonConvert.SerializeObject(requestData);
                var content = new StringContent(jsonContent, Encoding.UTF8, "application/json");

                // Make POST request to the Microsoft Graph API
                var response = await httpClient.PostAsync(requestUrl, content);

                // Check if request was successful
                if (response.IsSuccessStatusCode)
                {
                    // Read and return the response content
                    var jsonResponse = await response.Content.ReadAsStringAsync();
                    return Ok(jsonResponse);
                }
                else
                {
                    // Return error message if request failed
                    var errorResponse = await response.Content.ReadAsStringAsync();
                    return StatusCode((int)response.StatusCode, errorResponse);
                }
            }
            catch (Exception ex)
            {
                // Return 500 Internal Server Error if an exception occurs
                return StatusCode(500, ex.Message);
            }
        }
    }
}
