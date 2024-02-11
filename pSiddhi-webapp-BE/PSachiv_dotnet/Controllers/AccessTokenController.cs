using Microsoft.AspNetCore.Mvc;
using PSachiv_dotnet.Services;
using PSachiv_dotnet.Models;
using System;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;

namespace PSachiv_dotnet.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class AccessTokenController : ControllerBase
    {
        private readonly AccessTokenService _accessTokenService;

        public AccessTokenController(AccessTokenService accessTokenService)
        {
            _accessTokenService = accessTokenService;
        }

        [HttpGet]
        public async Task<IActionResult> GetAccessToken()
        {
            try
            {
                var accessToken = await _accessTokenService.GetAccessTokenAsync();
                return Ok(accessToken);
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message);
            }
        }
    }
      
}


/*[HttpGet]
     public async Task<IActionResult> GetAccessToken()
     {
         try

             // Create HTTP client
             var httpClient = _httpClientFactory.CreateClient();

             // Define token request parameters
             var tokenRequest = new
             {
                 client_id = "d322e6be-133a-4c9b-9bbb-df2566b83b14",
                 client_secret = "sZ98Q~~ErbJjZDOsj2FpK5ru3RN~5oohglejqbyl",
                 grant_type = "client_credentials",
                 scope = "https://graph.microsoft.com/.default" // Scope for Microsoft Graph API
             };


             // Convert token request parameters to JSON
             var jsonRequest = JsonSerializer.Serialize(tokenRequest);

             // Create HTTP request message
             var request = new HttpRequestMessage(HttpMethod.Post, "https://login.microsoftonline.com/8399c1c2-9c1b-4d0d-97fb-e0cfed231878/oauth2/v2.0/token")
             {
                 Content = new StringContent(jsonRequest, System.Text.Encoding.UTF8, "application/json")
             };

             // Send request to token endpoint
             var response = await httpClient.SendAsync(request);

             // Check if request was successful
             if (response.IsSuccessStatusCode)
             {
                 // Parse token response
                 var jsonResponse = await response.Content.ReadAsStringAsync();
                 var tokenResponse = JsonSerializer.Deserialize<TokenResponse>(jsonResponse);

                 // Access token
                 var accessToken = tokenResponse.access_token;

                 // Use the access token (e.g., call Microsoft Graph API)

                 return Ok(accessToken);
             }
             else
             {
                 return StatusCode((int)response.StatusCode, response.ReasonPhrase);
             }
         }
         catch (Exception ex)
         {
             return StatusCode(500, ex.Message);
         }
     }
 }*/



//[HttpGet("Sites")]
/*public async Task<IActionResult> GetSites()
{
    try
    {
        // Get access token
        var accessToken = await _accessTokenService.GetAccessTokenAsync();

        // Create HTTP client
        var httpClient = _httpClientFactory.CreateClient();

        // Create request message with authorization header
        var request = new HttpRequestMessage(HttpMethod.Get, "https://graph.microsoft.com/v1.0/sites?search=NewSite");
        request.Headers.Add("Authorization", $"Bearer {accessToken}");

        // Send request
        var response = await httpClient.SendAsync(request);

        // Check if request was successful
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
}*/

