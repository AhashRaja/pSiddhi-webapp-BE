using PSachiv_dotnet.Models;
using System;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;

namespace PSachiv_dotnet.Services
{
    public class AccessTokenService
    {
        private readonly IHttpClientFactory _httpClientFactory;

        public AccessTokenService(IHttpClientFactory httpClientFactory)
        {
            _httpClientFactory = httpClientFactory;
        }

        public async Task<string> GetAccessTokenAsync()
        {
            var httpClient = _httpClientFactory.CreateClient();
            var tokenRequest = new Dictionary<string, string>
            {
                { "client_id", "d322e6be-133a-4c9b-9bbb-df2566b83b14" },
                { "client_secret", "sZ98Q~~ErbJjZDOsj2FpK5ru3RN~5oohglejqbyl" },
                { "grant_type", "client_credentials" },
                { "scope", "https://graph.microsoft.com/.default" }
            };

            try
            {
                var requestContent = new FormUrlEncodedContent(tokenRequest);

                var response = await httpClient.PostAsync("https://login.microsoftonline.com/8399c1c2-9c1b-4d0d-97fb-e0cfed231878/oauth2/v2.0/token", requestContent);

                if (response.IsSuccessStatusCode)
                {
                    var jsonResponse = await response.Content.ReadAsStringAsync();
                    var tokenResponse = JsonSerializer.Deserialize<TokenResponse>(jsonResponse);
                    return tokenResponse.access_token;
                }
                else
                {
                    var errorResponse = await response.Content.ReadAsStringAsync();
                    throw new Exception($"Error: {response.StatusCode}, Reason: {errorResponse}");
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to retrieve access token: {ex.Message}");
            }
        }
    }
}
