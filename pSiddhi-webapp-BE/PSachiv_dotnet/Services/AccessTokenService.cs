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
                { "client_id", "48851c48-610d-4a79-bb2a-2f67757f9261" },
                { "client_secret", "n9.8Q~cNDgXHP9KtNz9Xy0YR6.Cq9vii1Y~3tadD" },
                { "grant_type", "client_credentials" },
                { "scope", "https://graph.microsoft.com/.default" }
            };

            try
            {
                var requestContent = new FormUrlEncodedContent(tokenRequest);

                var response = await httpClient.PostAsync("https://login.microsoftonline.com/041abf65-a74f-472f-ac6a-c0ab0f1f8679/oauth2/v2.0/token", requestContent);

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
