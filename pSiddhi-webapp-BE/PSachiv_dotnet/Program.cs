using Microsoft.Identity.Web;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using PSachiv_dotnet.Services;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddControllers();
// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddHttpClient(); // Add HttpClient service
builder.Services.AddSwaggerGen();
builder.Services.AddMicrosoftIdentityWebApiAuthentication(builder.Configuration); // Add Microsoft Graph API authentication
builder.Services.AddScoped<AccessTokenService>();

// Define the CORS policy name
var MyAllowSpecificOrigins = "_myAllowSpecificOrigins";

// Configure CORS policy
builder.Services.AddCors(options =>
{
    options.AddPolicy(name: MyAllowSpecificOrigins,
                      policy =>
                      {
                          // Get the UI domain from configuration
                          var uiDomain = builder.Configuration.GetSection("UI_Domain").Value;

                          if (uiDomain != null)
                          {
                              // Allow requests from the specified UI domain
                              policy.WithOrigins(uiDomain)
                                    .AllowAnyHeader()
                                    .AllowAnyMethod()
                                    .AllowCredentials();
                          }
                          else
                          {
                              throw new InvalidOperationException("UI_Domain is not configured correctly.");
                          }
                      });
});


var app = builder.Build();

// Use CORS with the specified policy
app.UseCors(MyAllowSpecificOrigins);

/*// Configure CORS
app.UseCors(options => 
        { 
            options.AllowAnyOrigin()
            .AllowAnyMethod()
            .AllowAnyHeader(); 
        });*/

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseHttpsRedirection();
app.UseRouting(); // Add UseRouting middleware
app.UseAuthorization();

app.MapControllers();

app.Run();
