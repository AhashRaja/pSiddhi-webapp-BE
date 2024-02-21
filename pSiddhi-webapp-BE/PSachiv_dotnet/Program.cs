using Microsoft.Identity.Web;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using PSachiv_dotnet.Services;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.Hosting;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddControllers();
// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddHttpClient(); // Add HttpClient service
builder.Services.AddSwaggerGen();
builder.Services.AddMicrosoftIdentityWebApiAuthentication(builder.Configuration); // Add Microsoft Graph API authentication
builder.Services.AddScoped<AccessTokenService>();

var app = builder.Build();

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
