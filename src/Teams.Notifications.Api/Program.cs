global using System;
global using System.Collections;
global using System.Collections.Concurrent;
global using System.Collections.Generic;
global using System.Data;
global using System.Diagnostics;
global using System.Globalization;
global using System.IdentityModel.Tokens.Jwt;
global using System.IO;
global using System.Linq;
global using System.Net;
global using System.Net.Http;
global using System.Net.Http.Json;
global using System.Reflection;
global using System.Text.Json;
global using System.Text.Json.Nodes;
global using System.Text.Json.Serialization;
global using System.Text.RegularExpressions;
global using System.Threading;
global using System.Threading.Tasks;
global using AdaptiveCards;
global using Azure.Core;
global using Azure.Identity;
global using Microsoft.Agents.Authentication;
global using Microsoft.Agents.Builder;
global using Microsoft.Agents.Builder.App;
global using Microsoft.Agents.Builder.App.AdaptiveCards;
global using Microsoft.Agents.Builder.State;
global using Microsoft.Agents.Core.Models;
global using Microsoft.Agents.Core.Serialization;
global using Microsoft.Agents.Extensions.Teams.Connector;
global using Microsoft.Agents.Extensions.Teams.Models;
global using Microsoft.Agents.Hosting.AspNetCore;
global using Microsoft.Agents.Storage;
global using Microsoft.AspNetCore.Authentication;
global using Microsoft.AspNetCore.Authorization;
global using Microsoft.AspNetCore.Builder;
global using Microsoft.AspNetCore.Http;
global using Microsoft.AspNetCore.Mvc;
global using Microsoft.AspNetCore.Mvc.ApplicationModels;
global using Microsoft.AspNetCore.OpenApi;
global using Microsoft.Azure.Cosmos;
global using Microsoft.Extensions.Configuration;
global using Microsoft.Extensions.DependencyInjection;
global using Microsoft.Extensions.Logging;
global using Microsoft.Extensions.Options;
global using Microsoft.Graph.Beta;
global using Microsoft.Graph.Beta.Models;
global using Microsoft.Identity.Client;
global using Microsoft.Identity.Web;
global using Microsoft.IdentityModel.Protocols;
global using Microsoft.IdentityModel.Protocols.OpenIdConnect;
global using Microsoft.IdentityModel.Tokens;
global using Microsoft.IdentityModel.Validators;
global using Microsoft.Kiota.Abstractions;
global using Microsoft.OpenApi;
global using Polly;
global using Polly.Retry;
global using Teams.Notifications.Api;
global using Teams.Notifications.Api.Action.Models;
global using Teams.Notifications.Api.Agents;
global using Teams.Notifications.Api.Agents.CardHandler;
global using Teams.Notifications.Api.DelegatingHandlers;
global using Teams.Notifications.Api.Extensions;
global using Teams.Notifications.Api.Filters;
global using Teams.Notifications.Api.Middlewares;
global using Teams.Notifications.Api.Models;
global using Teams.Notifications.Api.OpenApiTransformer;
global using Teams.Notifications.Api.Services;
global using Teams.Notifications.Api.Services.Interfaces;
global using Uniphar.Platform.Telemetry;
global using Activity = Microsoft.Agents.Core.Models.Activity;
global using Attachment = Microsoft.Agents.Core.Models.Attachment;
global using IMiddleware = Microsoft.Agents.Builder.IMiddleware;
global using RouteAttribute = Microsoft.AspNetCore.Mvc.RouteAttribute;
global using WebApplication = Microsoft.AspNetCore.Builder.WebApplication;


const string appPathPrefix = "platform-teams-notification-api";

var builder = WebApplication.CreateBuilder(args);

var environment = builder.Environment.EnvironmentName ?? throw new NoNullAllowedException("ASPNETCORE_ENVIRONMENT environment variable has to be set.");

TokenCredential credentials = new DefaultAzureCredential();

// this is what the bot is communicating on
builder.Services.AddHttpClient(typeof(RestChannelServiceClientFactory).FullName!).AddHttpMessageHandler<RequestAndResponseLoggerHandler>();


// Values from app registration
var clientId = builder.Configuration["AZURE_CLIENT_ID"] ?? throw new NoNullAllowedException("ClientId is required");
var tenantId = builder.Configuration["AZURE_TENANT_ID"] ?? throw new NoNullAllowedException("TenantId is required");

var clientSecret = builder.Configuration["ClientSecret"];

var environmentSuffix = environment == "prod" ? string.Empty : $".{environment}";

if (environment == "local") environmentSuffix = ".dev";

var apiUrl = new Uri($"https://api{environmentSuffix}.uniphar.ie/");


var jitterRandomizer = new Random();
builder.Services.AddHttpClient();
builder
    .Services
    .AddHttpClient("frontgate-api", client => client.BaseAddress = apiUrl)
    .AddStandardResilienceHandler();
// will use workload if available
if (!string.IsNullOrWhiteSpace(clientSecret))
{
    credentials = new ClientSecretCredential(tenantId, clientId, clientSecret);
    builder.Configuration.AddInMemoryCollection(new Dictionary<string, string?>
    {
        { "Connections:ServiceConnection:Settings:AuthType", "ClientSecret" },
        { "Connections:ServiceConnection:Settings:ClientId", clientId },
        { "Connections:ServiceConnection:Settings:TenantId", tenantId },
        { "Connections:ServiceConnection:Settings:ClientSecret", clientSecret },
        { "Connections:ServiceConnection:Settings:Scopes:0", "https://api.botframework.com/.default" }
    });
}
else
{
    var federatedTokenFile = builder.Configuration["AZURE_FEDERATED_TOKEN_FILE"] ?? throw new NoNullAllowedException("Token file is required");
    builder.Configuration.AddInMemoryCollection(new Dictionary<string, string?>
    {
        { "Connections:ServiceConnection:Settings:AuthType", "WorkloadIdentity" },
        { "Connections:ServiceConnection:Settings:ClientId", clientId },
        { "Connections:ServiceConnection:Settings:TenantId", tenantId },
        { "Connections:ServiceConnection:Settings:FederatedTokenFile", federatedTokenFile },
        { "Connections:ServiceConnection:Settings:Scopes:0", "https://api.botframework.com/.default" }
    });
}

builder.Services.AddSingleton(credentials);
builder.Services.AddSingleton(new GraphServiceClient(credentials));
builder.Services.AddTransient<RequestAndResponseLoggerHandler>();
builder.Services.AddTransient<ICardManagerService, CardManagerService>();
builder.Services.AddTransient<ITeamsManagerService, TeamsManagerService>();
builder.Services.AddTransient<IFrontgateApiService, FrontgateApiService>();
builder
    .Services
    .AddHttpClient("service-now-api", client => client.BaseAddress = apiUrl)
    .AddStandardResilienceHandler();
builder.Services.AddTransient<IServiceNowApiService, ServiceNowApiService>();

builder.Services.Configure<CosmosOptions>(builder.Configuration.GetSection(CosmosOptions.SectionName));

builder.Services.AddSingleton(sp =>
{
    var connectionString = sp.GetRequiredService<IOptions<CosmosOptions>>().Value.ConnectionString;
    return new CosmosClient(connectionString,
        new()
        {
            HttpClientFactory = () => new(new SocketsHttpHandler { PooledConnectionLifetime = TimeSpan.FromMinutes(5) }, false),
            SerializerOptions = new()
            {
                Indented = true,
                IgnoreNullValues = true,
                PropertyNamingPolicy = CosmosPropertyNamingPolicy.CamelCase
            }
        });
});
builder.Services.AddSingleton<ICosmosMessageStore, CosmosMessageStore>();
builder.Services.AddHealthChecks();
builder.Services.AddAgentAspNetAuthentication(builder.Configuration);
builder.Services.AddMemoryCache();
// Add ApplicationOptions
builder.AddAgentApplicationOptions();
builder.AddAgent<CardActionAgent>();
builder
    .Services
    .AddControllers(o =>
    {
        o.Conventions.Add(new HideChannelApi());
        o.Conventions.Add(new GlobalRouteConvention(appPathPrefix));
        o.Filters.Add<InvalidOperationExceptionFilter>();
    })
    .AddJsonOptions(options =>
    {
        options.JsonSerializerOptions.PropertyNamingPolicy = JsonNamingPolicy.CamelCase;
        options.JsonSerializerOptions.DictionaryKeyPolicy = JsonNamingPolicy.CamelCase;
        options.JsonSerializerOptions.DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull;
    });
if (environment != "local")
    // key vault is required for ApplicationInsights, since it needs the connection string, but locally we will remove it
    builder.Configuration.AddAzureKeyVault(new($"https://uni-devops-app-{environment}-kv.vault.azure.net/"), credentials);

builder.Services.AddSingleton<IMiddleware[]>(_ => [new CaptureMiddleware()]);
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddOpenApi(options =>
{
    options.AddDocumentTransformer((doc, _, _) =>
    {
        foreach (var server in doc.Servers ?? [])
            if (server.Url != null && server.Url.Contains("uniphar.ie"))
                server.Url = server.Url.Replace("http://", "https://");

        return Task.CompletedTask;
    });
    options.AddDocumentTransformer<AddAdaptiveCardDocsTransformer>();
    options.AddOperationTransformer<AddCorrectFileTransformer>();

    options.AddDocumentTransformer<BearerSecuritySchemeTransformer>();
});

// Register IStorage.  For development, MemoryStorage is suitable.
// For production Agents, persisted storage should be used so
// that state survives Agent restarts, and operate correctly
// in a cluster of Agent instances.
builder.Services.AddSingleton<IStorage, MemoryStorage>();

// Configure OpenTelemetry
builder.RegisterOpenTelemetry(appPathPrefix).Build();


var app = builder.Build();
await app.Services.GetRequiredService<ICosmosMessageStore>().EnsureContainerIsProvisioned();
app.MapHealthChecks("/health");
app.MapOpenApi(appPathPrefix + "/swagger/{documentName}/openapi.json");
app.UseSwaggerUI(c =>
{
    c.SwaggerEndpoint("v1/openapi.json", "V1");
    c.RoutePrefix = $"{appPathPrefix}/swagger";
});
app.MapControllers();
app.Run();