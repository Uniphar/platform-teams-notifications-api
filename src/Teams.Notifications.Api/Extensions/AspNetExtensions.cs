namespace Teams.Notifications.Api.Extensions;

internal static class AspNetExtensions
{
    private static readonly ConcurrentDictionary<string, ConfigurationManager<OpenIdConnectConfiguration>> _openIdMetadataCache = new();

    /// <summary>
    ///     Adds token validation typical for ABS and agent-to-agent.
    ///     default to Azure Public Cloud.
    /// </summary>
    /// <param name="services"></param>
    /// <param name="configuration"></param>
    public static void AddAgentAspNetAuthentication(this IServiceCollection services, IConfiguration configuration)
    {
        var tenantId = configuration["AZURE_TENANT_ID"] ?? throw new ArgumentNullException(nameof(configuration), "Missing AZURE_TENANT_ID");
        //ABS public cloud
        var validTokenIssuers = new List<string>
        {
            "https://api.botframework.com",
            "https://sts.windows.net/d6d49420-f39b-4df7-a1dc-d59a935871db/",
            "https://login.microsoftonline.com/d6d49420-f39b-4df7-a1dc-d59a935871db/v2.0",
            "https://sts.windows.net/f8cdef31-a31e-4b4a-93e4-5f571e91255a/",
            "https://login.microsoftonline.com/f8cdef31-a31e-4b4a-93e4-5f571e91255a/v2.0",
            "https://sts.windows.net/69e9b82d-4842-4902-8d1e-abc5b98a55e8/",
            "https://login.microsoftonline.com/69e9b82d-4842-4902-8d1e-abc5b98a55e8/v2.0",
            string.Format(CultureInfo.InvariantCulture, AuthenticationConstants.ValidTokenIssuerUrlTemplateV1, tenantId),
            string.Format(CultureInfo.InvariantCulture, AuthenticationConstants.ValidTokenIssuerUrlTemplateV2, tenantId)
        };

        var audiences = new List<string> { configuration["AZURE_CLIENT_ID"] ?? throw new NoNullAllowedException("ClientId is required") };

        const string azureBotServiceOpenIdMetadataUrl = AuthenticationConstants.PublicAzureBotServiceOpenIdMetadataUrl;

        // If the `OpenIdMetadataUrl` setting is not specified, use the default based on `IsGov`.  This is what is used to authenticate Entra ID tokens.
        const string openIdMetadataUrl = AuthenticationConstants.PublicOpenIdMetadataUrl;
        var openIdRefreshInterval = BaseConfigurationManager.DefaultAutomaticRefreshInterval;
        services
            .AddAuthentication("NotificationScheme")
            .AddMicrosoftIdentityWebApi(_ => { },
                microsoftIdentityOptions =>
                {
                    microsoftIdentityOptions.Instance = "https://login.microsoftonline.com/";
                    microsoftIdentityOptions.TenantId = configuration["AZURE_ENTRA_EXTERNAL_TENANT_ID"] ?? throw new NoNullAllowedException("EXTERNAL_TENANT_ID is required");
                    microsoftIdentityOptions.ClientId = configuration["platform-teams-notification-api-client-id"] ?? throw new NoNullAllowedException("ClientId is required");
                },
                "NotificationScheme");
        // authorization policies for the API
        services
            .AddAuthorizationBuilder()
            .AddPolicy("Teams.Notifications.Api.Writer",
                policy =>
                {
                    policy.RequireRole("teams.notifications.api.writer");
                    policy.AuthenticationSchemes.Add("NotificationScheme");
                });
        // authentication for the BOT
        services
            .AddAuthentication("AgentScheme")
            .AddJwtBearer("AgentScheme",
                options =>
                {
                    options.SaveToken = true;
                    options.TokenValidationParameters = new()
                    {
                        ValidateIssuer = true,
                        ValidateAudience = true,
                        ValidateLifetime = true,
                        ClockSkew = TimeSpan.FromMinutes(5),
                        ValidIssuers = validTokenIssuers,
                        ValidAudiences = audiences,
                        ValidateIssuerSigningKey = true,
                        RequireSignedTokens = true
                    };

                    // Using Microsoft.IdentityModel.Validators
                    options.TokenValidationParameters.EnableAadSigningKeyIssuerValidation();
                    EnableAadSigningKeyIssuerValidationForEntraTokens(options.TokenValidationParameters);

                    options.Events = new()
                    {
                        // Create a ConfigurationManager based on the requestor.  This is to handle ABS non-Entra tokens.
                        OnMessageReceived = async context =>
                        {
                            var authorizationHeader = context.Request.Headers.Authorization.ToString();

                            if (string.IsNullOrWhiteSpace(authorizationHeader))
                            {
                                // Default to AadTokenValidation handling
                                context.Options.TokenValidationParameters.ConfigurationManager ??= options.ConfigurationManager as BaseConfigurationManager;
                                await Task.CompletedTask.ConfigureAwait(false);
                                return;
                            }

                            var parts = authorizationHeader.Split(' ');
                            if (parts is not ["Bearer", _])
                            {
                                // Default to AadTokenValidation handling
                                context.Options.TokenValidationParameters.ConfigurationManager ??= options.ConfigurationManager as BaseConfigurationManager;
                                await Task.CompletedTask.ConfigureAwait(false);
                                return;
                            }

                            JwtSecurityToken token = new(parts[1]);
                            var issuer = token.Claims.FirstOrDefault(claim => claim.Type == AuthenticationConstants.IssuerClaim)?.Value;

                            if (AuthenticationConstants.BotFrameworkTokenIssuer.Equals(issuer))
                                // Use the Azure Bot authority for this configuration manager
                            {
                                context.Options.TokenValidationParameters.ConfigurationManager = _openIdMetadataCache.GetOrAdd(azureBotServiceOpenIdMetadataUrl,
                                    _ => new(azureBotServiceOpenIdMetadataUrl, new OpenIdConnectConfigurationRetriever(), new HttpClient())
                                    {
                                        AutomaticRefreshInterval = openIdRefreshInterval
                                    });
                            }
                            else
                            {
                                context.Options.TokenValidationParameters.ConfigurationManager = _openIdMetadataCache.GetOrAdd(openIdMetadataUrl,
                                    _ => new(openIdMetadataUrl, new OpenIdConnectConfigurationRetriever(), new HttpClient())
                                    {
                                        AutomaticRefreshInterval = openIdRefreshInterval
                                    });
                            }

                            await Task.CompletedTask.ConfigureAwait(false);
                        },

                        OnTokenValidated = _ => Task.CompletedTask,
                        OnForbidden = _ => Task.CompletedTask,
                        OnAuthenticationFailed = _ => Task.CompletedTask
                    };
                });
        services.AddAuthentication(options => { options.DefaultScheme = "NotificationScheme"; });
    }

    private static void EnableAadSigningKeyIssuerValidationForEntraTokens(TokenValidationParameters tokenValidationParameters)
    {
        var signingKeyValidator = tokenValidationParameters.IssuerSigningKeyValidator;
        if (signingKeyValidator is not null)
        {
            tokenValidationParameters.IssuerSigningKeyValidator = (securityKey, token, validationParameters) =>
                IsBotFrameworkToken(token) || signingKeyValidator(securityKey, token, validationParameters);
        }

        var signingKeyValidatorUsingConfiguration = tokenValidationParameters.IssuerSigningKeyValidatorUsingConfiguration;
        if (signingKeyValidatorUsingConfiguration is not null)
        {
            tokenValidationParameters.IssuerSigningKeyValidatorUsingConfiguration = (securityKey, token, configuration, validationParameters) =>
                IsBotFrameworkToken(token) || signingKeyValidatorUsingConfiguration(securityKey, token, configuration, validationParameters);
        }
    }

    private static bool IsBotFrameworkToken(SecurityToken token) => (token as JwtSecurityToken)?.Issuer == AuthenticationConstants.BotFrameworkTokenIssuer;
}