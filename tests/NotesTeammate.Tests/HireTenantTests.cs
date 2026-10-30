using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.AspNetCore.Builder;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Options;
using Microsoft.Identity.Abstractions;
using Microsoft.IdentityModel.JsonWebTokens;
using Microsoft.IdentityModel.Tokens;
using Microsoft.Teams.Apps;

namespace NotesTeammate.Tests;

public class HireTenantTests
{
    private const string Instance = "https://login.microsoftonline.com/";
    private const string BotIssuer = "https://api.botframework.com";
    private static readonly Guid Contoso = Guid.Parse("11111111-1111-1111-1111-111111111111");
    private static readonly Guid Fabrikam = Guid.Parse("22222222-2222-2222-2222-222222222222");

    [Fact]
    public void Token_requests_go_to_the_hire_tenant_only_inside_its_scope()
    {
        Assert.Null(HireTenantHeaders.InHireTenant(null));

        using (HireTenant.Use(Contoso))
        {
            Assert.Equal(Contoso.ToString(), HireTenantHeaders.InHireTenant(null)?.AcquireTokenOptions.Tenant);
        }

        Assert.Null(HireTenant.Current);
    }

    [Fact]
    public void A_tenant_the_caller_chose_is_kept()
    {
        var options = new AuthorizationHeaderProviderOptions { AcquireTokenOptions = new() { Tenant = Fabrikam.ToString() } };

        using (HireTenant.Use(Contoso))
        {
            Assert.Equal(Fabrikam.ToString(), HireTenantHeaders.InHireTenant(options)?.AcquireTokenOptions.Tenant);
        }
    }

    [Fact]
    public async Task Concurrent_hires_keep_their_own_tenant()
    {
        async Task<string?> TenantSeenBy(Guid tenant)
        {
            using (HireTenant.Use(tenant))
            {
                await Task.Yield();
                return HireTenantHeaders.InHireTenant(null)?.AcquireTokenOptions.Tenant;
            }
        }

        var seen = await Task.WhenAll(TenantSeenBy(Contoso), TenantSeenBy(Fabrikam));

        Assert.Equal(new[] { Contoso.ToString(), Fabrikam.ToString() }, (IEnumerable<string?>)seen);
    }

    [Theory]
    [InlineData(BotIssuer)]
    [InlineData("https://login.microsoftonline.com/11111111-1111-1111-1111-111111111111/v2.0")]
    [InlineData("https://sts.windows.net/11111111-1111-1111-1111-111111111111/")]
    public void Issuers_for_the_token_tenant_are_accepted(string issuer) =>
        Assert.Equal(issuer, HireTenant.ValidateIssuer(issuer, Contoso.ToString(), Instance, BotIssuer));

    [Theory]
    [InlineData("https://login.microsoftonline.com/22222222-2222-2222-2222-222222222222/v2.0", "11111111-1111-1111-1111-111111111111")]
    [InlineData("https://login.microsoftonline.com/common/v2.0", "common")]
    [InlineData("https://evil.example/11111111-1111-1111-1111-111111111111/v2.0", "11111111-1111-1111-1111-111111111111")]
    [InlineData("https://login.microsoftonline.com/11111111-1111-1111-1111-111111111111/v2.0", null)]
    public void Other_issuers_are_rejected(string issuer, string? tenantId) =>
        Assert.Throws<SecurityTokenInvalidIssuerException>(() => HireTenant.ValidateIssuer(issuer, tenantId, Instance, BotIssuer));

    [Fact]
    public void The_Teams_SDK_is_wired_to_serve_other_tenants()
    {
        var builder = WebApplication.CreateBuilder();
        builder.Configuration.AddInMemoryCollection(new Dictionary<string, string?>
        {
            ["AzureAd:Instance"] = Instance,
            ["AzureAd:TenantId"] = Contoso.ToString(),
            ["AzureAd:ClientId"] = Guid.NewGuid().ToString(),
        });
        builder.Services.AddTeamsBotApplication();
        builder.Services.AddHireTenants(builder.Configuration);
        using var app = builder.Build();

        Assert.IsType<HireTenantHeaders>(app.Services.GetRequiredService<IAuthorizationHeaderProvider>());

        // A request from a hire in Fabrikam, while the blueprint's home tenant is Contoso.
        var issuer = $"{Instance}{Fabrikam}/v2.0";
        var token = new JsonWebToken("""{"alg":"none"}""", $$"""{"iss":"{{issuer}}","tid":"{{Fabrikam}}"}""");
        var validation = app.Services.GetRequiredService<IOptionsMonitor<JwtBearerOptions>>().Get("AzureAd").TokenValidationParameters;
        Assert.Equal(issuer, validation.IssuerValidator(issuer, token, validation));
    }
}
