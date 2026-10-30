using System.Security.Claims;
using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.Identity.Abstractions;
using Microsoft.IdentityModel.JsonWebTokens;
using Microsoft.IdentityModel.Tokens;

namespace NotesTeammate;

/// <summary>
/// The tenant of the hire being served. A hire's agent identity and agent user live in the customer's tenant, so its
/// tokens have to be requested there, not in the tenant where the blueprint was registered.
/// </summary>
public static class HireTenant
{
    private static readonly AsyncLocal<string?> s_current = new();

    public static string? Current => s_current.Value;

    /// <summary>Requests agent user tokens in <paramref name="tenantId"/> until the result is disposed.</summary>
    public static IDisposable Use(Guid tenantId)
    {
        var previous = s_current.Value;
        s_current.Value = tenantId.ToString("D");
        return new Restore(previous);
    }

    /// <summary>
    /// Lets hires in any tenant use this service: agent user tokens go to the hire's tenant, and requests signed by a
    /// customer tenant are accepted. The audience (this blueprint) and the caller (the Agent 365 platform) are still
    /// checked by the Teams SDK. Call after <c>AddTeamsBotApplication</c>.
    /// </summary>
    public static IServiceCollection AddHireTenants(this IServiceCollection services, IConfiguration configuration)
    {
        var original = services.LastOrDefault(d => d.ServiceType == typeof(IAuthorizationHeaderProvider))
            ?? throw new InvalidOperationException("Call AddTeamsBotApplication first.");
        services.Remove(original);
        services.Add(new ServiceDescriptor(
            typeof(IAuthorizationHeaderProvider),
            sp => new HireTenantHeaders(Create(sp, original)),
            original.Lifetime));

        var instance = configuration["AzureAd:Instance"] ?? "https://login.microsoftonline.com/";
        var botTokenIssuer = configuration["BotFramework:BotTokenIssuer"] ?? "https://api.botframework.com";
        services.PostConfigure<JwtBearerOptions>("AzureAd", options =>
            options.TokenValidationParameters.IssuerValidator = (issuer, token, _) =>
                ValidateIssuer(issuer, TenantOf(token), instance, botTokenIssuer));

        return services;
    }

    /// <summary>Accepts Bot Framework tokens, and Entra tokens whose issuer is the token's own tenant.</summary>
    internal static string ValidateIssuer(string issuer, string? tenantId, string instance, string botTokenIssuer)
    {
        if (string.Equals(issuer, botTokenIssuer, StringComparison.OrdinalIgnoreCase))
        {
            return issuer;
        }

        if (Guid.TryParse(tenantId, out var tenant)
            && (issuer == $"{instance.TrimEnd('/')}/{tenant:D}/v2.0" || issuer == $"https://sts.windows.net/{tenant:D}/"))
        {
            return issuer;
        }

        throw new SecurityTokenInvalidIssuerException($"Issuer '{issuer}' doesn't match tenant '{tenantId}'.");
    }

    private static string? TenantOf(SecurityToken token) =>
        token is JsonWebToken jwt && jwt.TryGetPayloadValue<string>("tid", out var tid) ? tid : null;

    private static IAuthorizationHeaderProvider Create(IServiceProvider services, ServiceDescriptor descriptor) =>
        (IAuthorizationHeaderProvider)(descriptor.ImplementationInstance
            ?? descriptor.ImplementationFactory?.Invoke(services)
            ?? ActivatorUtilities.CreateInstance(services, descriptor.ImplementationType!));

    private sealed class Restore(string? previous) : IDisposable
    {
        public void Dispose() => s_current.Value = previous;
    }
}

/// <summary>Adds <see cref="HireTenant.Current"/> to agent user token requests that don't name a tenant.</summary>
internal sealed class HireTenantHeaders(IAuthorizationHeaderProvider inner) : IAuthorizationHeaderProvider
{
    public Task<string> CreateAuthorizationHeaderAsync(IEnumerable<string> scopes, AuthorizationHeaderProviderOptions? options = null,
        ClaimsPrincipal? claimsPrincipal = null, CancellationToken cancellationToken = default) =>
        inner.CreateAuthorizationHeaderAsync(scopes, InHireTenant(options), claimsPrincipal, cancellationToken);

    public Task<string> CreateAuthorizationHeaderForUserAsync(IEnumerable<string> scopes, AuthorizationHeaderProviderOptions? options = null,
        ClaimsPrincipal? claimsPrincipal = null, CancellationToken cancellationToken = default) =>
        inner.CreateAuthorizationHeaderForUserAsync(scopes, InHireTenant(options), claimsPrincipal, cancellationToken);

    public Task<string> CreateAuthorizationHeaderForAppAsync(string scopes, AuthorizationHeaderProviderOptions? options = null,
        CancellationToken cancellationToken = default) =>
        inner.CreateAuthorizationHeaderForAppAsync(scopes, options, cancellationToken);

    internal static AuthorizationHeaderProviderOptions? InHireTenant(AuthorizationHeaderProviderOptions? options)
    {
        if (HireTenant.Current is not { } tenant)
        {
            return options;
        }

        options ??= new AuthorizationHeaderProviderOptions();
        options.AcquireTokenOptions ??= new AcquireTokenOptions();
        if (string.IsNullOrEmpty(options.AcquireTokenOptions.Tenant))
        {
            options.AcquireTokenOptions.Tenant = tenant;
        }

        return options;
    }
}
