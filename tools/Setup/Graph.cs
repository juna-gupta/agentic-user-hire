using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Text.Json.Nodes;
using Azure.Core;
using Azure.Identity;

namespace Setup;

/// <summary>Microsoft Graph calls as the signed-in admin.</summary>
internal sealed class Graph(TokenCredential credential)
{
    // Microsoft Graph Command Line Tools, a public client available in every tenant.
    private const string ClientId = "14d82eec-204b-4c2f-b7e8-296a70dab67e";

    private static readonly string[] Scopes = new[]
    {
        "User.Read",
        "Application.Read.All",
        "AgentIdentityBlueprint.Create",
        "AgentIdentityBlueprint.AddRemoveCreds.All",
        "AgentIdentityBlueprint.DeleteRestore.All",
        "AgentIdentityBlueprintPrincipal.Create",
        "AgentIdentity.Create.All",
        "AgentIdentity.DeleteRestore.All",
        "AgentIdUser.ReadWrite.All",
        "AgentCommunicationConfiguration.ReadWrite",
        "DelegatedPermissionGrant.ReadWrite.All",
        "User.ReadWrite.All",
        "LicenseAssignment.ReadWrite.All",
    }.Select(scope => "https://graph.microsoft.com/" + scope).ToArray();

    private static readonly HttpClient Http = new() { BaseAddress = new Uri("https://graph.microsoft.com/") };

    /// <summary>
    /// Signs in once, then reuses the cached sign-in on later runs. Uses the browser, or a code to enter on any device
    /// when <paramref name="deviceCode"/> is set.
    /// </summary>
    public static async Task<Graph> SignInAsync(string? tenant, string recordPath, bool deviceCode = false)
    {
        var cache = new TokenCachePersistenceOptions { Name = "agentic-user-hire-setup" };
        AuthenticationRecord? saved = null;
        if (File.Exists(recordPath))
        {
            await using var stream = File.OpenRead(recordPath);
            saved = await AuthenticationRecord.DeserializeAsync(stream);
        }

        TokenCredential credential;
        Func<Task<AuthenticationRecord>> authenticate;
        if (deviceCode)
        {
            var device = new DeviceCodeCredential(new DeviceCodeCredentialOptions
            {
                ClientId = ClientId,
                TenantId = tenant ?? "organizations",
                TokenCachePersistenceOptions = cache,
                AuthenticationRecord = saved,
                DeviceCodeCallback = (code, _) =>
                {
                    Console.WriteLine(code.Message);
                    return Task.CompletedTask;
                },
            });
            (credential, authenticate) = (device, () => device.AuthenticateAsync(new TokenRequestContext(Scopes)));
        }
        else
        {
            var browser = new InteractiveBrowserCredential(new InteractiveBrowserCredentialOptions
            {
                ClientId = ClientId,
                TenantId = tenant ?? "organizations",
                TokenCachePersistenceOptions = cache,
                AuthenticationRecord = saved,
            });
            (credential, authenticate) = (browser, () => browser.AuthenticateAsync(new TokenRequestContext(Scopes)));
        }

        if (saved is null)
        {
            var record = await authenticate();
            await using var stream = File.Create(recordPath);
            await record.SerializeAsync(stream);
        }

        return new Graph(credential);
    }

    public Task<JsonNode?> GetAsync(string path) => SendAsync(HttpMethod.Get, path);

    public Task<JsonNode?> PostAsync(string path, JsonNode body, bool retry = false) => SendAsync(HttpMethod.Post, path, body, retry);

    public Task<JsonNode?> PutAsync(string path, JsonNode body, bool retry = false) => SendAsync(HttpMethod.Put, path, body, retry);

    public Task<JsonNode?> PatchAsync(string path, JsonNode body) => SendAsync(HttpMethod.Patch, path, body);

    public Task<JsonNode?> DeleteAsync(string path) => SendAsync(HttpMethod.Delete, path);

    /// <param name="retry">Retry for about a minute. New directory objects take a little while to show up everywhere.</param>
    private async Task<JsonNode?> SendAsync(HttpMethod method, string path, JsonNode? body = null, bool retry = false)
    {
        for (var attempt = 1; ; attempt++)
        {
            var token = await credential.GetTokenAsync(new TokenRequestContext(Scopes), default);
            using var request = new HttpRequestMessage(method, path);
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token.Token);
            request.Headers.Add("OData-Version", "4.0");
            if (body is not null)
            {
                request.Content = JsonContent.Create(body);
            }

            using var response = await Http.SendAsync(request);
            var text = await response.Content.ReadAsStringAsync();
            if (response.IsSuccessStatusCode)
            {
                return text.Length > 0 ? JsonNode.Parse(text) : null;
            }

            if (retry && attempt < 7 && (int)response.StatusCode is 400 or 403 or 404)
            {
                Console.WriteLine($"  waiting for the directory to catch up ({(int)response.StatusCode})");
                await Task.Delay(TimeSpan.FromSeconds(10));
                continue;
            }

            throw new GraphException($"{method} {path} failed with {(int)response.StatusCode}: {text}");
        }
    }
}

internal sealed class GraphException(string message) : Exception(message);
