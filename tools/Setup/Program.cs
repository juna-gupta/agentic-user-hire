using System.IO.Compression;
using System.Text.Json;
using System.Text.Json.Nodes;
using Microsoft.Extensions.Configuration.UserSecrets;
using Setup;

// Sets up a test tenant for the sample, signed in as an admin. Run from the repository root:
//   dotnet run --project tools/Setup -- <command>

const string WorkIqAppId = "fdcc1f02-fc51-4226-8753-f668596af7f7"; // Work IQ, for the agent user's own Microsoft 365 data
const string MessagingBotApiAppId = "5a807f24-c9de-44ee-a3a7-329e88a00ffc"; // Teams messages sent as the agent user
const string Agent365ManagerAppId = "e8be65d6-d430-4289-a665-51bf2a194bda"; // lets Agent 365 manage the blueprint's agents
const string UserSecretsId = "agentic-user-hire"; // the agent project's UserSecretsId

// What every instance may do: read its own profile through Work IQ, and send Teams messages as its agent user.
(string ResourceAppId, string Scope)[] permissions = [(WorkIqAppId, "WorkIQAgent.Ask"), (MessagingBotApiAppId, "AgentData.ReadWrite")];

var root = FindRoot();
var setupFolder = Path.Combine(root, "tools", "Setup");
var state = State.Load(Path.Combine(setupFolder, "state.json"));
var command = args.ElementAtOrDefault(0);
var argument = args.ElementAtOrDefault(1) is { } second && !second.StartsWith("--", StringComparison.Ordinal) ? second : null;

try
{
    switch (command)
    {
        case "blueprint":
            await CreateBlueprintAsync();
            break;
        case "endpoint" when argument?.StartsWith("https://", StringComparison.OrdinalIgnoreCase) == true:
            await SetEndpointAsync(argument);
            break;
        case "package":
            BuildPackage();
            break;
        case "consent" when Option("--tenant") is not null:
            await ConsentAsync();
            break;
        case "hire" when argument is not null:
            await HireAsync(argument);
            break;
        case "remove" when argument is not null:
            await RemoveAsync(argument);
            break;
        case "cleanup":
            await CleanupAsync();
            break;
        default:
            Console.WriteLine("""
                Usage: dotnet run --project tools/Setup -- <command>

                  blueprint [--tenant <id>]     Create the blueprint, its secret (saved to user secrets) and permissions
                  endpoint <https-url>          Send Teams traffic to your service, e.g. https://<tunnel>/api/messages
                  package                       Build appPackage/build/notes-teammate.zip for the admin center
                  consent --tenant <id>         As a customer tenant's admin, approve the blueprint's permissions there
                  hire <name> [--license <sku-id>] [--domain <domain>]
                                                Create an instance the way a store hire does
                  remove <name>                 Delete one instance
                  cleanup                       Delete every instance and the blueprint

                Add --device-code to sign in with a code instead of a browser window.
                """);
            return 1;
    }

    return 0;
}
catch (Exception ex) when (ex is GraphException or InvalidOperationException)
{
    Console.Error.WriteLine(ex.Message);
    return 1;
}

// The blueprint is the agent's definition in Entra. Every hire becomes an agent identity and agent user under it.
async Task CreateBlueprintAsync()
{
    var graph = await ConnectAsync();
    var me = await MeAsync(graph);

    if (state.BlueprintAppId is null)
    {
        state.TenantId = (await graph.GetAsync("v1.0/organization?$select=id"))?["value"]?[0]?["id"]?.GetValue<string>();
        var blueprint = await graph.PostAsync("beta/applications", new JsonObject
        {
            ["@odata.type"] = "Microsoft.Graph.AgentIdentityBlueprint",
            ["displayName"] = "Notes Teammate",
            ["signInAudience"] = "AzureADMultipleOrgs", // customers in other tenants can hire it
            ["managerApplications"] = new JsonArray(Agent365ManagerAppId),
            ["sponsors@odata.bind"] = new JsonArray(me.Ref),
            ["owners@odata.bind"] = new JsonArray(me.Ref),
        });
        state.BlueprintObjectId = Text(blueprint, "id");
        state.BlueprintAppId = Text(blueprint, "appId");
        state.Save();
        Console.WriteLine($"Blueprint created: {state.BlueprintAppId}");

        var password = await graph.PostAsync($"v1.0/applications/{state.BlueprintObjectId}/addPassword", new JsonObject
        {
            ["passwordCredential"] = new JsonObject
            {
                ["displayName"] = "local development",
                ["endDateTime"] = DateTimeOffset.UtcNow.AddDays(90).ToString("o"),
            },
        }, retry: true);
        SetUserSecrets(new()
        {
            ["AzureAd:TenantId"] = state.TenantId,
            ["AzureAd:ClientId"] = state.BlueprintAppId,
            ["AzureAd:ClientCredentials:0:ClientSecret"] = Text(password, "secretText"),
        });
        Console.WriteLine("Client secret saved to the agent's user secrets (expires in 90 days)");
    }

    if (state.BlueprintPrincipalId is null)
    {
        var principal = await graph.PostAsync("v1.0/servicePrincipals/microsoft.graph.agentIdentityBlueprintPrincipal",
            new JsonObject { ["appId"] = state.BlueprintAppId }, retry: true);
        state.BlueprintPrincipalId = Text(principal, "id");
        state.Save();
        Console.WriteLine("Blueprint principal created");
    }

    // Consent once on the blueprint, and mark the scopes inheritable: every instance gets them with no consent of its
    // own, including instances hired before a scope was added. Run 'blueprint' again after changing the list.
    var inheritablePath = $"beta/applications/{state.BlueprintObjectId}/microsoft.graph.agentIdentityBlueprint/inheritablePermissions";
    var inheritable = (await graph.GetAsync(inheritablePath))?["value"] as JsonArray ?? [];
    foreach (var (resourceAppId, scope) in permissions)
    {
        await GrantAsync(graph, state.BlueprintPrincipalId, resourceAppId, scope);
        var current = inheritable.FirstOrDefault(p => p?["resourceAppId"]?.GetValue<string>() == resourceAppId);
        var scopes = (current?["inheritableScopes"]?["scopes"] as JsonArray ?? []).Select(s => s?.GetValue<string>()).OfType<string>().ToList();
        if (scopes.Contains(scope, StringComparer.OrdinalIgnoreCase))
        {
            continue;
        }

        var body = new JsonObject
        {
            ["inheritableScopes"] = new JsonObject
            {
                ["@odata.type"] = "microsoft.graph.enumeratedScopes",
                ["scopes"] = new JsonArray(scopes.Append(scope).Select(s => (JsonNode?)s).ToArray()),
            },
        };
        if (current is null)
        {
            body["resourceAppId"] = resourceAppId;
            await graph.PostAsync(inheritablePath, body, retry: true);
        }
        else
        {
            await graph.PatchAsync($"{inheritablePath}/{resourceAppId}", body);
        }

        Console.WriteLine($"Made inheritable: {scope}");
    }

    Console.WriteLine("Blueprint ready. Next: start a tunnel to http://localhost:3978 and run 'endpoint https://<tunnel>/api/messages'.");
}

// Consent is per tenant. The grants made by 'blueprint' cover the home tenant only, so in a customer's tenant every
// hire fails to get a token (AADSTS65001) until that tenant's admin approves the same permissions once. The hires there
// then inherit them, just like at home.
async Task ConsentAsync()
{
    var blueprintAppId = Require(state.BlueprintAppId);
    var graph = await ConnectAsync();
    var existing = await graph.GetAsync($"v1.0/servicePrincipals?$filter=appId eq '{blueprintAppId}'&$select=id");
    var principalId = First(existing)?["id"]?.GetValue<string>();
    if (principalId is null)
    {
        var principal = await graph.PostAsync("v1.0/servicePrincipals/microsoft.graph.agentIdentityBlueprintPrincipal",
            new JsonObject { ["appId"] = blueprintAppId }, retry: true);
        principalId = Text(principal, "id");
        Console.WriteLine("Blueprint principal created in this tenant");
    }

    foreach (var (resourceAppId, scope) in permissions)
    {
        await GrantAsync(graph, principalId, resourceAppId, scope);
    }

    Console.WriteLine("Consent done. Instances hired in this tenant can now reply and read their profile through Work IQ.");
}

async Task GrantAsync(Graph graph, string? clientId, string resourceAppId, string scope)
{
    var resource = await graph.GetAsync($"v1.0/servicePrincipals?$filter=appId eq '{resourceAppId}'&$select=id");
    var resourceId = First(resource)?["id"]?.GetValue<string>()
        ?? throw new InvalidOperationException(
            $"{resourceAppId} has no service principal in this tenant. Add it with 'az ad sp create --id {resourceAppId}', then run this again.");

    var grants = await graph.GetAsync(
        $"v1.0/oauth2PermissionGrants?$filter=clientId eq '{clientId}' and resourceId eq '{resourceId}' and consentType eq 'AllPrincipals'");
    if (First(grants) is { } grant)
    {
        var scopes = (grant["scope"]?.GetValue<string>() ?? string.Empty).Split(' ', StringSplitOptions.RemoveEmptyEntries);
        if (!scopes.Contains(scope, StringComparer.OrdinalIgnoreCase))
        {
            await graph.PatchAsync($"v1.0/oauth2PermissionGrants/{Text(grant, "id")}",
                new JsonObject { ["scope"] = string.Join(' ', scopes.Append(scope)) });
        }
    }
    else
    {
        await graph.PostAsync("v1.0/oauth2PermissionGrants", new JsonObject
        {
            ["clientId"] = clientId,
            ["consentType"] = "AllPrincipals",
            ["resourceId"] = resourceId,
            ["scope"] = scope,
        }, retry: true);
    }

    Console.WriteLine($"Granted: {scope}");
}

async Task SetEndpointAsync(string url)
{
    var blueprintObjectId = Require(state.BlueprintObjectId);
    var graph = await ConnectAsync();
    await graph.PutAsync($"beta/applications/{blueprintObjectId}/microsoft.graph.agentIdentityBlueprint/communicationConfiguration", new JsonObject
    {
        ["endpointConfiguration"] = new JsonObject
        {
            ["configurationType"] = "apiBased",
            ["apiBased"] = new JsonObject { ["callbackUri"] = url },
        },
    });
    Console.WriteLine($"Teams now sends every instance's messages to {url}");
}

// The package an admin uploads in the Microsoft 365 admin center. It points the store listing at the blueprint.
void BuildPackage()
{
    var blueprintAppId = Require(state.BlueprintAppId);
    state.AppId ??= Guid.NewGuid().ToString();
    state.TemplateId ??= Guid.NewGuid().ToString();
    state.Save();

    var source = Path.Combine(root, "appPackage");
    var zipPath = Path.Combine(source, "build", "notes-teammate.zip");
    Directory.CreateDirectory(Path.GetDirectoryName(zipPath)!);
    File.Delete(zipPath);

    using var zip = ZipFile.Open(zipPath, ZipArchiveMode.Create);
    foreach (var file in new[] { "manifest.json", "agenticUserTemplateManifest.json" })
    {
        var text = File.ReadAllText(Path.Combine(source, file))
            .Replace("{{APP_ID}}", state.AppId, StringComparison.Ordinal)
            .Replace("{{TEMPLATE_ID}}", state.TemplateId, StringComparison.Ordinal)
            .Replace("{{BLUEPRINT_APP_ID}}", blueprintAppId, StringComparison.Ordinal);
        using var writer = new StreamWriter(zip.CreateEntry(file).Open());
        writer.Write(text);
    }

    foreach (var icon in new[] { "color.png", "outline.png" })
    {
        zip.CreateEntryFromFile(Path.Combine(source, icon), icon);
    }

    Console.WriteLine($"Package: {zipPath}");
}

// What the platform does when someone hires the agent from the store: a new agent identity under the blueprint, and
// an agent user for it, managed by the person who hired it. Use it where the store isn't available.
async Task HireAsync(string name)
{
    var blueprintAppId = Require(state.BlueprintAppId);
    var graph = await ConnectAsync();
    var me = await MeAsync(graph);
    var nickname = "notes-" + new string(name.ToLowerInvariant().Where(char.IsAsciiLetterOrDigit).ToArray());
    var upn = $"{nickname}@{Option("--domain") ?? me.Domain}";

    var identity = await graph.PostAsync("beta/servicePrincipals/microsoft.graph.agentIdentity", new JsonObject
    {
        ["displayName"] = $"Notes Teammate {name}",
        ["agentIdentityBlueprintId"] = blueprintAppId,
        ["sponsors@odata.bind"] = new JsonArray(me.Ref),
    }, retry: true);
    var user = await graph.PostAsync("beta/users/microsoft.graph.agentUser", new JsonObject
    {
        ["accountEnabled"] = true,
        ["displayName"] = $"{name} (Notes Teammate)",
        ["mailNickname"] = nickname,
        ["userPrincipalName"] = upn,
        ["identityParentId"] = Text(identity, "id"),
        ["usageLocation"] = "US",
    }, retry: true);
    var userId = Text(user, "id");
    state.Hires.Add(new HiredInstance(name, Text(identity, "id"), userId, upn));
    state.Save();

    await graph.PutAsync($"v1.0/users/{userId}/manager/$ref", new JsonObject { ["@odata.id"] = me.Ref }, retry: true);
    if (Option("--license") is { } sku)
    {
        await graph.PostAsync($"v1.0/users/{userId}/assignLicense", new JsonObject
        {
            ["addLicenses"] = new JsonArray(new JsonObject { ["skuId"] = sku, ["disabledPlans"] = new JsonArray() }),
            ["removeLicenses"] = new JsonArray(),
        }, retry: true);
    }

    Console.WriteLine($"Hired {name}: {upn} (agent user {userId}). Find it in Teams once its license is active.");
}

async Task RemoveAsync(string name)
{
    var hire = state.Hires.FirstOrDefault(h => string.Equals(h.Name, name, StringComparison.OrdinalIgnoreCase))
        ?? throw new InvalidOperationException($"No instance named {name}. Instances: {string.Join(", ", state.Hires.Select(h => h.Name))}");
    await DeleteHireAsync(await ConnectAsync(), hire);
}

async Task CleanupAsync()
{
    if (state.BlueprintObjectId is null && state.Hires.Count == 0)
    {
        Console.WriteLine("Nothing to clean up.");
        return;
    }

    var graph = await ConnectAsync();
    foreach (var hire in state.Hires.ToList())
    {
        await DeleteHireAsync(graph, hire);
    }

    if (state.BlueprintObjectId is not null)
    {
        await graph.DeleteAsync($"beta/applications/{state.BlueprintObjectId}/microsoft.graph.agentIdentityBlueprint");
        SetUserSecrets(new()
        {
            ["AzureAd:TenantId"] = null,
            ["AzureAd:ClientId"] = null,
            ["AzureAd:ClientCredentials:0:ClientSecret"] = null,
        });
        Console.WriteLine("Deleted the blueprint and cleared the agent's user secrets");
    }

    File.Delete(Path.Combine(setupFolder, "state.json"));
}

async Task DeleteHireAsync(Graph graph, HiredInstance hire)
{
    await graph.DeleteAsync($"v1.0/users/{hire.AgentUserId}");
    await graph.DeleteAsync($"beta/servicePrincipals/{hire.AgentIdentityId}/microsoft.graph.agentIdentity");
    state.Hires.Remove(hire);
    state.Save();
    Console.WriteLine($"Removed {hire.Name} ({hire.UserPrincipalName})");
}

Task<Graph> ConnectAsync()
{
    // One saved sign-in per tenant, so the blueprint's tenant and a customer's tenant can use different admins.
    var tenant = Option("--tenant") ?? state.TenantId;
    var home = tenant is null || state.TenantId is null || tenant == state.TenantId;
    var record = home ? "auth-record.json" : $"auth-record-{tenant}.json";
    return Graph.SignInAsync(tenant, Path.Combine(setupFolder, record), deviceCode: args.Contains("--device-code"));
}

async Task<(string Ref, string Domain)> MeAsync(Graph graph)
{
    var me = await graph.GetAsync("v1.0/me?$select=id,userPrincipalName");
    return ($"https://graph.microsoft.com/v1.0/users/{Text(me, "id")}", Text(me, "userPrincipalName").Split('@')[1]);
}

void SetUserSecrets(Dictionary<string, string?> values)
{
    var path = PathHelper.GetSecretsPathFromSecretsId(UserSecretsId);
    Directory.CreateDirectory(Path.GetDirectoryName(path)!);
    var secrets = File.Exists(path) ? JsonNode.Parse(File.ReadAllText(path))?.AsObject() ?? [] : [];
    foreach (var (key, value) in values)
    {
        if (value is null)
        {
            secrets.Remove(key);
        }
        else
        {
            secrets[key] = value;
        }
    }

    File.WriteAllText(path, secrets.ToJsonString(new JsonSerializerOptions { WriteIndented = true }));
}

string? Option(string name) => Array.IndexOf(args, name) is var i and >= 0 ? args.ElementAtOrDefault(i + 1) : null;

static string Require(string? value) =>
    value ?? throw new InvalidOperationException("Run 'blueprint' first.");

// The first item of a Graph list response, or null when the list is empty.
static JsonNode? First(JsonNode? list) => list?["value"] is JsonArray { Count: > 0 } items ? items[0] : null;

static string Text(JsonNode? node, string property) =>
    node?[property]?.GetValue<string>() ?? throw new InvalidOperationException($"Graph response had no '{property}'.");

static string FindRoot()
{
    for (var folder = new DirectoryInfo(Directory.GetCurrentDirectory()); folder is not null; folder = folder.Parent)
    {
        if (File.Exists(Path.Combine(folder.FullName, "AgenticUserHire.slnx")))
        {
            return folder.FullName;
        }
    }

    throw new InvalidOperationException("Run this from inside the repository.");
}
