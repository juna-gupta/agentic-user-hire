using System.Text.Json;
using Microsoft.Identity.Abstractions;
using Microsoft.Teams.Core.Hosting;
using Microsoft.Teams.Core.Schema;
using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;

namespace NotesTeammate;

/// <summary>Calls Work IQ tools. Tests replace it with canned answers.</summary>
public interface IWorkIqTools
{
    Task<CallToolResult> CallAsync(string tool, IReadOnlyDictionary<string, object?> arguments, CancellationToken cancellationToken);
}

/// <summary>
/// Connects to the Work IQ MCP server signed in as one hire's agent user. No person signs in: the blueprint's
/// credential is exchanged for a token for the agent identity and then for its agent user. Work IQ only shows that
/// user what it can access, and bills its calls to the hire's tenant.
/// </summary>
public sealed class WorkIq(IAuthorizationHeaderProvider headers, IHttpClientFactory httpClients)
{
    private const string Scope = "api://workiq.svc.cloud.microsoft/.default";
    private static readonly Uri Endpoint = new("https://workiq.svc.cloud.microsoft/mcp");

    private readonly BotTokenProvider _tokens = new(headers, "AzureAd");

    /// <returns>A session for one conversation turn, or null when there's no token for the agent user.</returns>
    public async Task<WorkIqSession?> ConnectAsync(AgenticIdentity identity, CancellationToken cancellationToken)
    {
        // WorkIQAgent.Ask comes from the blueprint's inheritable permissions, so a new hire needs no consent of its own.
        var token = await _tokens.GetAgenticUserTokenAsync(identity, Scope, cancellationToken);
        if (token is null)
        {
            return null;
        }

        var transport = new HttpClientTransport(new HttpClientTransportOptions
        {
            Endpoint = Endpoint,
            TransportMode = HttpTransportMode.StreamableHttp,
            AdditionalHeaders = new Dictionary<string, string> { ["Authorization"] = "Bearer " + token },
        }, httpClients.CreateClient());
        try
        {
            return new WorkIqSession(transport, await McpClient.CreateAsync(transport, cancellationToken: cancellationToken));
        }
        catch
        {
            await transport.DisposeAsync();
            throw;
        }
    }

    /// <summary>The JSON a tool returned, as structured content or as a JSON text block.</summary>
    internal static JsonElement? Json(CallToolResult result)
    {
        if (result.StructuredContent is { ValueKind: JsonValueKind.Object } structured)
        {
            return structured;
        }

        foreach (var block in result.Content.OfType<TextContentBlock>().Reverse())
        {
            try
            {
                using var document = JsonDocument.Parse(block.Text);
                if (document.RootElement.ValueKind == JsonValueKind.Object)
                {
                    return document.RootElement.Clone();
                }
            }
            catch (JsonException)
            {
            }
        }

        return null;
    }

    internal static string Describe(CallToolResult result)
    {
        var text = string.Join(' ', result.Content.OfType<TextContentBlock>().Select(block => block.Text));
        return text.Length <= 300 ? text : text[..300];
    }

    internal static string? Text(JsonElement parent, string property) =>
        parent.ValueKind == JsonValueKind.Object && parent.TryGetProperty(property, out var value) && value.ValueKind == JsonValueKind.String
            ? value.GetString()
            : null;

    internal static IEnumerable<JsonElement> Items(JsonElement parent, string property) =>
        parent.ValueKind == JsonValueKind.Object && parent.TryGetProperty(property, out var value) && value.ValueKind == JsonValueKind.Array
            ? value.EnumerateArray()
            : [];
}

/// <summary>One open Work IQ connection. Dispose it when the turn is done.</summary>
public sealed class WorkIqSession(HttpClientTransport transport, McpClient client) : IWorkIqTools, IAsyncDisposable
{
    public async Task<CallToolResult> CallAsync(string tool, IReadOnlyDictionary<string, object?> arguments, CancellationToken cancellationToken) =>
        await client.CallToolAsync(tool, arguments, cancellationToken: cancellationToken);

    public async ValueTask DisposeAsync()
    {
        await client.DisposeAsync();
        await transport.DisposeAsync();
    }
}
