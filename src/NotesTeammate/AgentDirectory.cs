using System.Text.Json;
using Microsoft.Teams.Core.Schema;
using ModelContextProtocol.Protocol;

namespace NotesTeammate;

/// <summary>Reads an instance's own profile through Work IQ, signed in as that instance's agent user.</summary>
public sealed class AgentDirectory(WorkIq workIq, ILogger<AgentDirectory> logger)
{
    private const string MePath = "/me?$select=displayName,userPrincipalName";
    private static readonly TimeSpan Timeout = TimeSpan.FromSeconds(30);

    public async Task<AgentProfile?> GetProfileAsync(AgenticIdentity identity, CancellationToken cancellationToken = default)
    {
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        timeout.CancelAfter(Timeout);
        try
        {
            await using var session = await workIq.ConnectAsync(identity, timeout.Token);
            if (session is null)
            {
                return null;
            }

            var result = await session.CallAsync("fetch", new Dictionary<string, object?> { ["entityUrls"] = new[] { MePath } }, timeout.Token);
            var profile = ReadProfile(result);
            if (profile is null)
            {
                logger.LogWarning("Work IQ didn't return the profile of agent user {AgentUser}: {Result}", identity.AgenticUserId, WorkIq.Describe(result));
            }

            return profile;
        }
        catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
        {
            logger.LogWarning("Work IQ didn't answer within {Seconds} seconds for agent user {AgentUser}", Timeout.TotalSeconds, identity.AgenticUserId);
            return null;
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            logger.LogWarning(ex, "Couldn't read the profile of agent user {AgentUser}", identity.AgenticUserId);
            return null;
        }
    }

    // Work IQ's fetch answers {"results":[{"statusCode":200,"data":{...}}]}, one result per path, as structured
    // content or as a JSON text block.
    internal static AgentProfile? ReadProfile(CallToolResult result)
    {
        if (WorkIq.Json(result) is not { } envelope || WorkIq.Items(envelope, "results").FirstOrDefault() is not { ValueKind: JsonValueKind.Object } first
            || !first.TryGetProperty("statusCode", out var status)
            || !status.TryGetInt32(out var code)
            || code != 200
            || !first.TryGetProperty("data", out var data)
            || data.ValueKind != JsonValueKind.Object)
        {
            return null;
        }

        return new AgentProfile(WorkIq.Text(data, "displayName"), WorkIq.Text(data, "userPrincipalName"));
    }
}
