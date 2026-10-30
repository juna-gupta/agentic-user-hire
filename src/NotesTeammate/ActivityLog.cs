using System.Text.Json;

namespace NotesTeammate;

/// <summary>
/// Development only: logs the type and name of every activity Teams posts, before the SDK routes it. Handy for seeing
/// the order of lifecycle events during a hire.
/// </summary>
public static class ActivityLog
{
    public static IApplicationBuilder UseActivityLog(this IApplicationBuilder app) => app.Use(async (context, next) =>
    {
        if (HttpMethods.IsPost(context.Request.Method) && context.Request.Path.StartsWithSegments("/api/messages"))
        {
            context.Request.EnableBuffering();
            try
            {
                using var activity = await JsonDocument.ParseAsync(context.Request.Body, cancellationToken: context.RequestAborted);
                var root = activity.RootElement;
                context.RequestServices.GetRequiredService<ILoggerFactory>().CreateLogger("NotesTeammate.Activity").LogInformation(
                    "Activity {Type} {Name} {Detail} to {Recipient} in tenant {Tenant}",
                    Text(root, "type"),
                    Text(root, "name") ?? "-",
                    Detail(root),
                    root.TryGetProperty("recipient", out var recipient) ? Text(recipient, "name") ?? Text(recipient, "id") : null,
                    root.TryGetProperty("conversation", out var conversation) ? Text(conversation, "tenantId") : null);
            }
            catch (JsonException)
            {
                // Not JSON. The SDK will reject it.
            }

            context.Request.Body.Position = 0;
        }

        await next(context);
    });

    private static readonly HashSet<string> IdFields = new(StringComparer.OrdinalIgnoreCase)
    {
        "eventType", "tenantId", "agenticUserId", "agenticAppInstanceId", "agentIdentityBlueprintId", "version",
    };

    // Lifecycle events all have the name "agentLifecycle"; which change it is lives in value.eventType. Anything
    // else in the value (updatedProperty, deletionReason, workload state, manager) is appended.
    private static string Detail(JsonElement root)
    {
        if (root.TryGetProperty("value", out var value) && value.ValueKind == JsonValueKind.Object
            && value.TryGetProperty("eventType", out var eventType))
        {
            var extra = value.EnumerateObject()
                .Where(p => !IdFields.Contains(p.Name) && p.Value.ValueKind != JsonValueKind.Null)
                .Select(p => $"{p.Name}={Shorten(p.Value.ToString())}");
            return string.Join(' ', extra.Prepend(eventType.ToString()));
        }

        var added = root.TryGetProperty("membersAdded", out var a) && a.ValueKind == JsonValueKind.Array ? a.GetArrayLength() : 0;
        var removed = root.TryGetProperty("membersRemoved", out var r) && r.ValueKind == JsonValueKind.Array ? r.GetArrayLength() : 0;
        return added + removed > 0 ? $"(members added {added}, removed {removed})" : "-";
    }

    private static string Shorten(string text) => text.Length <= 120 ? text : text[..120] + "...";

    private static string? Text(JsonElement element, string property) =>
        element.ValueKind == JsonValueKind.Object && element.TryGetProperty(property, out var value) && value.ValueKind == JsonValueKind.String
            ? value.GetString()
            : null;
}
