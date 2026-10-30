using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using Microsoft.Teams.Core.Schema;

namespace NotesTeammate;

/// <summary>Lines that answer a question, and the file or chat they came from.</summary>
public sealed record Finding(string Source, string? Url, IReadOnlyList<string> Lines);

/// <summary>Says what the hire is doing while it looks for an answer.</summary>
public delegate Task Step(string text, CancellationToken cancellationToken);

/// <summary>
/// Answers a question from what people shared with a hire, through Work IQ and without a language model. Work IQ's
/// retrieve searches as the hire's agent user, so a hire only finds files, messages and meetings that were shared
/// with it. The hire reads the files it finds and quotes the lines that match the question. When retrieve finds
/// nothing, the hire reads this chat and the files attached to it.
/// </summary>
public sealed partial class AnswerFinder(WorkIq workIq, ILogger<AnswerFinder> logger)
{
    private const int MaxFiles = 3;
    private const int MaxMessages = 50;
    private const string Retrieval = "application/vnd.ms-workiq.retrieval";
    private static readonly TimeSpan Timeout = TimeSpan.FromSeconds(60);

    /// <param name="chatId">The Teams chat the question was asked in, searched when retrieve finds nothing.</param>
    /// <param name="name">The hire's display name. Its own messages in a chat that retrieve finds are skipped.</param>
    /// <param name="step">Told what the hire is doing, such as "Reading Kestrel notes", so Teams can show it.</param>
    /// <returns>What was found, empty when nothing matched, or null when Work IQ couldn't be used.</returns>
    public async Task<IReadOnlyList<Finding>?> FindAsync(
        AgenticIdentity identity, string question, string? chatId, string? name = null, Step? step = null, CancellationToken cancellationToken = default)
    {
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        timeout.CancelAfter(Timeout);
        try
        {
            await using var session = await workIq.ConnectAsync(identity, timeout.Token);
            return session is null ? null : await FindAsync(session, question, chatId, identity.AgenticUserId, name, step, timeout.Token);
        }
        catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
        {
            logger.LogWarning("Work IQ didn't answer within {Seconds} seconds for agent user {AgentUser}", Timeout.TotalSeconds, identity.AgenticUserId);
            return null;
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            logger.LogWarning(ex, "Couldn't search for agent user {AgentUser}", identity.AgenticUserId);
            return null;
        }
    }

    /// <param name="self">The hire's agent user ID. Its own replies in the chat are skipped.</param>
    /// <param name="name">The hire's display name, used to leave its own messages out of a chat snippet.</param>
    internal static async Task<IReadOnlyList<Finding>> FindAsync(
        IWorkIqTools tools, string question, string? chatId, string? self, string? name, Step? step, CancellationToken cancellationToken)
    {
        var words = TextMatch.Words(question);
        if (words.Count == 0)
        {
            return [];
        }

        step ??= (_, _) => Task.CompletedTask;
        await step("Searching what's shared with me", cancellationToken);
        var found = await SearchAsync(tools, question, words, name, step, cancellationToken);
        if (found.Count > 0 || chatId is null)
        {
            return found;
        }

        await step("Nothing shared with me matched, so reading this chat", cancellationToken);
        return await ReadChatAsync(tools, chatId, words, self, step, cancellationToken);
    }

    // retrieve returns {"application/vnd.ms-workiq.retrieval":{"markdown":"...","retrievalHits":[...]}}. Each hit has
    // a link and, for files, the drive item to read. The Markdown has each hit's title and a snippet.
    private static async Task<List<Finding>> SearchAsync(IWorkIqTools tools, string question, HashSet<string> words, string? name, Step step, CancellationToken cancellationToken)
    {
        var result = await tools.CallAsync("retrieve", new Dictionary<string, object?> { ["query"] = question }, cancellationToken);
        if (WorkIq.Json(result) is not { } envelope || !envelope.TryGetProperty(Retrieval, out var retrieval))
        {
            throw new InvalidOperationException("Work IQ retrieve failed: " + WorkIq.Describe(result));
        }

        var summaries = Summaries(WorkIq.Text(retrieval, "markdown"));
        var found = new List<Finding>();
        var files = 0;
        foreach (var hit in WorkIq.Items(retrieval, "retrievalHits"))
        {
            summaries.TryGetValue(WorkIq.Text(hit, "id") ?? string.Empty, out var summary);
            var title = Title(hit) ?? summary.Title ?? "a shared item";
            List<string>? lines = null;
            if (files < MaxFiles && WorkIq.Text(hit, "entityUrl") is { } item && DriveItem().IsMatch(item))
            {
                files++;
                await step($"Reading {title}", cancellationToken);
                lines = await ReadFileAsync(tools, item + "/content", cancellationToken);
            }

            lines ??= TextMatch.Snippet(summary.Snippet, name);
            if (TextMatch.Pick(lines, words, title) is { Count: > 0 } picked)
            {
                found.Add(new Finding(title, WorkIq.Text(hit, "webUrl"), picked));
            }
        }

        return found;
    }

    // The chat's latest messages, newest first. Questions and the hire's own replies are skipped, and files attached
    // to a message are read through their sharing URL.
    private static async Task<List<Finding>> ReadChatAsync(IWorkIqTools tools, string chatId, HashSet<string> words, string? self, Step step, CancellationToken cancellationToken)
    {
        var path = $"/chats/{chatId}/messages?$top={MaxMessages}";
        var result = await tools.CallAsync("fetch", new Dictionary<string, object?> { ["entityUrls"] = new[] { path } }, cancellationToken);
        if (WorkIq.Json(result) is not { } envelope
            || WorkIq.Items(envelope, "results").FirstOrDefault() is not { ValueKind: JsonValueKind.Object } first
            || !first.TryGetProperty("data", out var data))
        {
            throw new InvalidOperationException("Work IQ couldn't read the chat: " + WorkIq.Describe(result));
        }

        var said = new List<string>();
        var found = new List<Finding>();
        var files = 0;
        foreach (var message in WorkIq.Items(data, "value"))
        {
            if (WorkIq.Text(message, "messageType") != "message"
                || (message.TryGetProperty("from", out var from) && from.ValueKind == JsonValueKind.Object
                    && from.TryGetProperty("user", out var user) && string.Equals(WorkIq.Text(user, "id"), self, StringComparison.OrdinalIgnoreCase)))
            {
                continue;
            }

            if (message.TryGetProperty("body", out var body))
            {
                said.AddRange(TextMatch.Lines(TextMatch.FromHtml(WorkIq.Text(body, "content"))).Select(TextMatch.WithoutRemember));
            }

            foreach (var attachment in WorkIq.Items(message, "attachments"))
            {
                if (files >= MaxFiles || WorkIq.Text(attachment, "contentType") != "reference"
                    || WorkIq.Text(attachment, "contentUrl") is not { } url || !url.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                files++;
                var name = WorkIq.Text(attachment, "name") ?? "a file in this chat";
                await step($"Reading {name} from this chat", cancellationToken);
                if (await ReadFileAsync(tools, $"/shares/{ShareId(url)}/driveItem/content", cancellationToken) is { } lines
                    && TextMatch.Pick(lines, words, name) is { Count: > 0 } picked)
                {
                    found.Add(new Finding(name, url, picked));
                }
            }
        }

        if (TextMatch.Pick(said, words, title: null) is { Count: > 0 } fromChat)
        {
            found.Insert(0, new Finding("This chat", null, fromChat));
        }

        return found;
    }

    // fetch_blob answers {"statusCode":200,"blobName":"...","contentType":"...","base64Content":"..."}.
    private static async Task<List<string>?> ReadFileAsync(IWorkIqTools tools, string path, CancellationToken cancellationToken)
    {
        var result = await tools.CallAsync("fetch_blob", new Dictionary<string, object?> { ["path"] = path }, cancellationToken);
        if (WorkIq.Json(result) is not { } blob
            || !blob.TryGetProperty("statusCode", out var status) || !status.TryGetInt32(out var code) || code != 200
            || WorkIq.Text(blob, "base64Content") is not { } content)
        {
            return null;
        }

        var text = TextMatch.ReadDocument(Convert.FromBase64String(content), WorkIq.Text(blob, "blobName"), WorkIq.Text(blob, "contentType"));
        return text is null ? null : TextMatch.Lines(text);
    }

    private static string? Title(JsonElement hit) =>
        hit.TryGetProperty("resourceMetadata", out var metadata) && metadata.ValueKind == JsonValueKind.Object
        && metadata.TryGetProperty("document", out var document) ? WorkIq.Text(document, "title") : null;

    // The Markdown lists each hit as "- **Title** [^h1]" followed by "  - **snippet:** ..." and other fields.
    internal static Dictionary<string, (string? Title, string? Snippet)> Summaries(string? markdown)
    {
        var summaries = new Dictionary<string, (string? Title, string? Snippet)>(StringComparer.Ordinal);
        string? id = null;
        foreach (var line in (markdown ?? string.Empty).Split('\n'))
        {
            if (HitHeading().Match(line) is { Success: true } heading)
            {
                id = heading.Groups["id"].Value;
                summaries[id] = (heading.Groups["title"].Value, null);
            }
            else if (id is not null && SnippetLine().Match(line) is { Success: true } snippet)
            {
                summaries[id] = (summaries[id].Title, snippet.Groups["text"].Value);
            }
        }

        return summaries;
    }

    // The sharing ID for a URL: "u!" and the URL in unpadded base64url.
    internal static string ShareId(string url) =>
        "u!" + Convert.ToBase64String(Encoding.UTF8.GetBytes(url)).TrimEnd('=').Replace('/', '_').Replace('+', '-');

    [GeneratedRegex(@"^/drives/[^/]+/items/[^/]+$")]
    private static partial Regex DriveItem();

    [GeneratedRegex(@"^- \*\*(?<title>.+?)\*\* \[\^(?<id>h\d+)\]")]
    private static partial Regex HitHeading();

    [GeneratedRegex(@"^\s+- \*\*snippet:\*\*\s*(?<text>.+)$")]
    private static partial Regex SnippetLine();
}
