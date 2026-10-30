using System.Globalization;
using System.Text;

namespace NotesTeammate;

/// <summary>The agent user's own directory profile, read with that instance's token.</summary>
public sealed record AgentProfile(string? DisplayName, string? UserPrincipalName);

/// <summary>A file or chat an answer quotes. Teams shows it as a numbered citation.</summary>
public sealed record Citation(int Position, string Title, string? Url, string Abstract, IReadOnlyList<string> Keywords);

/// <summary>What a hire says back, and the extras Teams can show with it.</summary>
public sealed record Reply(string Text)
{
    /// <summary>The sources an answer quotes, marked in the text as [1], [2] and so on.</summary>
    public IReadOnlyList<Citation> Citations { get; init; } = [];

    /// <summary>Messages to try next, offered as buttons in a 1:1 chat.</summary>
    public IReadOnlyList<string> Suggestions { get; init; } = [];

    /// <summary>The reply answers a question from what the hire found, so it gets the AI label and feedback buttons.</summary>
    public bool Answer { get; init; }

    /// <summary>A note was saved, so the message that asked for it gets a check mark.</summary>
    public bool Noted { get; init; }
}

/// <summary>
/// What one hired instance does: it remembers notes, tells you who it is and answers questions from what's shared
/// with it. Every method takes the instance key, and everything the instance knows lives in that instance's record,
/// so two hires never share notes, even in the same chat.
/// </summary>
public sealed class Teammate(IInstanceStore store, TimeProvider clock)
{
    public const string Help =
        "Try \"remember <something>\", \"notes\", \"who are you\", or ask about a file you shared with me (end with a question mark).";

    internal const string Remember = "remember ";

    private static readonly string[] Starters = ["notes", "who are you"];

    /// <summary>The instance was hired. <paramref name="hiredBy"/> is its manager, the person who hired it.</summary>
    public Task HiredAsync(InstanceKey key, string? hiredBy, CancellationToken cancellationToken = default) =>
        store.UpdateAsync(key, record =>
        {
            record.HiredBy = hiredBy ?? record.HiredBy;
            record.HiredAt ??= clock.GetUtcNow();
            record.Enabled = true;
        }, cancellationToken);

    /// <summary>An admin turned the agent user off or back on.</summary>
    public Task SetEnabledAsync(InstanceKey key, bool enabled, CancellationToken cancellationToken = default) =>
        store.UpdateAsync(key, record => record.Enabled = enabled, cancellationToken);

    /// <summary>The agent user was deleted, so the hire is over. Its data goes with it; only the removal is remembered.</summary>
    public Task RemovedAsync(InstanceKey key, CancellationToken cancellationToken = default) =>
        store.UpdateAsync(key, record =>
        {
            record.RemovedAt ??= clock.GetUtcNow();
            record.HiredBy = null;
            record.HiredAt = null;
            record.Greeted = false;
            record.Notes.Clear();
        }, cancellationToken);

    /// <summary>
    /// An admin restored the deleted agent user. Its notes were deleted with it, so it starts over. A restored hire is
    /// back in service, so this also turns the instance on.
    /// </summary>
    public Task RestoredAsync(InstanceKey key, CancellationToken cancellationToken = default) =>
        store.UpdateAsync(key, record =>
        {
            record.RemovedAt = null;
            record.Enabled = true;
        }, cancellationToken);

    /// <summary>Answers a message sent to this instance, or returns null when it is paused or removed.</summary>
    /// <param name="profile">Reads the instance's own profile. Only called for "who are you".</param>
    /// <param name="find">Searches what's shared with the instance. Only called for questions.</param>
    public async Task<Reply?> ReplyAsync(
        InstanceKey key,
        string text,
        string? from,
        Func<CancellationToken, Task<AgentProfile?>> profile,
        Func<string, CancellationToken, Task<IReadOnlyList<Finding>?>>? find = null,
        CancellationToken cancellationToken = default)
    {
        text = text.Trim();
        var firstContact = false;

        // Don't rely on handling the hire event first: the service may have been down when it was sent. The first
        // message also starts the record.
        var record = await store.UpdateAsync(key, record =>
        {
            if (!record.Active)
            {
                return;
            }

            record.HiredAt ??= clock.GetUtcNow();
            firstContact = !record.Greeted;
            record.Greeted = true;
            if (text.StartsWith(Remember, StringComparison.OrdinalIgnoreCase) && text[Remember.Length..].Trim() is { Length: > 0 } note)
            {
                record.Notes.Add(new Note(note, from, clock.GetUtcNow()));
            }
        }, cancellationToken);

        // Check the state on every message, not only in the lifecycle handlers. A paused or deleted instance stays
        // silent.
        if (!record.Active)
        {
            return null;
        }

        var reply = new StringBuilder();
        if (firstContact)
        {
            reply.Append(CultureInfo.InvariantCulture, $"Hi {from ?? "there"}, I'm your new notes teammate");
            reply.Append(record.HiredBy is { } hiredBy ? $", hired by {hiredBy}. " : ". ");
            reply.Append("I keep notes for this hire only.\n\n");
        }

        if (text.StartsWith(Remember, StringComparison.OrdinalIgnoreCase))
        {
            reply.Append(CultureInfo.InvariantCulture, $"Got it. I have {Count(record.Notes.Count)}.");
            return new Reply(reply.ToString()) { Noted = true, Suggestions = ["notes"] };
        }

        if (text.Equals("notes", StringComparison.OrdinalIgnoreCase) || text.StartsWith("what do you know", StringComparison.OrdinalIgnoreCase))
        {
            AppendNotes(reply, record);
        }
        else if (text.StartsWith("who are you", StringComparison.OrdinalIgnoreCase))
        {
            AppendIdentity(reply, record, await profile(cancellationToken));
        }
        else if (text.EndsWith('?') && find is not null)
        {
            return Answer(reply, text, await find(text, cancellationToken));
        }
        else
        {
            reply.Append(Help);
            return new Reply(reply.ToString()) { Suggestions = Starters };
        }

        return new Reply(reply.ToString()) { Suggestions = firstContact ? Starters : [] };
    }

    private static void AppendNotes(StringBuilder reply, InstanceRecord record)
    {
        if (record.Notes.Count == 0)
        {
            reply.Append("I don't have any notes yet. ").Append(Help);
            return;
        }

        reply.Append("Here's what I know:\n");
        foreach (var note in record.Notes)
        {
            reply.Append(CultureInfo.InvariantCulture, $"\n- {note.Text}{(note.From is null ? string.Empty : $" (from {note.From})")}");
        }
    }

    private static void AppendIdentity(StringBuilder reply, InstanceRecord record, AgentProfile? me)
    {
        reply.Append(me?.DisplayName is { } name
            ? $"I'm {name}{(me.UserPrincipalName is { } upn ? $" ({upn})" : string.Empty)}."
            : "I couldn't read my own profile. Check that Work IQ is on in this tenant and that its permission was approved.");
        reply.Append(CultureInfo.InvariantCulture,
            $"\n\n{record.HiredBy ?? "Someone"} hired me on {record.HiredAt:MMM d}. I have {Count(record.Notes.Count)}, kept for this hire only.");
        reply.Append(CultureInfo.InvariantCulture,
            $"\n\nInstance: tenant {record.Key.TenantId}, agent user {record.Key.AgentUserId}.");
    }

    /// <summary>
    /// Lists the lines found, each followed by the number of the source it came from. A line already quoted from one
    /// file isn't repeated for a copy of it, and a source with nothing new to add gets no number.
    /// </summary>
    private static Reply Answer(StringBuilder reply, string question, IReadOnlyList<Finding>? findings)
    {
        if (findings is null)
        {
            reply.Append("I couldn't search right now. Check that Work IQ is on in this tenant.");
            return new Reply(reply.ToString());
        }

        if (findings.Count == 0)
        {
            reply.Append("I didn't find that in what's shared with me or in this chat.");
            return new Reply(reply.ToString());
        }

        reply.Append("Here's what I found.\n");
        var words = TextMatch.Words(question);
        var quoted = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var citations = new List<Citation>();
        foreach (var finding in findings)
        {
            var lines = finding.Lines.Where(quoted.Add).ToList();
            if (lines.Count == 0)
            {
                continue;
            }

            var position = citations.Count + 1;
            foreach (var line in lines)
            {
                reply.Append(CultureInfo.InvariantCulture, $"\n- {line} [{position}]");
            }

            var summary = string.Join(' ', finding.Lines);
            citations.Add(new Citation(position, finding.Source, finding.Url, summary, [.. TextMatch.Words(summary).Where(words.Contains)]));
        }

        return new Reply(reply.ToString()) { Citations = citations, Answer = true };
    }

    private static string Count(int notes) => notes == 1 ? "1 note" : $"{notes} notes";
}
