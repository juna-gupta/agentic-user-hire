using System.Text.RegularExpressions;
using Microsoft.Teams.Apps;
using Microsoft.Teams.Apps.Schema;
using Microsoft.Teams.Apps.Schema.Entities;
using Microsoft.Teams.Core.Schema;

namespace NotesTeammate;

/// <summary>
/// Turns a <see cref="Reply"/> into a Teams message. Every reply is markdown. An answer also gets the AI label,
/// feedback buttons and a citation for each file or chat it quotes, and in a 1:1 chat a reply can offer suggested
/// messages to send next.
/// </summary>
internal static partial class TeamsReply
{
    public static MessageActivityInput Render(Reply reply, bool personal, string? askerId)
    {
        var input = new MessageActivityInput().WithText(reply.Text, TextFormats.Markdown);
        if (reply.Answer)
        {
            input.AddAIGenerated();
            input.AddFeedback(true);
            foreach (var citation in reply.Citations)
            {
                input.AddCitation(citation.Position, new CitationAppearance
                {
                    Name = Clip(citation.Title, 80),
                    Abstract = Clip(citation.Abstract, 160),
                    Url = citation.Url is { } url ? new Uri(url) : null,
                    Icon = IconFor(citation),
                    Keywords = [.. citation.Keywords.Take(3).Select(k => Clip(k, 28))],
                });
            }

            // Teams reads the label, feedback buttons and citations from the message entity, a schema.org Message.
            // Set its type and context explicitly so they're always sent.
            var root = input.Entities!.OfType<OMessageEntity>().First(e => e.OType == "Message");
            root.Properties["@type"] = "Message";
            root.Properties["@context"] = "https://schema.org";
            root.Properties["@id"] = "";
        }

        // Suggested messages are addressed to the person who asked, in 1:1 chats only.
        if (personal && reply.Suggestions.Count > 0 && askerId is not null)
        {
            var suggestions = new SuggestedActions { Actions = [.. reply.Suggestions.Select(s => new SuggestedAction(ActionTypes.IMBack, s, s))] };
            input.WithSuggestedActions(suggestions.AddRecipients(askerId));
        }

        return input;
    }

    /// <summary>The icon Teams shows on a citation, from the file type in its title or link. A chat gets none.</summary>
    internal static CitationIcon? IconFor(Citation citation)
    {
        var match = Extension().Match(citation.Title);
        if (!match.Success && citation.Url is { } url)
        {
            match = Extension().Match(Uri.UnescapeDataString(url));
        }

        return match.Success
            ? match.Groups[1].Value.ToLowerInvariant() switch
            {
                "doc" or "docx" => CitationIcon.MicrosoftWord,
                "xls" or "xlsx" => CitationIcon.MicrosoftExcel,
                "ppt" or "pptx" => CitationIcon.MicrosoftPowerPoint,
                "pdf" => CitationIcon.Pdf,
                _ => CitationIcon.Text,
            }
            : null;
    }

    private static string Clip(string text, int max) => text.Length <= max ? text : string.Concat(text.AsSpan(0, max - 3), "...");

    [GeneratedRegex(@"\.(docx?|xlsx?|pptx?|pdf|md|txt)\b", RegexOptions.IgnoreCase)]
    private static partial Regex Extension();
}
