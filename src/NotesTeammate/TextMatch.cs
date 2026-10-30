using System.IO.Compression;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace NotesTeammate;

/// <summary>
/// Picks the lines of a file or chat that answer a question, by the words they share with it. There's no language
/// model, so it can't rephrase or combine facts; it quotes the lines and says where they came from.
/// </summary>
internal static partial class TextMatch
{
    private const int MaxLine = 300;

    private static readonly HashSet<string> Common = new(StringComparer.Ordinal)
    {
        "about", "all", "and", "any", "are", "can", "could", "did", "does", "for", "from", "get", "got", "had", "has",
        "have", "how", "into", "its", "know", "many", "much", "not", "our", "please", "should", "tell", "than", "that",
        "the", "their", "them", "then", "there", "they", "this", "was", "were", "what", "whats", "when", "where",
        "which", "who", "whom", "why", "will", "with", "would", "you", "your",
    };

    private static readonly XNamespace Word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /// <summary>The words that carry meaning: lowercase, three letters or more, common words and plural s removed.</summary>
    public static HashSet<string> Words(string? text)
    {
        var words = new HashSet<string>(StringComparer.Ordinal);
        foreach (Match match in WordPattern().Matches(text ?? string.Empty))
        {
            var word = match.Value.ToLowerInvariant().Replace("'", string.Empty, StringComparison.Ordinal);
            if (word.Length >= 3 && !Common.Contains(word))
            {
                words.Add(word.Length > 4 && word.EndsWith('s') && !word.EndsWith("ss", StringComparison.Ordinal) ? word[..^1] : word);
            }
        }

        return words;
    }

    /// <summary>Splits text into lines and sentences, without Markdown markers.</summary>
    public static List<string> Lines(string? text)
    {
        var lines = new List<string>();
        foreach (var part in Separator().Split(text ?? string.Empty))
        {
            var line = Marker().Replace(part, string.Empty).Trim();
            if (line.Length >= 3)
            {
                lines.Add(line.Length <= MaxLine ? line : line[..(MaxLine - 3)] + "...");
            }
        }

        return lines;
    }

    /// <summary>
    /// The lines of a search snippet. A chat's snippet is a transcript, "From:Ava (30 October 2026 05:37 AM) text
    /// From:Nina (...) text", so it's split by message. The hire's own messages are left out, and so is "remember" at
    /// the start of a message.
    /// </summary>
    public static List<string> Snippet(string? text, string? self)
    {
        text ??= string.Empty;
        var speakers = Speaker().Matches(text);
        if (speakers.Count == 0)
        {
            return Lines(text);
        }

        var lines = Lines(text[..speakers[0].Index]);
        for (var i = 0; i < speakers.Count; i++)
        {
            var start = speakers[i].Index + speakers[i].Length;
            var end = i + 1 < speakers.Count ? speakers[i + 1].Index : text.Length;
            if (!string.Equals(speakers[i].Groups["name"].Value.Trim(), self, StringComparison.OrdinalIgnoreCase))
            {
                lines.AddRange(Lines(WithoutRemember(text[start..end])));
            }
        }

        return lines;
    }

    /// <summary>A note as it was said, without the "remember" that asked for it.</summary>
    public static string WithoutRemember(string text)
    {
        var trimmed = text.TrimStart();
        return trimmed.StartsWith(Teammate.Remember, StringComparison.OrdinalIgnoreCase) ? trimmed[Teammate.Remember.Length..] : trimmed;
    }

    /// <summary>
    /// The lines that share the most words with the question, in their original order. Each line has to bring a word
    /// the lines picked before it didn't, so "Docs owner" isn't quoted next to "Release owner" just for "owner".
    /// Words in the source's title are left out when the question has others, since every line of "Kestrel notes" is
    /// about Kestrel. Questions in the source don't answer anything, so they're skipped.
    /// </summary>
    public static List<string> Pick(IEnumerable<string> lines, IReadOnlySet<string> question, string? title, int max = 3)
    {
        var keys = new HashSet<string>(question);
        keys.ExceptWith(Words(title));
        if (keys.Count == 0)
        {
            keys.UnionWith(question);
        }

        var candidates = lines
            .Where(line => !line.EndsWith('?'))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Select((line, index) => (line, index, matched: Words(line).Where(keys.Contains).ToList()))
            .Where(candidate => candidate.matched.Count > 0)
            .OrderByDescending(candidate => candidate.matched.Count)
            .ThenBy(candidate => candidate.index);

        var covered = new HashSet<string>(StringComparer.Ordinal);
        var picked = new List<(string Line, int Index)>();
        foreach (var (line, index, matched) in candidates)
        {
            if (picked.Count < max && matched.Any(word => !covered.Contains(word)))
            {
                covered.UnionWith(matched);
                picked.Add((line, index));
            }
        }

        return [.. picked.OrderBy(p => p.Index).Select(p => p.Line)];
    }

    /// <summary>The text of a Word document or a text file, or null for other kinds of file.</summary>
    public static string? ReadDocument(byte[] content, string? name, string? contentType)
    {
        var extension = Path.GetExtension(name ?? string.Empty).ToLowerInvariant();
        if (extension == ".docx" || contentType?.Contains("wordprocessingml", StringComparison.OrdinalIgnoreCase) == true)
        {
            return ReadWord(content);
        }

        return extension is ".md" or ".txt" or ".csv" or ".log" || contentType?.StartsWith("text/", StringComparison.OrdinalIgnoreCase) == true
            ? Encoding.UTF8.GetString(content)
            : null;
    }

    /// <summary>The text of a Teams message body, which is HTML.</summary>
    public static string FromHtml(string? html)
    {
        var text = Break().Replace(html ?? string.Empty, "\n");
        return WebUtility.HtmlDecode(Tag().Replace(text, string.Empty));
    }

    // A .docx file is a zip; the body is word/document.xml, with one w:p element per paragraph.
    private static string? ReadWord(byte[] content)
    {
        try
        {
            using var zip = new ZipArchive(new MemoryStream(content), ZipArchiveMode.Read);
            if (zip.GetEntry("word/document.xml") is not { } body)
            {
                return null;
            }

            using var stream = body.Open();
            var document = XDocument.Load(stream);
            return string.Join('\n', document.Descendants(Word + "p").Select(paragraph => string.Concat(paragraph.Descendants(Word + "t").Select(t => t.Value))));
        }
        catch (Exception ex) when (ex is InvalidDataException or System.Xml.XmlException)
        {
            return null;
        }
    }

    [GeneratedRegex(@"[\p{L}\p{N}']+")]
    private static partial Regex WordPattern();

    // New lines, sentence ends, and the headings and list items that search snippets flatten onto one line.
    [GeneratedRegex(@"\r?\n|(?<=[.!])\s+|\s+#{1,6}\s+|\s+[-*]\s+")]
    private static partial Regex Separator();

    [GeneratedRegex(@"^\s*(#{1,6}|[-*]|\d+\.)\s+")]
    private static partial Regex Marker();

    [GeneratedRegex(@"<br\s*/?>|</p>|</div>|</li>", RegexOptions.IgnoreCase)]
    private static partial Regex Break();

    [GeneratedRegex(@"<[^>]+>")]
    private static partial Regex Tag();

    // A message header in a chat snippet: "From:Ava Smith (30 October 2026 05:37 AM)".
    [GeneratedRegex(@"From:\s*(?<name>[^()\r\n]+?)\s*\([^()]*\d{4}[^()]*\)\s*")]
    private static partial Regex Speaker();
}
