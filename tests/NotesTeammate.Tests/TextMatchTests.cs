using System.IO.Compression;
using System.Text;

namespace NotesTeammate.Tests;

public class TextMatchTests
{
    private const string Notebook = """
        # Kestrel team notebook

        ## Launch

        - Project Kestrel goes public on Tuesday, November 18, 2026.
        - The launch review is on Friday, November 14, at 2 PM Pacific.

        ## People

        - Release owner: Priya Raman.
        - Docs owner: Marco Diaz.
        - Support lead: Lena Okafor.

        ## Open questions

        - Who signs off the press release, Priya or the comms team?
        """;

    [Fact]
    public void Common_words_and_plurals_are_dropped()
    {
        Assert.Equal(new[] { "project", "kestrel", "public", "release", "owner" },
            TextMatch.Words("When does Project Kestrel go public and who are the release owners?"));
    }

    [Fact]
    public void The_lines_that_match_the_question_are_picked()
    {
        var picked = Ask("When does Project Kestrel go public and who is the release owner?", TextMatch.Lines(Notebook));

        Assert.Equal(new[] { "Project Kestrel goes public on Tuesday, November 18, 2026.", "Release owner: Priya Raman." }, picked);
    }

    [Fact]
    public void Words_from_the_title_alone_match_nothing()
    {
        Assert.Empty(Ask("What is the marketing budget for Kestrel?", TextMatch.Lines(Notebook)));
    }

    [Fact]
    public void A_search_snippet_is_split_like_the_file()
    {
        var snippet = "# Kestrel team notebook ## Launch - Project Kestrel goes public on Tuesday, November 18, 2026. "
            + "- The launch review is on Friday. ## People - Release owner: Priya Raman. - Support lead: Lena Okafor.";

        Assert.Equal(new[] { "Support lead: Lena Okafor." }, Ask("Who is the support lead?", TextMatch.Lines(snippet)));
    }

    [Fact]
    public void A_chat_snippet_is_split_by_message_without_the_hires_own()
    {
        var snippet = "From:Dana Lee (30 October 2026 05:37 AM) remember the demo moved to Monday "
            + "From:NinaNotes (30 October 2026 05:37 AM) Got it. The demo is noted.";

        Assert.Equal(new[] { "the demo moved to Monday" }, TextMatch.Snippet(snippet, "NinaNotes"));
    }

    [Fact]
    public void A_word_document_is_read_paragraph_by_paragraph()
    {
        var text = TextMatch.ReadDocument(Word("Launch", "Project Kestrel goes public on Tuesday."), "notebook.docx", null);

        Assert.Equal("Launch\nProject Kestrel goes public on Tuesday.", text);
    }

    [Fact]
    public void Other_kinds_of_file_are_skipped()
    {
        Assert.Null(TextMatch.ReadDocument([1, 2, 3], "slides.pptx", "application/vnd.ms-powerpoint"));
        Assert.Null(TextMatch.ReadDocument([1, 2, 3], "broken.docx", null));
    }

    [Fact]
    public void A_teams_message_is_read_without_its_html()
    {
        Assert.Equal("remember the demo moved to Monday\nand Tom & Ana present\n",
            TextMatch.FromHtml("<p>remember the demo moved to <b>Monday</b></p><p>and Tom &amp; Ana present</p>"));
    }

    private static List<string> Ask(string question, IEnumerable<string> lines) =>
        TextMatch.Pick(lines, TextMatch.Words(question), "Kestrel team notebook.docx");

    internal static byte[] Word(params string[] paragraphs)
    {
        var body = string.Concat(paragraphs.Select(p => $"<w:p><w:r><w:t>{p}</w:t></w:r></w:p>"));
        using var stream = new MemoryStream();
        using (var zip = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            using var writer = new StreamWriter(zip.CreateEntry("word/document.xml").Open(), Encoding.UTF8);
            writer.Write($"""<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{body}</w:body></w:document>""");
        }

        return stream.ToArray();
    }
}
