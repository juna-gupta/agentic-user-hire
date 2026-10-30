using Microsoft.Teams.Apps.Schema.Entities;

namespace NotesTeammate.Tests;

public class TeamsReplyTests
{
    private static readonly Reply Answer = new("Here's what I found.\n\n- Launch is on November 18. [1]")
    {
        Answer = true,
        Citations = [new Citation(1, "Kestrel notebook.docx", "https://contoso.example/Kestrel notebook.docx", "Launch is on November 18.", ["launch"])],
    };

    [Fact]
    public void An_answer_is_markdown_with_the_AI_label_feedback_and_its_citations()
    {
        var message = TeamsReply.Render(Answer, personal: true, askerId: "dana");

        Assert.Equal(Answer.Text, message.Text);
        Assert.Equal("markdown", message.TextFormat?.ToString(), ignoreCase: true);
        var root = Assert.Single(message.Entities!.OfType<OMessageEntity>(), e => e.OType == "Message");
        Assert.Equal("Message", root.Properties["@type"]?.ToString());
        Assert.Equal("https://schema.org", root.Properties["@context"]?.ToString());
    }

    [Fact]
    public void A_plain_reply_has_no_entities_and_suggestions_only_in_a_1_1_chat()
    {
        var help = new Reply("Try notes.") { Suggestions = ["notes", "who are you"] };

        var personal = TeamsReply.Render(help, personal: true, askerId: "dana");
        var group = TeamsReply.Render(help, personal: false, askerId: "dana");

        Assert.True(personal.Entities is null || personal.Entities.Count == 0);
        Assert.Equal(2, personal.SuggestedActions?.Actions.Count);
        Assert.Null(group.SuggestedActions);
    }

    [Theory]
    [InlineData("Kestrel notebook.docx", null, "Microsoft Word")]
    [InlineData("Kestrel notes", "https://contoso.example/Kestrel%20notes.md", "Text")]
    [InlineData("Budget.xlsx", null, "Microsoft Excel")]
    [InlineData("This chat", null, null)]
    public void The_citation_icon_follows_the_file_type(string title, string? url, string? icon)
    {
        Assert.Equal(icon, TeamsReply.IconFor(new Citation(1, title, url, "", []))?.ToString());
    }
}
