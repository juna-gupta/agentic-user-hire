using System.Text;
using System.Text.Json;
using ModelContextProtocol.Protocol;

namespace NotesTeammate.Tests;

public class AnswerFinderTests
{
    private const string Self = "a7d1c0de-0000-4000-8000-000000000001";
    private const string Chat = "19:a7d1c0de_dana@unq.gbl.spaces";
    private const string Question = "When does Project Kestrel go public?";

    [Fact]
    public async Task Files_that_retrieve_finds_are_read_and_quoted()
    {
        var tools = new FakeWorkIq();
        tools.Answers["retrieve"] = Retrieval(hits: """
            [{"id":"h1","webUrl":"https://contoso.example/Kestrel.docx","entityUrl":"/drives/b!d1/items/01AB",
              "resourceMetadata":{"document":{"title":"Kestrel team notebook"}}}]
            """);
        tools.Answers["fetch_blob"] = Blob("Kestrel.docx", TextMatchTests.Word("Launch", "Project Kestrel goes public on Tuesday, November 18."));

        var found = await AnswerFinder.FindAsync(tools, Question, Chat, Self, null, null, default);

        var finding = Assert.Single(found);
        Assert.Equal(new[] { "Project Kestrel goes public on Tuesday, November 18." }, finding.Lines);
        Assert.Equal("https://contoso.example/Kestrel.docx", finding.Url);
        Assert.Equal("/drives/b!d1/items/01AB/content", tools.Calls.Last().Arguments["path"]);
        Assert.DoesNotContain(tools.Calls, call => call.Tool == "fetch");
    }

    [Fact]
    public async Task A_hit_that_isnt_a_file_is_quoted_from_its_snippet()
    {
        var tools = new FakeWorkIq();
        tools.Answers["retrieve"] = Retrieval(
            hits: """[{"id":"h1","webUrl":"https://outlook.example/1"}]""",
            markdown: "- **Kestrel launch date** [^h1]\n  - **snippet:** Hi all. Project Kestrel goes public on November 18.");

        var found = await AnswerFinder.FindAsync(tools, Question, Chat, Self, null, null, default);

        Assert.Equal("Kestrel launch date", Assert.Single(found).Source);
        Assert.Equal(new[] { "Project Kestrel goes public on November 18." }, found[0].Lines);
    }

    [Fact]
    public async Task When_retrieve_finds_nothing_the_chat_and_its_files_are_read()
    {
        var tools = new FakeWorkIq();
        tools.Answers["retrieve"] = Retrieval(hits: "[]");
        tools.Answers["fetch"] = Json($$$"""
            {"results":[{"statusCode":200,"data":{"value":[
              {"messageType":"message","from":{"user":{"id":"{{{Self}}}"}},"body":{"content":"Project Kestrel: I don't know yet"}},
              {"messageType":"message","from":{"user":{"id":"dana"}},"body":{"content":"<p>When does Project Kestrel go public?</p>"}},
              {"messageType":"message","from":{"user":{"id":"dana"}},"body":{"content":"<p>remember Kestrel goes public on the 18th</p>"},
               "attachments":[{"contentType":"reference","name":"plan.md","contentUrl":"https://contoso.example/plan.md"}]}
            ]}}]}
            """);
        tools.Answers["fetch_blob"] = Blob("plan.md", Encoding.UTF8.GetBytes("# Plan\n\n- Project Kestrel goes public in November."));

        var found = await AnswerFinder.FindAsync(tools, Question, Chat, Self, null, null, default);

        Assert.Equal(new[] { "This chat", "plan.md" }, found.Select(finding => finding.Source));
        Assert.Equal(new[] { "Kestrel goes public on the 18th" }, found[0].Lines);
        Assert.Equal(new[] { "Project Kestrel goes public in November." }, found[1].Lines);
        Assert.Equal($"/chats/{Chat}/messages?$top=50", ((string[])tools.Calls.Single(call => call.Tool == "fetch").Arguments["entityUrls"]!)[0]);
        Assert.Equal("/shares/" + AnswerFinder.ShareId("https://contoso.example/plan.md") + "/driveItem/content",
            tools.Calls.Single(call => call.Tool == "fetch_blob").Arguments["path"]);
    }

    [Fact]
    public async Task Each_step_is_reported_so_Teams_can_show_it()
    {
        var tools = new FakeWorkIq();
        tools.Answers["retrieve"] = Retrieval(hits: "[]");
        tools.Answers["fetch"] = Json($$$"""
            {"results":[{"statusCode":200,"data":{"value":[
              {"messageType":"message","from":{"user":{"id":"dana"}},"body":{"content":"<p>See the plan</p>"},
               "attachments":[{"contentType":"reference","name":"plan.md","contentUrl":"https://contoso.example/plan.md"}]}
            ]}}]}
            """);
        tools.Answers["fetch_blob"] = Blob("plan.md", Encoding.UTF8.GetBytes("- Project Kestrel goes public in November."));
        var steps = new List<string>();

        await AnswerFinder.FindAsync(tools, Question, Chat, Self, null, (text, _) => { steps.Add(text); return Task.CompletedTask; }, default);

        Assert.Equal(new[] { "Searching what's shared with me", "Nothing shared with me matched, so reading this chat", "Reading plan.md from this chat" }, steps);
    }

    [Fact]
    public async Task A_failed_retrieve_is_an_error_not_an_empty_answer()
    {
        var tools = new FakeWorkIq();
        tools.Answers["retrieve"] = new CallToolResult { Content = [new TextContentBlock { Text = "No billing policy" }], IsError = true };

        await Assert.ThrowsAsync<InvalidOperationException>(() => AnswerFinder.FindAsync(tools, Question, Chat, Self, null, null, default));
    }

    [Fact]
    public void A_sharing_id_is_unpadded_base64url()
    {
        Assert.Equal("u!aHR0cHM6Ly9vbmVkcml2ZS5leGFtcGxlL2E_Yj1j", AnswerFinder.ShareId("https://onedrive.example/a?b=c"));
    }

    private static CallToolResult Retrieval(string hits, string markdown = "") =>
        Json($$$"""{"application/vnd.ms-workiq.retrieval":{"markdown":{{{JsonSerializer.Serialize(markdown)}}},"retrievalHits":{{{hits}}}}}""");

    private static CallToolResult Blob(string name, byte[] content) =>
        Json(JsonSerializer.Serialize(new { statusCode = 200, blobName = name, base64Content = Convert.ToBase64String(content) }));

    private static CallToolResult Json(string json) => new() { StructuredContent = JsonDocument.Parse(json).RootElement.Clone() };

    private sealed class FakeWorkIq : IWorkIqTools
    {
        public Dictionary<string, CallToolResult> Answers { get; } = [];

        public List<(string Tool, IReadOnlyDictionary<string, object?> Arguments)> Calls { get; } = [];

        public Task<CallToolResult> CallAsync(string tool, IReadOnlyDictionary<string, object?> arguments, CancellationToken cancellationToken)
        {
            Calls.Add((tool, arguments));
            return Task.FromResult(Answers[tool]);
        }
    }
}
