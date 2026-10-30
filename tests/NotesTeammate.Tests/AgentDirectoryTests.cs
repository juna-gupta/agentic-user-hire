using System.Text.Json;
using ModelContextProtocol.Protocol;

namespace NotesTeammate.Tests;

public class AgentDirectoryTests
{
    private const string Profile = """
        {"results":[{"data":{"displayName":"Ava (Notes Teammate)","userPrincipalName":"notes-ava@contoso.example"},"statusCode":200}]}
        """;

    [Fact]
    public void The_profile_is_read_from_the_fetch_text()
    {
        var profile = AgentDirectory.ReadProfile(Text(Profile));

        Assert.Equal(new AgentProfile("Ava (Notes Teammate)", "notes-ava@contoso.example"), profile);
    }

    [Fact]
    public void Structured_content_is_used_when_present()
    {
        var result = new CallToolResult { StructuredContent = JsonDocument.Parse(Profile).RootElement.Clone() };

        Assert.Equal("Ava (Notes Teammate)", AgentDirectory.ReadProfile(result)?.DisplayName);
    }

    [Theory]
    [InlineData("""{"results":[{"statusCode":403,"error":{"message":"Forbidden"}}]}""")]
    [InlineData("""{"results":[]}""")]
    [InlineData("""{"error":"Work IQ is turned off for this tenant."}""")]
    [InlineData("Work IQ is turned off for this tenant.")]
    public void Anything_but_a_profile_reads_as_none(string text)
    {
        Assert.Null(AgentDirectory.ReadProfile(Text(text, isError: true)));
    }

    private static CallToolResult Text(string text, bool isError = false) =>
        new() { Content = [new TextContentBlock { Text = text }], IsError = isError };
}
