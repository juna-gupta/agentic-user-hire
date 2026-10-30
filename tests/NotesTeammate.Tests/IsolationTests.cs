namespace NotesTeammate.Tests;

/// <summary>Two people hire the agent: Ava's instance and Ben's instance must never see each other's data.</summary>
public sealed class IsolationTests : IDisposable
{
    private static readonly Guid Contoso = Guid.NewGuid();
    private static readonly Guid Fabrikam = Guid.NewGuid();

    private readonly string _folder = Path.Combine(Path.GetTempPath(), "notes-teammate-tests", Guid.NewGuid().ToString("N"));
    private readonly FileInstanceStore _store;
    private readonly Teammate _teammate;

    private readonly InstanceKey _ava = new(Contoso, Guid.NewGuid());
    private readonly InstanceKey _ben = new(Contoso, Guid.NewGuid());

    public IsolationTests()
    {
        _store = new FileInstanceStore(_folder);
        _teammate = new Teammate(_store, TimeProvider.System);
    }

    public void Dispose()
    {
        if (Directory.Exists(_folder))
        {
            Directory.Delete(_folder, recursive: true);
        }
    }

    [Fact]
    public async Task TwoHiresInTheSameTenantKeepSeparateNotes()
    {
        await Say(_ava, "remember the launch moved to Friday");

        Assert.Contains("the launch moved to Friday", await Say(_ava, "notes"));
        Assert.DoesNotContain("launch", await Say(_ben, "notes"));
    }

    [Fact]
    public async Task TheSameAgentUserIdInAnotherTenantIsAnotherInstance()
    {
        var elsewhere = _ava with { TenantId = Fabrikam };
        await Say(_ava, "remember budget is approved");

        Assert.DoesNotContain("budget", await Say(elsewhere, "notes"));
    }

    [Fact]
    public async Task RemovingAHireDeletesOnlyItsData()
    {
        await Say(_ava, "remember Ava's note");
        await Say(_ben, "remember Ben's note");

        await _teammate.RemovedAsync(_ava);

        var removed = await _store.GetAsync(_ava);
        Assert.Empty(removed!.Notes);
        Assert.Null(removed.HiredBy);
        Assert.Single((await _store.GetAsync(_ben))!.Notes);
    }

    [Fact]
    public async Task ARemovedInstanceIgnoresLateMessages()
    {
        await _teammate.HiredAsync(_ava, hiredBy: "Dana");
        await _teammate.RemovedAsync(_ava);

        Assert.Null(await Say(_ava, "remember this"));
        Assert.Empty((await _store.GetAsync(_ava))!.Notes);
    }

    [Fact]
    public async Task ARestoredInstanceStartsOver()
    {
        await Say(_ava, "remember Ava's note");
        await _teammate.RemovedAsync(_ava);
        await _teammate.RestoredAsync(_ava);

        Assert.Contains("I don't have any notes yet", await Say(_ava, "notes"));
    }

    [Fact]
    public async Task TheFirstMessageGreetsAndNamesTheHirer()
    {
        await _teammate.HiredAsync(_ava, hiredBy: "Dana");

        var first = await Say(_ava, "hello");
        var second = await Say(_ava, "hello");

        Assert.Contains("hired by Dana", first);
        Assert.DoesNotContain("Hi ", second);
    }

    [Fact]
    public async Task HelpAndSavedNotesOfferWhatToTryNext()
    {
        await _teammate.HiredAsync(_ava, hiredBy: "Dana");
        await Say(_ava, "hello");

        var help = await _teammate.ReplyAsync(_ava, "hello", "Dana", _ => Task.FromResult<AgentProfile?>(null));
        var saved = await _teammate.ReplyAsync(_ava, "remember the demo is on Monday", "Dana", _ => Task.FromResult<AgentProfile?>(null));

        Assert.Equal(new[] { "notes", "who are you" }, help!.Suggestions);
        Assert.False(help.Noted);
        Assert.True(saved!.Noted);
        Assert.Equal(new[] { "notes" }, saved.Suggestions);
        Assert.False(saved.Answer);
    }

    [Fact]
    public async Task WhoAreYouAnswersWithTheInstancesOwnProfile()
    {
        await _teammate.HiredAsync(_ava, hiredBy: "Dana");

        var reply = await _teammate.ReplyAsync(_ava, "who are you?", "Dana",
            _ => Task.FromResult<AgentProfile?>(new AgentProfile("Ava (Notes)", "ava.notes@contoso.example")));

        Assert.Contains("I'm Ava (Notes) (ava.notes@contoso.example)", reply!.Text);
        Assert.Contains(_ava.AgentUserId.ToString(), reply.Text);
    }

    [Fact]
    public async Task AQuestionIsAnsweredFromWhatsSharedWithTheInstance()
    {
        IReadOnlyList<Finding> found =
        [
            new Finding("Kestrel notebook.docx", "https://contoso.example/Kestrel notebook.docx", ["Launch is on November 18.", "Release owner: Priya."]),
            new Finding("Kestrel notebook.md", "https://contoso.example/Kestrel notebook.md", ["Launch is on November 18."]),
            new Finding("This chat", null, ["Launch party is on November 19."]),
        ];
        string? asked = null;

        var reply = await _teammate.ReplyAsync(_ava, "When is the launch?", "Dana", _ => Task.FromResult<AgentProfile?>(null),
            (question, _) => { asked = question; return Task.FromResult<IReadOnlyList<Finding>?>(found); });

        Assert.Equal("When is the launch?", asked);
        Assert.True(reply!.Answer);
        Assert.Equal("Here's what I found.\n\n- Launch is on November 18. [1]\n- Release owner: Priya. [1]\n- Launch party is on November 19. [2]", reply.Text[reply.Text.IndexOf("Here's", StringComparison.Ordinal)..]);
        Assert.Collection(reply.Citations,
            c =>
            {
                Assert.Equal((1, "Kestrel notebook.docx", "https://contoso.example/Kestrel notebook.docx"), (c.Position, c.Title, c.Url));
                Assert.Equal(new[] { "launch" }, c.Keywords);
            },
            c => Assert.Equal((2, "This chat", (string?)null), (c.Position, c.Title, c.Url)));
    }

    [Fact]
    public async Task AQuestionWithNoAnswerSaysSo()
    {
        await _teammate.HiredAsync(_ava, hiredBy: "Dana");
        await Say(_ava, "hello");

        var reply = await _teammate.ReplyAsync(_ava, "What's the budget?", "Dana", _ => Task.FromResult<AgentProfile?>(null),
            (_, _) => Task.FromResult<IReadOnlyList<Finding>?>([]));

        Assert.StartsWith("I didn't find that", reply!.Text);
        Assert.False(reply.Answer);
    }

    [Fact]
    public async Task ADisabledInstanceDoesNothing()
    {
        await _teammate.SetEnabledAsync(_ava, enabled: false);

        Assert.Null(await Say(_ava, "remember this"));
        Assert.Empty((await _store.GetAsync(_ava))!.Notes);
    }

    private async Task<string?> Say(InstanceKey instance, string text) =>
        (await _teammate.ReplyAsync(instance, text, "Dana", _ => Task.FromResult<AgentProfile?>(null)))?.Text;
}
