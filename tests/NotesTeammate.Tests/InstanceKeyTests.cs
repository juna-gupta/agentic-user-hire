using Microsoft.Teams.Core.Schema;

namespace NotesTeammate.Tests;

public class InstanceKeyTests
{
    private const string Tenant = "11111111-1111-1111-1111-111111111111";
    private const string AgentUser = "22222222-2222-2222-2222-222222222222";

    [Fact]
    public void KeyComesFromTheRecipient()
    {
        var recipient = new ChannelAccount { Id = "29:agent", AgenticUserId = AgentUser, AgenticAppId = Guid.NewGuid().ToString() };

        var key = InstanceKey.ForRecipient(Tenant, recipient);

        Assert.Equal(new InstanceKey(Guid.Parse(Tenant), Guid.Parse(AgentUser)), key);
    }

    [Fact]
    public void RecipientThatIsNotAnAgentUserHasNoKey()
    {
        Assert.Null(InstanceKey.ForRecipient(Tenant, new ChannelAccount { Id = "28:classic-bot" }));
        Assert.Null(InstanceKey.ForRecipient(Tenant, null));
    }

    [Theory]
    [InlineData(null, AgentUser)]
    [InlineData(Tenant, null)]
    [InlineData("../other-tenant", AgentUser)]
    [InlineData(Tenant, "..\\..\\secrets")]
    [InlineData("00000000-0000-0000-0000-000000000000", AgentUser)]
    public void OnlyGuidsMakeAKey(string? tenantId, string? agentUserId)
    {
        Assert.Null(InstanceKey.Create(tenantId, agentUserId));
    }
}
