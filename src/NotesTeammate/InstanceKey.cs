using Microsoft.Teams.Core.Schema;

namespace NotesTeammate;

/// <summary>
/// One hired copy of the agent: the customer's tenant plus the agent user that the hire created. Every read and write
/// of instance data goes through this key, so one hire never sees another's data.
/// </summary>
public readonly record struct InstanceKey(Guid TenantId, Guid AgentUserId)
{
    /// <summary>Null unless both IDs are GUIDs. They name files and partitions, so nothing else gets through.</summary>
    public static InstanceKey? Create(string? tenantId, string? agentUserId) =>
        Guid.TryParse(tenantId, out var tenant) && tenant != Guid.Empty
        && Guid.TryParse(agentUserId, out var agentUser) && agentUser != Guid.Empty
            ? new InstanceKey(tenant, agentUser)
            : null;

    /// <summary>
    /// The instance a Teams activity is for: the agent user it was sent to (the recipient), never the sender or
    /// anything in the message text. Null when the recipient isn't an agent user.
    /// </summary>
    public static InstanceKey? ForRecipient(string? tenantId, ChannelAccount? recipient) =>
        Create(tenantId, recipient?.AgenticUserId);

    public override string ToString() => $"{TenantId:D}/{AgentUserId:D}";
}
