using Microsoft.Teams.Apps;
using Microsoft.Teams.Apps.Clients;
using Microsoft.Teams.Apps.Schema;
using Microsoft.Teams.Core.Schema;

namespace NotesTeammate;

/// <summary>
/// Connects Teams to the teammate. Lifecycle events say when an instance is hired, paused, removed or restored, and
/// each message is answered by the instance it was sent to. Teams only sends this service activities for its own
/// blueprint; the SDK checks every request's token before these handlers run.
/// </summary>
public sealed class TeamsRoutes(Teammate teammate, AgentDirectory directory, AnswerFinder finder, ILogger<TeamsRoutes> logger)
{
    public void Map(TeamsBotApplication teams)
    {
        teams.OnAgenticUserIdentityCreated(OnHiredAsync);
        teams.OnAgenticUserEnabled((context, cancellationToken) => SetEnabledAsync(context.Activity.Value, true, cancellationToken));
        teams.OnAgenticUserDisabled((context, cancellationToken) => SetEnabledAsync(context.Activity.Value, false, cancellationToken));
        teams.OnAgenticUserDeleted(OnRemovedAsync);
        teams.OnAgenticUserUndeleted(OnRestoredAsync);

        // The other lifecycle events need nothing from this teammate, so they are only logged. Handling them keeps the
        // log a complete record of what Teams sends for each hire.
        teams.OnAgenticUserWorkloadOnboardingUpdated((context, cancellationToken) => LogOnlyAsync(context.Activity.Value,
            $"{context.Activity.Value?.WorkloadName} setup {context.Activity.Value?.WorkloadOnboardingState}"));
        teams.OnAgenticUserIdentityUpdated((context, cancellationToken) => LogOnlyAsync(context.Activity.Value,
            $"Profile updated ({context.Activity.Value?.UpdatedProperty?.PropertyName})"));
        teams.OnAgenticUserManagerUpdated((context, cancellationToken) => LogOnlyAsync(context.Activity.Value, "Manager changed"));
        teams.OnMessage(OnMessageAsync);
        teams.OnMessageSubmitFeedback(OnFeedbackAsync);
    }

    /// <summary>Reaction IDs in the form Teams uses for emoji: the Unicode code point, then the emoji's name.</summary>
    private static readonly ReactionType Searching = new("1f440_eyes");
    private static readonly ReactionType Saved = new("2705_whiteheavycheckmark");

    private Task LogOnlyAsync(AgentLifecycleValueBase? value, string change)
    {
        if (Key(value) is { } key)
        {
            logger.LogInformation("{Change}: instance {Instance}", change, key);
        }

        return Task.CompletedTask;
    }

    private Task OnHiredAsync(Context<AgentLifecycleEventActivity<AgenticUserIdentityCreatedValue>> context, CancellationToken cancellationToken)
    {
        var hire = context.Activity.Value;
        if (Key(hire) is not { } key)
        {
            return Task.CompletedTask;
        }

        logger.LogInformation("Hired: instance {Instance}", key);
        return teammate.HiredAsync(key, hire?.Manager?.DisplayName ?? hire?.Manager?.Email, cancellationToken);
    }

    private Task SetEnabledAsync(AgentLifecycleValueBase? change, bool enabled, CancellationToken cancellationToken)
    {
        if (Key(change) is not { } key)
        {
            return Task.CompletedTask;
        }

        logger.LogInformation("{Change}: instance {Instance}", enabled ? "Enabled" : "Disabled", key);
        return teammate.SetEnabledAsync(key, enabled, cancellationToken);
    }

    private Task OnRemovedAsync(Context<AgentLifecycleEventActivity<AgenticUserDeletedValue>> context, CancellationToken cancellationToken)
    {
        if (Key(context.Activity.Value) is not { } key)
        {
            return Task.CompletedTask;
        }

        logger.LogInformation("Removed: instance {Instance}, its notes deleted", key);
        return teammate.RemovedAsync(key, cancellationToken);
    }

    private Task OnRestoredAsync(Context<AgentLifecycleEventActivity<AgenticUserUndeletedValue>> context, CancellationToken cancellationToken)
    {
        if (Key(context.Activity.Value) is not { } key)
        {
            return Task.CompletedTask;
        }

        logger.LogInformation("Restored: instance {Instance}, starting over", key);
        return teammate.RestoredAsync(key, cancellationToken);
    }

    private async Task OnMessageAsync(Context<MessageActivity> context, CancellationToken cancellationToken)
    {
        var activity = context.Activity;
        if (!string.Equals(activity.ChannelId, "msteams", StringComparison.OrdinalIgnoreCase)
            || InstanceKey.ForRecipient(activity.TenantId, activity.Recipient) is not { } key
            || AgenticIdentity.FromAccount(activity.Recipient) is not { } identity)
        {
            logger.LogInformation("Ignored a message that isn't a Teams message to an agent user");
            return;
        }

        // The hire may be in a customer's tenant. Its tokens, for Work IQ and for the reply, are requested there.
        using var tenant = HireTenant.Use(key.TenantId);
        var text = activity.TextWithoutMentions ?? activity.Text ?? string.Empty;
        var conversationId = activity.Conversation?.Id;
        var chatId = conversationId is { } id && id.StartsWith("19:", StringComparison.Ordinal) ? id : null;
        var personal = string.Equals(activity.Conversation?.ConversationType?.Value ?? "personal", "personal", StringComparison.OrdinalIgnoreCase);

        // context.Api is bound to the recipient, so status steps, reactions and the reply all come from this instance's
        // agent user. In a 1:1 chat a search shows its steps; in a group chat, where steps would be noisy, the
        // question gets an eyes reaction until the answer is posted.
        var conversations = context.Api.Conversations;
        var stream = personal && conversationId is not null && activity.Id is not null
            ? new StatusStream(conversations, conversationId, activity.Id, logger)
            : null;
        var searching = false;
        var reply = await teammate.ReplyAsync(key, text, activity.From?.Name,
            token => directory.GetProfileAsync(identity, token),
            async (question, token) =>
            {
                if (stream is null)
                {
                    searching = await ReactAsync(conversations, conversationId, activity.Id, Searching, true, token);
                    await context.TypingAsync(token);
                }

                return await finder.FindAsync(identity, question, chatId, activity.Recipient?.Name, stream is null ? null : stream.StepAsync, token);
            },
            cancellationToken);
        if (reply is null)
        {
            logger.LogInformation("Ignored a message to paused or removed instance {Instance}", key);
            return;
        }

        var message = TeamsReply.Render(reply, personal, activity.From?.Id);
        if (stream is null || !await stream.FinishAsync(message, cancellationToken))
        {
            await context.SendAsync(message, cancellationToken);
        }

        if (searching)
        {
            await ReactAsync(conversations, conversationId, activity.Id, Searching, false, cancellationToken);
        }

        if (reply.Noted)
        {
            await ReactAsync(conversations, conversationId, activity.Id, Saved, true, cancellationToken);
        }
    }

    /// <summary>The thumbs up or down under an answer. Only the rating is logged, not the comment.</summary>
    private Task<InvokeResponse> OnFeedbackAsync(Context<InvokeActivity<MessageSubmitFeedbackValue>> context, CancellationToken cancellationToken)
    {
        logger.LogInformation("Feedback on answer {MessageId}: {Reaction}", context.Activity.ReplyToId, context.Activity.Value?.Reaction ?? "none");
        return Task.FromResult(InvokeResponse.Ok());
    }

    /// <summary>Adds or removes this agent user's reaction on a message. Best effort, so a refusal is only logged.</summary>
    private async Task<bool> ReactAsync(ConversationApiClient conversations, string? conversationId, string? messageId, ReactionType reaction, bool add, CancellationToken cancellationToken)
    {
        if (conversationId is null || messageId is null)
        {
            return false;
        }

        try
        {
            if (add)
            {
                await conversations.AddReactionAsync(conversationId, messageId, reaction.Value, cancellationToken: cancellationToken);
            }
            else
            {
                await conversations.DeleteReactionAsync(conversationId, messageId, reaction.Value, cancellationToken: cancellationToken);
            }

            return true;
        }
        catch (HttpRequestException ex)
        {
            logger.LogWarning("Couldn't {Action} the {Reaction} reaction ({Status})", add ? "add" : "remove", reaction.Value, ex.StatusCode);
            return false;
        }
    }

    private InstanceKey? Key(AgentLifecycleValueBase? value)
    {
        var key = InstanceKey.Create(value?.TenantId, value?.AgenticUserId);
        if (key is null)
        {
            logger.LogWarning("Ignored a lifecycle event without a tenant and agent user");
        }

        return key;
    }
}
