using System.Net;
using Microsoft.Teams.Apps;
using Microsoft.Teams.Apps.Clients;
using Microsoft.Teams.Apps.Schema;
using Microsoft.Teams.Apps.Schema.Entities;

namespace NotesTeammate;

/// <summary>
/// Shows what a hire is doing while it searches, such as "Reading Kestrel notes", with the Teams streaming protocol.
/// Each step replaces the one before, with a Stop button, and the reply replaces the last step. Updates go through
/// the turn's API client, which is bound to the agent user, so every step comes from the hire itself.
/// </summary>
/// <remarks>
/// Best effort: if Teams refuses an update, the steps stop and the reply is posted as a normal message.
/// See https://learn.microsoft.com/microsoftteams/platform/bots/streaming-ux.
/// </remarks>
internal sealed class StatusStream(ConversationApiClient conversations, string conversationId, string requestId, ILogger logger)
{
    /// <summary>Teams throttles streaming to about one update a second.</summary>
    private static readonly TimeSpan MinInterval = TimeSpan.FromSeconds(1);

    private string? _streamId;
    private int _sequence;
    private DateTimeOffset _lastSent = DateTimeOffset.MinValue;
    private bool _broken;
    private bool _stopped;

    public async Task StepAsync(string text, CancellationToken cancellationToken)
    {
        if (_broken || _stopped)
        {
            return;
        }

        var wait = _lastSent + MinInterval - DateTimeOffset.UtcNow;
        if (wait > TimeSpan.Zero)
        {
            await Task.Delay(wait, cancellationToken);
        }

        _sequence++;
        var update = new MessageActivityInput
        {
            Type = "typing",
            Text = text,
            ReplyToId = requestId,
            ChannelData = new TeamsOutboundChannelData { StreamId = _streamId, StreamType = StreamTypes.Informative, StreamSequence = _sequence },
        };
        update.AddEntity(new StreamInfoEntity { StreamId = _streamId, StreamType = StreamTypes.Informative, StreamSequence = _sequence });
        try
        {
            var sent = await conversations.CreateActivityAsync(conversationId, update, cancellationToken: cancellationToken);
            _streamId ??= sent?.Id;
            _lastSent = DateTimeOffset.UtcNow;
        }
        catch (HttpRequestException ex) when (Stopped(ex))
        {
            _stopped = true;
            logger.LogInformation("The person stopped the reply");
        }
        catch (HttpRequestException ex)
        {
            _broken = true;
            logger.LogWarning("Teams refused a status update ({Status}), so the reply will be posted normally", ex.StatusCode);
        }
    }

    /// <summary>
    /// Ends the stream with the reply. Returns false when there's no stream to end or Teams refused the reply, so the
    /// caller posts it normally, and true when it was posted or the person pressed Stop.
    /// </summary>
    public async Task<bool> FinishAsync(MessageActivityInput reply, CancellationToken cancellationToken)
    {
        if (_stopped)
        {
            return true;
        }

        if (_broken || _streamId is null)
        {
            return false;
        }

        var final = new StreamInfoEntity { StreamId = _streamId, StreamType = StreamTypes.Final };
        reply.ReplyToId = requestId;
        reply.AddEntity(final);
        try
        {
            await conversations.CreateActivityAsync(conversationId, reply, cancellationToken: cancellationToken);
            return true;
        }
        catch (HttpRequestException ex) when (Stopped(ex))
        {
            logger.LogInformation("The person stopped the reply");
            return true;
        }
        catch (HttpRequestException ex) when (ex.StatusCode == HttpStatusCode.BadRequest)
        {
            // Close the stream with the text alone rather than leave the last step showing.
            logger.LogWarning("Teams refused the end of the stream ({Message}), so ending it with the text only", ex.Message);
            var plain = new MessageActivityInput { Text = reply.Text, TextFormat = reply.TextFormat, ReplyToId = requestId };
            plain.AddEntity(new StreamInfoEntity { StreamId = _streamId, StreamType = StreamTypes.Final });
            try
            {
                await conversations.CreateActivityAsync(conversationId, plain, cancellationToken: cancellationToken);
                return true;
            }
            catch (HttpRequestException)
            {
                reply.Entities?.Remove(final);
                return false;
            }
        }
        catch (HttpRequestException ex)
        {
            logger.LogWarning("Teams refused the end of the stream ({Status}), so posting the reply normally", ex.StatusCode);
            reply.Entities?.Remove(final);
            return false;
        }
    }

    // The soft stops from https://learn.microsoft.com/microsoftteams/platform/bots/streaming-ux#error-codes.
    private static bool Stopped(HttpRequestException ex) =>
        ex.StatusCode is HttpStatusCode.NoContent or HttpStatusCode.Gone
        || (ex.StatusCode == HttpStatusCode.Forbidden && ex.Message.Contains("cancel", StringComparison.OrdinalIgnoreCase));
}
