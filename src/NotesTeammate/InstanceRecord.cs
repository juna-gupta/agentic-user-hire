using System.Text.Json.Serialization;

namespace NotesTeammate;

/// <summary>Everything one hired instance knows. Nothing here is shared with other instances.</summary>
public sealed class InstanceRecord
{
    public required InstanceKey Key { get; init; }

    /// <summary>Who hired this instance (its manager), from the hire event. Null until that event is handled.</summary>
    public string? HiredBy { get; set; }

    public DateTimeOffset? HiredAt { get; set; }

    /// <summary>False while an admin has the agent user disabled.</summary>
    public bool Enabled { get; set; } = true;

    /// <summary>
    /// When the agent user was deleted. The notes are gone by then; the record is kept only so a message that reaches
    /// the deleted instance later is ignored instead of starting a new record. In production, expire these after the
    /// 30 days a deleted user can be restored.
    /// </summary>
    public DateTimeOffset? RemovedAt { get; set; }

    /// <summary>True when the instance should answer: not paused and not removed.</summary>
    [JsonIgnore]
    public bool Active => Enabled && RemovedAt is null;

    public bool Greeted { get; set; }

    public List<Note> Notes { get; init; } = [];
}

public sealed record Note(string Text, string? From, DateTimeOffset At);
