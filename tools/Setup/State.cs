using System.Text.Json;

namespace Setup;

/// <summary>IDs of what this tool created, in tools/Setup/state.json. No secrets.</summary>
internal sealed class State
{
    private static readonly JsonSerializerOptions Json = new(JsonSerializerDefaults.Web) { WriteIndented = true };

    private string _path = string.Empty;

    public string? TenantId { get; set; }

    public string? BlueprintAppId { get; set; }

    public string? BlueprintObjectId { get; set; }

    public string? BlueprintPrincipalId { get; set; }

    public string? AppId { get; set; }

    public string? TemplateId { get; set; }

    public List<HiredInstance> Hires { get; set; } = [];

    public static State Load(string path)
    {
        var state = File.Exists(path) ? JsonSerializer.Deserialize<State>(File.ReadAllText(path), Json) ?? new State() : new State();
        state._path = path;
        return state;
    }

    public void Save() => File.WriteAllText(_path, JsonSerializer.Serialize(this, Json));
}

internal sealed record HiredInstance(string Name, string AgentIdentityId, string AgentUserId, string UserPrincipalName);
