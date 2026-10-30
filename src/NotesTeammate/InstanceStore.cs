using System.Text.Json;

namespace NotesTeammate;

/// <summary>Per-instance storage. Every call takes an <see cref="InstanceKey"/>; there is no way to list or read across instances.</summary>
public interface IInstanceStore
{
    Task<InstanceRecord?> GetAsync(InstanceKey key, CancellationToken cancellationToken = default);

    /// <summary>Loads the instance's record (or starts a new one), applies <paramref name="change"/> and saves it.</summary>
    Task<InstanceRecord> UpdateAsync(InstanceKey key, Action<InstanceRecord> change, CancellationToken cancellationToken = default);
}

/// <summary>
/// One JSON file per instance, at <c>{folder}/{tenantId}/{agentUserId}.json</c>. Good enough to see the shape. In
/// production use a database partitioned by the same key, for example Cosmos DB with a tenantId/agentUserId
/// hierarchical partition key.
/// </summary>
public sealed class FileInstanceStore(string folder) : IInstanceStore
{
    private static readonly JsonSerializerOptions Json = new(JsonSerializerDefaults.Web) { WriteIndented = true };

    // One lock for the whole store keeps the sample simple. Updates are short.
    private readonly SemaphoreSlim _lock = new(1, 1);

    public async Task<InstanceRecord?> GetAsync(InstanceKey key, CancellationToken cancellationToken = default)
    {
        await _lock.WaitAsync(cancellationToken);
        try
        {
            return await ReadAsync(key, cancellationToken);
        }
        finally
        {
            _lock.Release();
        }
    }

    public async Task<InstanceRecord> UpdateAsync(InstanceKey key, Action<InstanceRecord> change, CancellationToken cancellationToken = default)
    {
        await _lock.WaitAsync(cancellationToken);
        try
        {
            var record = await ReadAsync(key, cancellationToken) ?? new InstanceRecord { Key = key };
            change(record);

            var path = PathOf(key);
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            var temp = path + ".tmp";
            await File.WriteAllTextAsync(temp, JsonSerializer.Serialize(record, Json), cancellationToken);
            File.Move(temp, path, overwrite: true);
            return record;
        }
        finally
        {
            _lock.Release();
        }
    }

    private async Task<InstanceRecord?> ReadAsync(InstanceKey key, CancellationToken cancellationToken)
    {
        var path = PathOf(key);
        if (!File.Exists(path))
        {
            return null;
        }

        await using var stream = File.OpenRead(path);
        return await JsonSerializer.DeserializeAsync<InstanceRecord>(stream, Json, cancellationToken);
    }

    private string PathOf(InstanceKey key) =>
        Path.Combine(folder, key.TenantId.ToString("D"), key.AgentUserId.ToString("D") + ".json");
}
