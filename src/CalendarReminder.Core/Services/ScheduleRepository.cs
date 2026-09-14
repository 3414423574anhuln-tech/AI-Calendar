using CalendarReminder.Core.Models;
using Microsoft.Data.Sqlite;
using System.Globalization;

namespace CalendarReminder.Core.Services;

public sealed class ScheduleRepository
{
    private const string DbFormat = "O";
    public string DatabasePath { get; }
    private string ConnectionString => new SqliteConnectionStringBuilder { DataSource = DatabasePath, Mode = SqliteOpenMode.ReadWriteCreate }.ToString();

    public ScheduleRepository(string? databasePath = null) => DatabasePath = databasePath ?? AppPaths.DatabasePath;

    public async Task InitializeAsync()
    {
        Directory.CreateDirectory(Path.GetDirectoryName(DatabasePath)!);
        await using var connection = new SqliteConnection(ConnectionString);
        await connection.OpenAsync();
        var command = connection.CreateCommand();
        command.CommandText = """
            PRAGMA journal_mode=WAL;
            CREATE TABLE IF NOT EXISTS Schedules(
              Id TEXT PRIMARY KEY, ScheduledAt TEXT NOT NULL, Content TEXT NOT NULL,
              CreatedAt TEXT NOT NULL, UpdatedAt TEXT NOT NULL, ReminderTriggered INTEGER NOT NULL DEFAULT 0,
              ReminderAt TEXT NOT NULL);
            CREATE INDEX IF NOT EXISTS IX_Schedules_ScheduledAt ON Schedules(ScheduledAt);
            """;
        await command.ExecuteNonQueryAsync();
        var columns = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var schema = connection.CreateCommand(); schema.CommandText = "PRAGMA table_info(Schedules)";
        await using (var reader = await schema.ExecuteReaderAsync()) while (await reader.ReadAsync()) columns.Add(reader.GetString(1));
        if (!columns.Contains("ReminderAt"))
        {
            var migration = connection.CreateCommand();
            migration.CommandText = "ALTER TABLE Schedules ADD COLUMN ReminderAt TEXT; UPDATE Schedules SET ReminderAt=ScheduledAt WHERE ReminderAt IS NULL OR ReminderAt='';";
            await migration.ExecuteNonQueryAsync();
        }
        var reminderIndex = connection.CreateCommand();
        reminderIndex.CommandText = "DROP INDEX IF EXISTS IX_Schedules_Reminder; CREATE INDEX IX_Schedules_Reminder ON Schedules(ReminderTriggered, ReminderAt);";
        await reminderIndex.ExecuteNonQueryAsync();
    }

    public async Task<List<ScheduleItem>> GetAllAsync() => await QueryAsync("SELECT * FROM Schedules ORDER BY ScheduledAt, CreatedAt");

    public async Task<List<ScheduleItem>> GetRangeAsync(DateTime start, DateTime end)
    {
        await using var c = await OpenAsync();
        var cmd = c.CreateCommand();
        cmd.CommandText = "SELECT * FROM Schedules WHERE ScheduledAt >= $s AND ScheduledAt <= $e ORDER BY ScheduledAt, CreatedAt";
        cmd.Parameters.AddWithValue("$s", start.ToString(DbFormat, CultureInfo.InvariantCulture));
        cmd.Parameters.AddWithValue("$e", end.ToString(DbFormat, CultureInfo.InvariantCulture));
        return await ReadAsync(cmd);
    }

    public async Task AddAsync(ScheduleItem item)
    {
        item.Content = item.Content.Trim(); if (item.ReminderAt == default) item.ReminderAt = item.ScheduledAt; item.CreatedAt = item.UpdatedAt = DateTime.Now;
        await using var c = await OpenAsync(); await InsertAsync(c, null, item);
    }

    public async Task UpdateAsync(ScheduleItem item)
    {
        await using var c = await OpenAsync();
        if (item.ReminderAt == default) item.ReminderAt = item.ScheduledAt;
        var previous = c.CreateCommand(); previous.CommandText = "SELECT ScheduledAt,ReminderAt FROM Schedules WHERE Id=$id"; previous.Parameters.AddWithValue("$id", item.Id.ToString());
        await using (var reader = await previous.ExecuteReaderAsync())
            if (await reader.ReadAsync() && (Parse(reader.GetString(0)) != item.ScheduledAt || Parse(reader.GetString(1)) != item.ReminderAt)) item.ReminderTriggered = false;
        item.Content = item.Content.Trim(); item.UpdatedAt = DateTime.Now;
        var cmd = c.CreateCommand();
        cmd.CommandText = "UPDATE Schedules SET ScheduledAt=$at,ReminderAt=$reminder,Content=$content,UpdatedAt=$updated,ReminderTriggered=$triggered WHERE Id=$id";
        Bind(cmd, item); await cmd.ExecuteNonQueryAsync();
    }

    public async Task DeleteAsync(Guid id)
    {
        await using var c = await OpenAsync(); var cmd = c.CreateCommand(); cmd.CommandText = "DELETE FROM Schedules WHERE Id=$id"; cmd.Parameters.AddWithValue("$id", id.ToString()); await cmd.ExecuteNonQueryAsync();
    }

    public async Task<int> DeleteManyAsync(IEnumerable<Guid> ids)
    {
        var values = ids.Distinct().ToList(); if (values.Count == 0) return 0;
        await using var c = await OpenAsync(); await using var tx = (SqliteTransaction)await c.BeginTransactionAsync(); var deleted = 0;
        try
        {
            foreach (var id in values)
            {
                var cmd = c.CreateCommand(); cmd.Transaction = tx; cmd.CommandText = "DELETE FROM Schedules WHERE Id=$id"; cmd.Parameters.AddWithValue("$id", id.ToString()); deleted += await cmd.ExecuteNonQueryAsync();
            }
            await tx.CommitAsync(); return deleted;
        }
        catch { await tx.RollbackAsync(); throw; }
    }

    public async Task<List<ScheduleItem>> GetDueAsync(DateTime now)
    {
        await using var c = await OpenAsync(); var cmd = c.CreateCommand();
        cmd.CommandText = "SELECT * FROM Schedules WHERE ReminderTriggered=0 AND ReminderAt <= $now ORDER BY ReminderAt,ScheduledAt";
        cmd.Parameters.AddWithValue("$now", now.ToString(DbFormat, CultureInfo.InvariantCulture)); return await ReadAsync(cmd);
    }

    public async Task<List<ScheduleItem>> GetDueRangeAsync(DateTime startExclusive, DateTime endInclusive)
    {
        await using var c = await OpenAsync(); var cmd = c.CreateCommand();
        cmd.CommandText = "SELECT * FROM Schedules WHERE ReminderTriggered=0 AND ReminderAt > $start AND ReminderAt <= $end ORDER BY ReminderAt,ScheduledAt,CreatedAt";
        cmd.Parameters.AddWithValue("$start", startExclusive.ToString(DbFormat, CultureInfo.InvariantCulture));
        cmd.Parameters.AddWithValue("$end", endInclusive.ToString(DbFormat, CultureInfo.InvariantCulture));
        return await ReadAsync(cmd);
    }

    public async Task MarkTriggeredAsync(Guid id)
    {
        await using var c = await OpenAsync(); var cmd = c.CreateCommand(); cmd.CommandText = "UPDATE Schedules SET ReminderTriggered=1,UpdatedAt=$u WHERE Id=$id"; cmd.Parameters.AddWithValue("$id", id.ToString()); cmd.Parameters.AddWithValue("$u", DateTime.Now.ToString(DbFormat, CultureInfo.InvariantCulture)); await cmd.ExecuteNonQueryAsync();
    }

    public async Task MarkTriggeredAsync(IEnumerable<Guid> ids)
    {
        var values = ids.Distinct().ToList(); if (values.Count == 0) return;
        await using var c = await OpenAsync(); await using var tx = (SqliteTransaction)await c.BeginTransactionAsync();
        try
        {
            foreach (var id in values)
            {
                var cmd = c.CreateCommand(); cmd.Transaction = tx; cmd.CommandText = "UPDATE Schedules SET ReminderTriggered=1,UpdatedAt=$u WHERE Id=$id";
                cmd.Parameters.AddWithValue("$id", id.ToString()); cmd.Parameters.AddWithValue("$u", DateTime.Now.ToString(DbFormat, CultureInfo.InvariantCulture)); await cmd.ExecuteNonQueryAsync();
            }
            await tx.CommitAsync();
        }
        catch { await tx.RollbackAsync(); throw; }
    }

    public async Task<ImportResult> BatchImportAsync(IEnumerable<ImportCandidate> candidates, DuplicateAction action)
    {
        await using var c = await OpenAsync(); await using var tx = (SqliteTransaction)await c.BeginTransactionAsync();
        int imported = 0, skipped = 0, replaced = 0;
        try
        {
            foreach (var row in candidates.Where(x => x.IsSelected))
            {
                if (row.Status == ImportStatus.高度重合) { skipped++; continue; }
                if (row.ScheduledAt is null || string.IsNullOrWhiteSpace(row.Content)) throw new InvalidDataException("导入内容包含无效日期或空安排。");
                var duplicateId = await FindDuplicateAsync(c, tx, row.ScheduledAt.Value, row.Content.Trim());
                if (duplicateId.HasValue && action == DuplicateAction.跳过重复项) { skipped++; continue; }
                if (duplicateId.HasValue && action == DuplicateAction.替换原日程)
                {
                    var del = c.CreateCommand(); del.Transaction = tx; del.CommandText = "DELETE FROM Schedules WHERE Id=$id"; del.Parameters.AddWithValue("$id", duplicateId.Value.ToString()); await del.ExecuteNonQueryAsync(); replaced++;
                }
                var item = new ScheduleItem { ScheduledAt = row.ScheduledAt.Value, ReminderAt = row.ReminderAt ?? row.ScheduledAt.Value, Content = row.Content.Trim() };
                await InsertAsync(c, tx, item); imported++;
            }
            await tx.CommitAsync(); return new(imported, skipped, replaced);
        }
        catch { await tx.RollbackAsync(); throw; }
    }

    public async Task MarkDuplicatesAsync(IEnumerable<ImportCandidate> candidates)
    {
        await using var c = await OpenAsync();
        foreach (var row in candidates.Where(x => x.ScheduledAt.HasValue && !string.IsNullOrWhiteSpace(x.Content)))
        {
            row.IsDuplicate = (await FindDuplicateAsync(c, null, row.ScheduledAt!.Value, row.Content.Trim())).HasValue;
            if (row.IsDuplicate && row.Status != ImportStatus.高度重合) row.Status = ImportStatus.疑似重复;
        }
    }

    private async Task<Guid?> FindDuplicateAsync(SqliteConnection c, SqliteTransaction? tx, DateTime at, string content)
    {
        var cmd = c.CreateCommand(); cmd.Transaction = tx; cmd.CommandText = "SELECT Id FROM Schedules WHERE ScheduledAt=$at AND TRIM(Content)=$content LIMIT 1";
        cmd.Parameters.AddWithValue("$at", at.ToString(DbFormat, CultureInfo.InvariantCulture)); cmd.Parameters.AddWithValue("$content", content.Trim());
        var value = (string?)await cmd.ExecuteScalarAsync(); return value is null ? null : Guid.Parse(value);
    }

    private async Task<SqliteConnection> OpenAsync() { var c = new SqliteConnection(ConnectionString); await c.OpenAsync(); return c; }
    private async Task<List<ScheduleItem>> QueryAsync(string sql) { await using var c = await OpenAsync(); var cmd = c.CreateCommand(); cmd.CommandText = sql; return await ReadAsync(cmd); }
    private static async Task<List<ScheduleItem>> ReadAsync(SqliteCommand cmd)
    {
        var result = new List<ScheduleItem>(); await using var reader = await cmd.ExecuteReaderAsync();
        while (await reader.ReadAsync()) result.Add(new ScheduleItem { Id=Guid.Parse(reader.GetString(0)), ScheduledAt=Parse(reader.GetString(1)), Content=reader.GetString(2), CreatedAt=Parse(reader.GetString(3)), UpdatedAt=Parse(reader.GetString(4)), ReminderTriggered=reader.GetInt32(5)!=0, ReminderAt=reader.IsDBNull(6)?Parse(reader.GetString(1)):Parse(reader.GetString(6)) });
        return result;
    }
    private static DateTime Parse(string value) => DateTime.Parse(value, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind);
    private static void Bind(SqliteCommand cmd, ScheduleItem i) { cmd.Parameters.AddWithValue("$id", i.Id.ToString()); cmd.Parameters.AddWithValue("$at", i.ScheduledAt.ToString(DbFormat, CultureInfo.InvariantCulture)); cmd.Parameters.AddWithValue("$reminder", (i.ReminderAt==default?i.ScheduledAt:i.ReminderAt).ToString(DbFormat, CultureInfo.InvariantCulture)); cmd.Parameters.AddWithValue("$content", i.Content); cmd.Parameters.AddWithValue("$updated", i.UpdatedAt.ToString(DbFormat, CultureInfo.InvariantCulture)); cmd.Parameters.AddWithValue("$triggered", i.ReminderTriggered ? 1 : 0); }
    private static async Task InsertAsync(SqliteConnection c, SqliteTransaction? tx, ScheduleItem i) { if(i.ReminderAt==default)i.ReminderAt=i.ScheduledAt; var cmd=c.CreateCommand(); cmd.Transaction=tx; cmd.CommandText="INSERT INTO Schedules(Id,ScheduledAt,Content,CreatedAt,UpdatedAt,ReminderTriggered,ReminderAt) VALUES($id,$at,$content,$created,$updated,$triggered,$reminder)"; Bind(cmd,i); cmd.Parameters.AddWithValue("$created",i.CreatedAt.ToString(DbFormat,CultureInfo.InvariantCulture)); await cmd.ExecuteNonQueryAsync(); }
}
