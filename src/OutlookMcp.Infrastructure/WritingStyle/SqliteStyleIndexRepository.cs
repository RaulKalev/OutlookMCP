using System.Globalization;
using System.Text.RegularExpressions;
using Microsoft.Data.Sqlite;
using Microsoft.Extensions.Logging;
using OutlookMcp.Application.Configuration;
using OutlookMcp.Application.WritingStyle;
using OutlookMcp.Contracts;

namespace OutlookMcp.Infrastructure.WritingStyle;

public sealed partial class SqliteStyleIndexRepository : IStyleIndexRepository, IDisposable
{
    private readonly ILogger<SqliteStyleIndexRepository> _logger;
    private readonly bool _fullTextEnabled;
    private readonly SemaphoreSlim _initializationLock = new(1, 1);
    private bool _initialized;

    public SqliteStyleIndexRepository(OutlookMcpOptions options, ILogger<SqliteStyleIndexRepository> logger)
    {
        DatabasePath = AppPaths.ExpandPath(options.WritingStyle.DatabasePath);
        _fullTextEnabled = options.WritingStyle.EnableFullTextSearch;
        _logger = logger;
    }

    public string DatabasePath { get; }

    [GeneratedRegex(@"[\p{L}\p{N}_/-]{2,}", RegexOptions.CultureInvariant)]
    private static partial Regex SearchTerms();

    public async Task InitializeAsync(CancellationToken cancellationToken)
    {
        if (_initialized) return;
        await _initializationLock.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            if (_initialized) return;
            Directory.CreateDirectory(Path.GetDirectoryName(DatabasePath)!);
            await using var connection = await OpenAsync(cancellationToken).ConfigureAwait(false);
            await using var command = connection.CreateCommand();
            command.CommandText = Schema;
            await command.ExecuteNonQueryAsync(cancellationToken).ConfigureAwait(false);
            _initialized = true;
        }
        finally { _initializationLock.Release(); }
    }

    public async Task<IReadOnlyList<StyleScanCheckpointDto>> GetCheckpointsAsync(CancellationToken cancellationToken)
    {
        await InitializeAsync(cancellationToken).ConfigureAwait(false);
        await using var connection = await OpenAsync(cancellationToken).ConfigureAwait(false);
        await using var command = connection.CreateCommand();
        command.CommandText = "SELECT store_id, folder_id, folder_path, next_offset, total_discovered, processed, failed, complete, last_processed_at FROM scan_folders ORDER BY folder_path";
        await using var reader = await command.ExecuteReaderAsync(cancellationToken).ConfigureAwait(false);
        var result = new List<StyleScanCheckpointDto>();
        while (await reader.ReadAsync(cancellationToken).ConfigureAwait(false)) result.Add(new(
            reader.GetString(0), reader.GetString(1), reader.GetString(2), reader.GetInt32(3), reader.GetInt32(4), reader.GetInt32(5), reader.GetInt32(6), reader.GetBoolean(7), ReadDate(reader, 8)));
        return result;
    }

    public async Task UpsertCheckpointAsync(StyleScanCheckpointDto checkpoint, CancellationToken cancellationToken)
    {
        await InitializeAsync(cancellationToken).ConfigureAwait(false);
        await using var connection = await OpenAsync(cancellationToken).ConfigureAwait(false);
        await using var command = connection.CreateCommand();
        command.CommandText = """
            INSERT INTO scan_folders(store_id,folder_id,folder_path,next_offset,total_discovered,processed,failed,complete,last_processed_at)
            VALUES($store,$folder,$path,$offset,$total,$processed,$failed,$complete,$last)
            ON CONFLICT(store_id,folder_id) DO UPDATE SET folder_path=excluded.folder_path,next_offset=excluded.next_offset,total_discovered=excluded.total_discovered,processed=excluded.processed,failed=excluded.failed,complete=excluded.complete,last_processed_at=excluded.last_processed_at;
            """;
        Add(command, "$store", checkpoint.StoreId); Add(command, "$folder", checkpoint.FolderId); Add(command, "$path", checkpoint.FolderPath);
        Add(command, "$offset", checkpoint.NextOffset); Add(command, "$total", checkpoint.TotalDiscovered); Add(command, "$processed", checkpoint.Processed);
        Add(command, "$failed", checkpoint.Failed); Add(command, "$complete", checkpoint.Complete); Add(command, "$last", WriteDate(checkpoint.LastProcessedAt));
        await command.ExecuteNonQueryAsync(cancellationToken).ConfigureAwait(false);
    }

    public async Task<(int Inserted, int Updated)> UpsertMessagesAsync(IReadOnlyList<IndexedSentEmailDto> messages, CancellationToken cancellationToken)
    {
        if (messages.Count == 0) return (0, 0);
        await InitializeAsync(cancellationToken).ConfigureAwait(false);
        await using var connection = await OpenAsync(cancellationToken).ConfigureAwait(false);
        await using var transaction = await connection.BeginTransactionAsync(cancellationToken).ConfigureAwait(false);
        await using var insert = CreateInsertCommand(connection, transaction);
        await using var update = CreateUpdateCommand(connection, transaction);
        var inserted = 0;
        var updated = 0;
        foreach (var message in messages)
        {
            BindMessage(insert, message);
            if (await insert.ExecuteNonQueryAsync(cancellationToken).ConfigureAwait(false) == 1) inserted++;
            else
            {
                BindMessage(update, message);
                updated += await update.ExecuteNonQueryAsync(cancellationToken).ConfigureAwait(false);
            }
        }
        await transaction.CommitAsync(cancellationToken).ConfigureAwait(false);
        _logger.LogInformation("Indexed sent-email batch; Inserted={Inserted}, Updated={Updated}", inserted, updated);
        return (inserted, updated);
    }

    public async Task<StyleRepositoryCountsDto> GetCountsAsync(CancellationToken cancellationToken)
    {
        await InitializeAsync(cancellationToken).ConfigureAwait(false);
        await using var connection = await OpenAsync(cancellationToken).ConfigureAwait(false);
        await using var command = connection.CreateCommand();
        command.CommandText = """
            SELECT COUNT(*),
              SUM(CASE WHEN processing_status='successfully_processed' THEN 1 ELSE 0 END),
              SUM(CASE WHEN length(authored_text)>0 THEN 1 ELSE 0 END),
              SUM(CASE WHEN extraction_confidence<0.6 THEN 1 ELSE 0 END),
              SUM(CASE WHEN processing_status IN ('processing_failed','body_unavailable','awaiting_outlook_synchronisation') THEN 1 ELSE 0 END),
              COUNT(*)-COUNT(DISTINCT authored_hash), MIN(sent_at), MAX(sent_at),
              (SELECT COUNT(*) FROM recurring_blocks WHERE block_type='signature' AND occurrence_count>=3),
              (SELECT COUNT(*) FROM recurring_blocks WHERE block_type='disclaimer' AND occurrence_count>=3),
              (SELECT COUNT(*) FROM missing_sent_emails)
            FROM sent_emails;
            """;
        await using var reader = await command.ExecuteReaderAsync(cancellationToken).ConfigureAwait(false);
        await reader.ReadAsync(cancellationToken).ConfigureAwait(false);
        return new(ReadInt(reader, 0), ReadInt(reader, 1), ReadInt(reader, 2), ReadInt(reader, 3), ReadInt(reader, 4), Math.Max(0, ReadInt(reader, 5)), ReadDate(reader, 6), ReadDate(reader, 7), ReadInt(reader, 8), ReadInt(reader, 9), ReadInt(reader, 10));
    }

    public async Task<string?> GetStateAsync(string key, CancellationToken cancellationToken)
    {
        await InitializeAsync(cancellationToken).ConfigureAwait(false);
        await using var connection = await OpenAsync(cancellationToken).ConfigureAwait(false);
        await using var command = connection.CreateCommand(); command.CommandText = "SELECT value FROM style_state WHERE key=$key"; Add(command, "$key", key);
        return Convert.ToString(await command.ExecuteScalarAsync(cancellationToken).ConfigureAwait(false), CultureInfo.InvariantCulture);
    }

    public async Task SetStateAsync(string key, string value, CancellationToken cancellationToken)
    {
        await InitializeAsync(cancellationToken).ConfigureAwait(false);
        await using var connection = await OpenAsync(cancellationToken).ConfigureAwait(false);
        await using var command = connection.CreateCommand(); command.CommandText = "INSERT INTO style_state(key,value) VALUES($key,$value) ON CONFLICT(key) DO UPDATE SET value=excluded.value";
        Add(command, "$key", key); Add(command, "$value", value); await command.ExecuteNonQueryAsync(cancellationToken).ConfigureAwait(false);
    }

    public async Task RebuildRecurringBlocksAsync(CancellationToken cancellationToken)
    {
        await InitializeAsync(cancellationToken).ConfigureAwait(false);
        await using var connection = await OpenAsync(cancellationToken).ConfigureAwait(false);
        await using var transaction = await connection.BeginTransactionAsync(cancellationToken).ConfigureAwait(false);
        await using var command = connection.CreateCommand(); command.Transaction = (SqliteTransaction)transaction;
        command.CommandText = """
            DELETE FROM recurring_blocks;
            INSERT INTO recurring_blocks(block_hash,block_text,block_type,occurrence_count)
              SELECT lower(hex(sha3(signature_text,256))),signature_text,'signature',COUNT(*) FROM sent_emails WHERE length(signature_text)>0 GROUP BY signature_text;
            INSERT INTO recurring_blocks(block_hash,block_text,block_type,occurrence_count)
              SELECT lower(hex(sha3(disclaimer_text,256))),disclaimer_text,'disclaimer',COUNT(*) FROM sent_emails WHERE length(disclaimer_text)>0 GROUP BY disclaimer_text;
            """;
        try { await command.ExecuteNonQueryAsync(cancellationToken).ConfigureAwait(false); }
        catch (SqliteException)
        {
            command.CommandText = """
                DELETE FROM recurring_blocks;
                INSERT INTO recurring_blocks(block_hash,block_text,block_type,occurrence_count)
                  SELECT signature_text,signature_text,'signature',COUNT(*) FROM sent_emails WHERE length(signature_text)>0 GROUP BY signature_text;
                INSERT INTO recurring_blocks(block_hash,block_text,block_type,occurrence_count)
                  SELECT disclaimer_text,disclaimer_text,'disclaimer',COUNT(*) FROM sent_emails WHERE length(disclaimer_text)>0 GROUP BY disclaimer_text;
                """;
            await command.ExecuteNonQueryAsync(cancellationToken).ConfigureAwait(false);
        }
        await transaction.CommitAsync(cancellationToken).ConfigureAwait(false);
    }

    public async Task<StyleStatisticsDto> GetStatisticsAsync(CancellationToken cancellationToken)
    {
        var counts = await GetCountsAsync(cancellationToken).ConfigureAwait(false);
        await using var connection = await OpenAsync(cancellationToken).ConfigureAwait(false);
        var averageCharacters = await ScalarDoubleAsync(connection, "SELECT COALESCE(AVG(authored_length),0) FROM sent_emails WHERE authored_length>0", cancellationToken).ConfigureAwait(false);
        var median = await ScalarIntAsync(connection, "SELECT COALESCE(authored_length,0) FROM sent_emails WHERE authored_length>0 ORDER BY authored_length LIMIT 1 OFFSET (SELECT COUNT(*)/2 FROM sent_emails WHERE authored_length>0)", cancellationToken).ConfigureAwait(false);
        var averageParagraphs = await ScalarDoubleAsync(connection, "SELECT COALESCE(AVG(paragraph_count),0) FROM sent_emails WHERE authored_length>0", cancellationToken).ConfigureAwait(false);
        var questionPercent = counts.Authored == 0 ? 0 : 100d * await ScalarIntAsync(connection, "SELECT COUNT(*) FROM sent_emails WHERE authored_length>0 AND question_count>0", cancellationToken).ConfigureAwait(false) / counts.Authored;
        var listPercent = counts.Authored == 0 ? 0 : 100d * await ScalarIntAsync(connection, "SELECT COUNT(*) FROM sent_emails WHERE authored_length>0 AND list_item_count>0", cancellationToken).ConfigureAwait(false) / counts.Authored;
        return new(counts.Total, counts.Success, counts.Authored, counts.LowConfidence, averageCharacters, median, averageParagraphs, questionPercent, listPercent,
            await GroupCountsAsync(connection, "greeting", cancellationToken).ConfigureAwait(false), await GroupCountsAsync(connection, "closing", cancellationToken).ConfigureAwait(false),
            await GroupCountsAsync(connection, "communication_intent", cancellationToken).ConfigureAwait(false), counts.Oldest, counts.Newest, counts.RecurringSignatures, counts.RecurringDisclaimers, counts.Duplicates);
    }

    public async Task<IReadOnlyList<StyleDatasetExampleDto>> GetRepresentativeExamplesAsync(int maximumExamples, int maximumCharactersPerExample, CancellationToken cancellationToken)
    {
        await InitializeAsync(cancellationToken).ConfigureAwait(false);
        await using var connection = await OpenAsync(cancellationToken).ConfigureAwait(false);
        var usable = await ScalarIntAsync(connection, "SELECT COUNT(*) FROM sent_emails WHERE authored_length>0", cancellationToken).ConfigureAwait(false);
        if (usable == 0) return [];
        var step = Math.Max(1, usable / Math.Max(1, maximumExamples));
        await using var command = connection.CreateCommand();
        command.CommandText = """
            WITH ordered AS (
              SELECT id,message_id,sent_at,subject,to_recipients,cc_recipients,authored_text,authored_length,communication_intent,extraction_confidence,
                     ROW_NUMBER() OVER (ORDER BY sent_at,id) AS rn
              FROM sent_emails WHERE authored_length>0)
            SELECT id,message_id,sent_at,subject,to_recipients,cc_recipients,substr(authored_text,1,$chars),authored_length,communication_intent,extraction_confidence
            FROM ordered WHERE (rn-1) % $step = 0 ORDER BY sent_at LIMIT $limit;
            """;
        Add(command, "$chars", maximumCharactersPerExample); Add(command, "$step", step); Add(command, "$limit", maximumExamples * 2);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken).ConfigureAwait(false);
        var result = new List<StyleDatasetExampleDto>();
        while (await reader.ReadAsync(cancellationToken).ConfigureAwait(false) && result.Count < maximumExamples)
        {
            var sent = ReadDate(reader, 2);
            var length = reader.GetInt32(7);
            result.Add(new(reader.GetInt64(0).ToString(CultureInfo.InvariantCulture), sent, reader.GetString(3), JoinRecipients(reader.GetString(4), reader.GetString(5)), reader.GetString(6), length, reader.GetString(8), reader.GetDouble(9), TimeBucket(sent), LengthBucket(length)));
        }
        return result;
    }

    public async Task<IReadOnlyDictionary<string, int>> GetCommonPhrasesAsync(int maximumPhrases, CancellationToken cancellationToken)
    {
        await InitializeAsync(cancellationToken).ConfigureAwait(false);
        await using var connection = await OpenAsync(cancellationToken).ConfigureAwait(false);
        await using var command = connection.CreateCommand();
        command.CommandText = """
            SELECT trim(substr(authored_text,1,instr(authored_text||char(10),char(10))-1)) AS phrase, COUNT(*) AS uses
            FROM sent_emails WHERE authored_length>0 GROUP BY phrase
            HAVING length(phrase) BETWEEN 2 AND 160 AND uses>=2 ORDER BY uses DESC, phrase LIMIT $limit;
            """;
        Add(command, "$limit", maximumPhrases);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken).ConfigureAwait(false);
        var result = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        while (await reader.ReadAsync(cancellationToken).ConfigureAwait(false)) result[reader.GetString(0)] = reader.GetInt32(1);
        return result;
    }

    public async Task<IReadOnlyList<StyleSearchCandidateDto>> SearchExamplesAsync(string query, int candidateLimit, CancellationToken cancellationToken)
    {
        await InitializeAsync(cancellationToken).ConfigureAwait(false);
        await using var connection = await OpenAsync(cancellationToken).ConfigureAwait(false);
        var fts = _fullTextEnabled ? BuildFtsQuery(query) : string.Empty;
        await using var command = connection.CreateCommand();
        command.CommandText = string.IsNullOrWhiteSpace(fts) ? RecentSearchSql : FtsSearchSql;
        if (!string.IsNullOrWhiteSpace(fts)) Add(command, "$query", fts);
        Add(command, "$limit", candidateLimit);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken).ConfigureAwait(false);
        var result = new List<StyleSearchCandidateDto>();
        while (await reader.ReadAsync(cancellationToken).ConfigureAwait(false)) result.Add(new(
            reader.GetInt64(0), reader.GetString(1), reader.GetString(2), reader.GetString(3), ReadDate(reader, 4), reader.GetString(5), reader.GetString(6), reader.GetString(7), reader.GetString(8), reader.GetString(9), reader.GetString(10), reader.GetString(11), reader.GetString(12), reader.GetDouble(13), reader.GetString(14), reader.GetDouble(15)));
        return result;
    }

    public async Task<IReadOnlySet<string>> GetEntryIdsAsync(string storeId, string folderId, CancellationToken cancellationToken)
    {
        await InitializeAsync(cancellationToken).ConfigureAwait(false);
        await using var connection = await OpenAsync(cancellationToken).ConfigureAwait(false);
        await using var command = connection.CreateCommand();
        command.CommandText = "SELECT entry_id FROM sent_emails WHERE store_id=$store AND folder_id=$folder";
        Add(command, "$store", storeId); Add(command, "$folder", folderId);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken).ConfigureAwait(false);
        var result = new HashSet<string>(StringComparer.Ordinal);
        while (await reader.ReadAsync(cancellationToken).ConfigureAwait(false)) result.Add(reader.GetString(0));
        return result;
    }

    public async Task ReplaceMissingEntryIdsAsync(string storeId, string folderId, IReadOnlyCollection<string> entryIds, CancellationToken cancellationToken)
    {
        await InitializeAsync(cancellationToken).ConfigureAwait(false);
        await using var connection = await OpenAsync(cancellationToken).ConfigureAwait(false);
        await using var transaction = (SqliteTransaction)await connection.BeginTransactionAsync(cancellationToken).ConfigureAwait(false);
        await using (var delete = connection.CreateCommand())
        {
            delete.Transaction = transaction;
            delete.CommandText = "DELETE FROM missing_sent_emails WHERE store_id=$store AND folder_id=$folder";
            Add(delete, "$store", storeId); Add(delete, "$folder", folderId);
            await delete.ExecuteNonQueryAsync(cancellationToken).ConfigureAwait(false);
        }
        await using var insert = connection.CreateCommand();
        insert.Transaction = transaction;
        insert.CommandText = "INSERT INTO missing_sent_emails(store_id,folder_id,entry_id,detected_at) VALUES($store,$folder,$entry,$detected)";
        Add(insert, "$store", storeId); Add(insert, "$folder", folderId); Add(insert, "$entry", string.Empty); Add(insert, "$detected", WriteDate(DateTimeOffset.UtcNow));
        foreach (var entryId in entryIds)
        {
            insert.Parameters["$entry"].Value = entryId;
            await insert.ExecuteNonQueryAsync(cancellationToken).ConfigureAwait(false);
        }
        await transaction.CommitAsync(cancellationToken).ConfigureAwait(false);
    }

    private async Task<SqliteConnection> OpenAsync(CancellationToken cancellationToken)
    {
        var connection = new SqliteConnection(new SqliteConnectionStringBuilder { DataSource = DatabasePath, Mode = SqliteOpenMode.ReadWriteCreate, Cache = SqliteCacheMode.Shared, Pooling = false }.ToString());
        await connection.OpenAsync(cancellationToken).ConfigureAwait(false);
        await using var command = connection.CreateCommand(); command.CommandText = "PRAGMA foreign_keys=ON; PRAGMA journal_mode=WAL; PRAGMA synchronous=NORMAL; PRAGMA busy_timeout=5000;";
        await command.ExecuteNonQueryAsync(cancellationToken).ConfigureAwait(false);
        return connection;
    }

    private static SqliteCommand CreateInsertCommand(SqliteConnection connection, System.Data.Common.DbTransaction transaction)
    {
        var command = connection.CreateCommand(); command.Transaction = (SqliteTransaction)transaction;
        command.CommandText = $"INSERT OR IGNORE INTO sent_emails({Columns}) VALUES({Parameters})"; AddMessageParameters(command); return command;
    }

    private static SqliteCommand CreateUpdateCommand(SqliteConnection connection, System.Data.Common.DbTransaction transaction)
    {
        var command = connection.CreateCommand(); command.Transaction = (SqliteTransaction)transaction;
        command.CommandText = $"UPDATE sent_emails SET {UpdateAssignments} WHERE store_id=$store_id AND entry_id=$entry_id"; AddMessageParameters(command); return command;
    }

    private static void AddMessageParameters(SqliteCommand command)
    {
        foreach (var name in ColumnNames) command.Parameters.Add(new SqliteParameter("$" + name, null));
    }

    private static void BindMessage(SqliteCommand command, IndexedSentEmailDto value)
    {
        var values = new Dictionary<string, object?>(StringComparer.Ordinal)
        {
            ["entry_id"] = value.EntryId,
            ["message_id"] = value.MessageId,
            ["store_id"] = value.StoreId,
            ["folder_id"] = value.FolderId,
            ["folder_path"] = value.FolderPath,
            ["internet_message_id"] = value.InternetMessageId,
            ["conversation_id"] = value.ConversationId,
            ["conversation_topic"] = value.ConversationTopic,
            ["subject"] = value.Subject,
            ["normalised_subject"] = value.NormalisedSubject,
            ["sent_at"] = WriteDate(value.SentAt),
            ["last_modified_at"] = WriteDate(value.LastModifiedAt),
            ["sender_name"] = value.SenderName,
            ["sender_address"] = value.SenderAddress,
            ["to_recipients"] = value.ToRecipients,
            ["cc_recipients"] = value.CcRecipients,
            ["bcc_recipients"] = value.BccRecipients,
            ["recipient_domains"] = value.RecipientDomains,
            ["clean_body"] = value.CleanBody,
            ["authored_text"] = value.AuthoredText,
            ["quoted_text"] = value.QuotedText,
            ["signature_text"] = value.SignatureText,
            ["disclaimer_text"] = value.DisclaimerText,
            ["body_preview"] = value.BodyPreview,
            ["attachment_names"] = value.AttachmentNames,
            ["attachment_extensions"] = value.AttachmentExtensions,
            ["project_keywords"] = value.ProjectKeywords,
            ["detected_entities"] = value.DetectedEntities,
            ["communication_intent"] = value.CommunicationIntent,
            ["processing_status"] = value.ProcessingStatus,
            ["processing_reason"] = value.ProcessingReason,
            ["extraction_confidence"] = value.ExtractionConfidence,
            ["indexed_at"] = WriteDate(value.IndexedAt),
            ["content_hash"] = value.ContentHash,
            ["authored_hash"] = value.AuthoredHash,
            ["greeting"] = value.Greeting,
            ["closing"] = value.Closing,
            ["paragraph_count"] = value.ParagraphCount,
            ["list_item_count"] = value.ListItemCount,
            ["question_count"] = value.QuestionCount,
            ["authored_length"] = value.AuthoredText.Length
        };
        foreach (var (name, item) in values) command.Parameters["$" + name].Value = item ?? DBNull.Value;
    }

    private static async Task<int> ScalarIntAsync(SqliteConnection connection, string sql, CancellationToken cancellationToken)
    { await using var command = connection.CreateCommand(); command.CommandText = sql; return Convert.ToInt32(await command.ExecuteScalarAsync(cancellationToken).ConfigureAwait(false) ?? 0, CultureInfo.InvariantCulture); }
    private static async Task<double> ScalarDoubleAsync(SqliteConnection connection, string sql, CancellationToken cancellationToken)
    { await using var command = connection.CreateCommand(); command.CommandText = sql; return Convert.ToDouble(await command.ExecuteScalarAsync(cancellationToken).ConfigureAwait(false) ?? 0d, CultureInfo.InvariantCulture); }
    private static async Task<IReadOnlyDictionary<string, int>> GroupCountsAsync(SqliteConnection connection, string column, CancellationToken cancellationToken)
    {
        await using var command = connection.CreateCommand(); command.CommandText = $"SELECT {column},COUNT(*) FROM sent_emails WHERE length({column})>0 GROUP BY {column} ORDER BY COUNT(*) DESC LIMIT 20";
        await using var reader = await command.ExecuteReaderAsync(cancellationToken).ConfigureAwait(false); var result = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        while (await reader.ReadAsync(cancellationToken).ConfigureAwait(false)) result[reader.GetString(0)] = reader.GetInt32(1); return result;
    }

    private static string BuildFtsQuery(string query) => string.Join(" OR ", SearchTerms().Matches(query).Select(match => '"' + match.Value.Replace("\"", "\"\"", StringComparison.Ordinal) + '"'));
    private static string JoinRecipients(string to, string cc) => string.Join("; ", new[] { to, cc }.Where(value => !string.IsNullOrWhiteSpace(value)));
    private static string TimeBucket(DateTimeOffset? sent) => sent is null ? "unknown" : sent.Value.Year.ToString(CultureInfo.InvariantCulture);
    private static string LengthBucket(int length) => length < 200 ? "short" : length < 1_000 ? "medium" : "long";
    private static void Add(SqliteCommand command, string name, object? value) => command.Parameters.AddWithValue(name, value ?? DBNull.Value);
    private static int ReadInt(SqliteDataReader reader, int ordinal) => reader.IsDBNull(ordinal) ? 0 : Convert.ToInt32(reader.GetValue(ordinal), CultureInfo.InvariantCulture);
    private static DateTimeOffset? ReadDate(SqliteDataReader reader, int ordinal) => reader.IsDBNull(ordinal) ? null : DateTimeOffset.TryParse(reader.GetString(ordinal), CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out var value) ? value : null;
    private static string? WriteDate(DateTimeOffset? value) => value?.ToString("O", CultureInfo.InvariantCulture);

    private const string Columns = "entry_id,message_id,store_id,folder_id,folder_path,internet_message_id,conversation_id,conversation_topic,subject,normalised_subject,sent_at,last_modified_at,sender_name,sender_address,to_recipients,cc_recipients,bcc_recipients,recipient_domains,clean_body,authored_text,quoted_text,signature_text,disclaimer_text,body_preview,attachment_names,attachment_extensions,project_keywords,detected_entities,communication_intent,processing_status,processing_reason,extraction_confidence,indexed_at,content_hash,authored_hash,greeting,closing,paragraph_count,list_item_count,question_count,authored_length";
    private static readonly string[] ColumnNames = Columns.Split(',');
    private static readonly string Parameters = string.Join(',', ColumnNames.Select(value => "$" + value));
    private static readonly string UpdateAssignments = string.Join(',', ColumnNames.Where(value => value is not "entry_id" and not "store_id").Select(value => value + "=$" + value));

    private const string FtsSearchSql = """
        SELECT e.id,e.entry_id,e.message_id,e.store_id,e.sent_at,e.subject,e.to_recipients,e.cc_recipients,e.recipient_domains,e.authored_text,e.clean_body,e.project_keywords,e.communication_intent,e.extraction_confidence,e.content_hash,-bm25(sent_email_fts,5.0,4.0,1.0,2.0,2.0,1.0,1.0,1.0,2.0)
        FROM sent_email_fts JOIN sent_emails e ON e.id=sent_email_fts.rowid
        WHERE sent_email_fts MATCH $query AND e.authored_length>0 ORDER BY bm25(sent_email_fts) LIMIT $limit;
        """;
    private const string RecentSearchSql = """
        SELECT id,entry_id,message_id,store_id,sent_at,subject,to_recipients,cc_recipients,recipient_domains,authored_text,clean_body,project_keywords,communication_intent,extraction_confidence,content_hash,0.0
        FROM sent_emails WHERE authored_length>0 ORDER BY sent_at DESC LIMIT $limit;
        """;

    private const string Schema = """
        CREATE TABLE IF NOT EXISTS sent_emails(
          id INTEGER PRIMARY KEY, entry_id TEXT NOT NULL, message_id TEXT NOT NULL, store_id TEXT NOT NULL, folder_id TEXT NOT NULL, folder_path TEXT NOT NULL,
          internet_message_id TEXT, conversation_id TEXT, conversation_topic TEXT, subject TEXT NOT NULL, normalised_subject TEXT NOT NULL,
          sent_at TEXT, last_modified_at TEXT, sender_name TEXT, sender_address TEXT, to_recipients TEXT NOT NULL, cc_recipients TEXT NOT NULL, bcc_recipients TEXT NOT NULL,
          recipient_domains TEXT NOT NULL, clean_body TEXT NOT NULL, authored_text TEXT NOT NULL, quoted_text TEXT NOT NULL, signature_text TEXT NOT NULL,
          disclaimer_text TEXT NOT NULL, body_preview TEXT NOT NULL, attachment_names TEXT NOT NULL, attachment_extensions TEXT NOT NULL,
          project_keywords TEXT NOT NULL, detected_entities TEXT NOT NULL, communication_intent TEXT NOT NULL, processing_status TEXT NOT NULL, processing_reason TEXT,
          extraction_confidence REAL NOT NULL, indexed_at TEXT NOT NULL, content_hash TEXT NOT NULL, authored_hash TEXT NOT NULL,
          greeting TEXT, closing TEXT, paragraph_count INTEGER NOT NULL, list_item_count INTEGER NOT NULL, question_count INTEGER NOT NULL, authored_length INTEGER NOT NULL,
          UNIQUE(store_id,entry_id));
        CREATE INDEX IF NOT EXISTS ix_sent_email_sent_at ON sent_emails(sent_at);
        CREATE INDEX IF NOT EXISTS ix_sent_email_modified ON sent_emails(last_modified_at);
        CREATE INDEX IF NOT EXISTS ix_sent_email_internet_id ON sent_emails(internet_message_id);
        CREATE INDEX IF NOT EXISTS ix_sent_email_content_hash ON sent_emails(content_hash);
        CREATE INDEX IF NOT EXISTS ix_sent_email_authored_hash ON sent_emails(authored_hash);
        CREATE INDEX IF NOT EXISTS ix_sent_email_intent ON sent_emails(communication_intent);
        CREATE TABLE IF NOT EXISTS scan_folders(store_id TEXT NOT NULL,folder_id TEXT NOT NULL,folder_path TEXT NOT NULL,next_offset INTEGER NOT NULL,total_discovered INTEGER NOT NULL,processed INTEGER NOT NULL,failed INTEGER NOT NULL,complete INTEGER NOT NULL,last_processed_at TEXT,PRIMARY KEY(store_id,folder_id));
        CREATE TABLE IF NOT EXISTS style_state(key TEXT PRIMARY KEY,value TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS recurring_blocks(block_hash TEXT NOT NULL,block_text TEXT NOT NULL,block_type TEXT NOT NULL,occurrence_count INTEGER NOT NULL,PRIMARY KEY(block_hash,block_type));
        CREATE TABLE IF NOT EXISTS missing_sent_emails(store_id TEXT NOT NULL,folder_id TEXT NOT NULL,entry_id TEXT NOT NULL,detected_at TEXT NOT NULL,PRIMARY KEY(store_id,folder_id,entry_id));
        CREATE VIRTUAL TABLE IF NOT EXISTS sent_email_fts USING fts5(subject,authored_text,clean_body,to_recipients,cc_recipients,recipient_domains,attachment_names,conversation_topic,project_keywords,content='sent_emails',content_rowid='id',tokenize='unicode61 remove_diacritics 2');
        CREATE TRIGGER IF NOT EXISTS sent_email_ai AFTER INSERT ON sent_emails BEGIN INSERT INTO sent_email_fts(rowid,subject,authored_text,clean_body,to_recipients,cc_recipients,recipient_domains,attachment_names,conversation_topic,project_keywords) VALUES(new.id,new.subject,new.authored_text,new.clean_body,new.to_recipients,new.cc_recipients,new.recipient_domains,new.attachment_names,new.conversation_topic,new.project_keywords); END;
        CREATE TRIGGER IF NOT EXISTS sent_email_ad AFTER DELETE ON sent_emails BEGIN INSERT INTO sent_email_fts(sent_email_fts,rowid,subject,authored_text,clean_body,to_recipients,cc_recipients,recipient_domains,attachment_names,conversation_topic,project_keywords) VALUES('delete',old.id,old.subject,old.authored_text,old.clean_body,old.to_recipients,old.cc_recipients,old.recipient_domains,old.attachment_names,old.conversation_topic,old.project_keywords); END;
        CREATE TRIGGER IF NOT EXISTS sent_email_au AFTER UPDATE ON sent_emails BEGIN INSERT INTO sent_email_fts(sent_email_fts,rowid,subject,authored_text,clean_body,to_recipients,cc_recipients,recipient_domains,attachment_names,conversation_topic,project_keywords) VALUES('delete',old.id,old.subject,old.authored_text,old.clean_body,old.to_recipients,old.cc_recipients,old.recipient_domains,old.attachment_names,old.conversation_topic,old.project_keywords); INSERT INTO sent_email_fts(rowid,subject,authored_text,clean_body,to_recipients,cc_recipients,recipient_domains,attachment_names,conversation_topic,project_keywords) VALUES(new.id,new.subject,new.authored_text,new.clean_body,new.to_recipients,new.cc_recipients,new.recipient_domains,new.attachment_names,new.conversation_topic,new.project_keywords); END;
        """;

    public void Dispose() => _initializationLock.Dispose();
}
