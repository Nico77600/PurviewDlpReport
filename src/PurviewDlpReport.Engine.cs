// =============================================================================
//  PurviewDlpReport.Engine.cs
// -----------------------------------------------------------------------------
//  Author  : Nicolas Fabert
//  Version : 2.1.1
//  Purpose : Performance-critical part of the Purview DLP Report tool.
//            This file is compiled automatically by PurviewDlpReport.psm1 the
//            first time it is used (and again whenever it changes). The
//            compiled DLL is cached in the .\bin folder.
//
//  Contents
//    1. EventNormalizer : turns one Activity Explorer JSON record into a DlpEvent
//    2. Coverage        : interval arithmetic used to find missing periods
//    3. DlpStore        : SQLite database (schema, ingestion, coverage, queries)
//    4. ReportWriter    : CSV and HTML files, split into parts
//
//  Conventions
//    - Every time stored in the database is a Unix epoch in MILLISECONDS (UTC).
//    - Time ranges are always [start, end) : start included, end excluded.
//    - The database keeps ONE row per Activity Explorer event (RecordIdentity)
//      and ONE row per message (MessageID). Reports show one row per message.
// =============================================================================
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.RegularExpressions;
using Microsoft.Data.Sqlite;

namespace PurviewDlpReport
{
    // -------------------------------------------------------------------------
    // 0. Console
    // -------------------------------------------------------------------------

    /// <summary>
    /// Font of the classic Windows console (conhost). The console has no font fallback, so the
    /// module chooses the frame characters from it. Returns null when the output is not a classic
    /// console window (Windows Terminal, VS Code, redirected output) or on any error.
    /// </summary>
    public static class ConsoleFont
    {
        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential, CharSet = System.Runtime.InteropServices.CharSet.Unicode)]
        struct ConsoleFontInfoEx
        {
            public uint Size; public uint Font; public short Width; public short Height; public int Family; public int Weight;
            [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst = 32)] public string FaceName;
        }

        [System.Runtime.InteropServices.DllImport("kernel32.dll", SetLastError = true)]
        static extern IntPtr GetStdHandle(int handle);

        [System.Runtime.InteropServices.DllImport("kernel32.dll", SetLastError = true)]
        static extern bool GetCurrentConsoleFontEx(IntPtr output, bool maximumWindow, ref ConsoleFontInfoEx info);

        public static string FaceName()
        {
            try
            {
                if (Console.IsOutputRedirected || !OperatingSystem.IsWindows()) return null;
                var info = new ConsoleFontInfoEx { Size = (uint)System.Runtime.InteropServices.Marshal.SizeOf<ConsoleFontInfoEx>() };
                return GetCurrentConsoleFontEx(GetStdHandle(-11), false, ref info) ? info.FaceName : null;
            }
            catch { return null; }
        }
    }

    // -------------------------------------------------------------------------
    // 1. Normalization of Activity Explorer records
    // -------------------------------------------------------------------------

    /// <summary>One DLP rule match, reduced to the fields the report needs.</summary>
    public sealed class DlpEvent
    {
        public string RecordIdentity;
        public string MessageId;
        public long HappenedMs;
        public string Sender;
        public string Subject;
        public string Recipients;          // combined list (To/Cc/Bcc are not separated by the source), "; " separated
        public int RecipientListLength;
        public string PolicyName;
        public string RuleName;
        public long? RecipientCount;       // null when the source does not report a reliable value
        public string CountSource;         // where RecipientCount comes from, or why it is null
    }

    public sealed class NormalizationResult
    {
        public DlpEvent Event;
        public string Problem;             // not null => the record is quarantined
    }

    public static class EventNormalizer
    {
        // Activity Explorer reports the observed value of the rule condition as text,
        // for example "RecipientCountOver 38" for a message sent to 38 recipients.
        // Validated against the audit log RecipientCount on 149,800 messages (lab, 2026-09-29).
        static readonly Regex RecipientCountOver = new Regex(@"\A\s*RecipientCountOver\s+([0-9]{1,9})\s*\z",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);

        public static NormalizationResult Normalize(JsonElement record, string policyId, string ruleId)
        {
            if (record.ValueKind != JsonValueKind.Object) return Fail("RecordIsNotAnObject");

            string recordIdentity = Text(record, "RecordIdentity");
            if (string.IsNullOrWhiteSpace(recordIdentity)) return Fail("MissingRecordIdentity");

            string workload = Text(record, "Workload");
            if (!Same(workload, "Exchange")) return Fail("UnexpectedWorkload:" + (workload ?? "null"));

            string activityId = Text(record, "ActivityId");
            string activity = Text(record, "Activity");
            bool isRuleMatch = activityId != null
                ? Same(activityId, "DLPRuleMatch")
                : (Same(activity, "DLPRuleMatch") || Same(activity, "DLP rule matched"));
            if (!isRuleMatch) return Fail("UnexpectedActivity:" + (activityId ?? activity ?? "null"));

            long happened;
            if (!TryParseTime(Text(record, "Happened"), out happened)) return Fail("MissingOrInvalidHappened");

            JsonElement email = ObjectOrParsedJson(record, "EmailInfo");
            string messageId = Text(email, "MessageID") ?? Text(record, "MessageID");
            if (string.IsNullOrWhiteSpace(messageId)) return Fail("MissingMessageID");

            // Only the policy/rule configured in the tool is trusted as evidence.
            bool targetFound = false;
            string policyName = null, ruleName = null;
            var counts = new SortedSet<long>();
            bool invalidCount = false;
            bool countFromConditions = false;
            foreach (JsonElement entry in PolicyEntries(record))
            {
                if (!Same(Text(entry, "PolicyId"), policyId) || !Same(Text(entry, "RuleId"), ruleId)) continue;
                targetFound = true;
                policyName = policyName ?? Text(entry, "PolicyName");
                ruleName = ruleName ?? Text(entry, "RuleName");
                JsonElement conditions;
                if (!entry.TryGetProperty("OtherConditions", out conditions)) continue;
                foreach (string condition in ConditionTexts(conditions, 0))
                {
                    if (condition.IndexOf("RecipientCount", StringComparison.OrdinalIgnoreCase) < 0) continue;
                    Match m = RecipientCountOver.Match(condition);
                    long n;
                    if (m.Success && long.TryParse(m.Groups[1].Value, NumberStyles.None, CultureInfo.InvariantCulture, out n))
                    {
                        counts.Add(n);
                        countFromConditions = true;
                    }
                    else invalidCount = true;
                }
            }
            if (!targetFound) return Fail("TargetRuleNotInRecord");

            // An explicit RecipientCount field is not returned today, but is honoured if Microsoft adds one.
            bool countFromField = false;
            foreach (JsonElement holder in new[] { email, record })
            {
                JsonElement value;
                if (holder.ValueKind != JsonValueKind.Object || !holder.TryGetProperty("RecipientCount", out value)) continue;
                long n;
                if (TryParseCount(value, out n)) { counts.Add(n); countFromField = true; }
                else if (value.ValueKind != JsonValueKind.Null) invalidCount = true;
            }

            var ev = new DlpEvent
            {
                RecordIdentity = recordIdentity.Trim(),
                MessageId = messageId.Trim(),
                HappenedMs = happened,
                Sender = Text(email, "Sender") ?? Text(record, "Sender"),
                Subject = Text(email, "Subject") ?? Text(record, "Subject"),
                PolicyName = policyName,
                RuleName = ruleName
            };
            List<string> receivers = Strings(email, "Receivers") ?? Strings(record, "Receivers") ?? new List<string>();
            ev.RecipientListLength = receivers.Count;
            ev.Recipients = string.Join("; ", receivers);

            if (invalidCount) { ev.RecipientCount = null; ev.CountSource = "InvalidValue"; }
            else if (counts.Count > 1) { ev.RecipientCount = null; ev.CountSource = "ConflictingValues"; }
            else if (counts.Count == 1)
            {
                ev.RecipientCount = counts.Min;
                ev.CountSource = countFromField && countFromConditions ? "RecipientCountField+OtherConditions"
                               : countFromField ? "RecipientCountField" : "OtherConditions";
            }
            else { ev.RecipientCount = null; ev.CountSource = "NotReported"; }

            return new NormalizationResult { Event = ev };
        }

        static NormalizationResult Fail(string problem) { return new NormalizationResult { Problem = problem }; }

        static bool Same(string a, string b)
        {
            return a != null && b != null && string.Equals(a.Trim(), b.Trim(), StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>Returns a string property, or null if absent/empty/not a scalar.</summary>
        public static string Text(JsonElement holder, string name)
        {
            if (holder.ValueKind != JsonValueKind.Object) return null;
            JsonElement value;
            if (!holder.TryGetProperty(name, out value)) return null;
            switch (value.ValueKind)
            {
                case JsonValueKind.String:
                    string s = value.GetString();
                    return string.IsNullOrEmpty(s) ? null : s;
                case JsonValueKind.Number:
                case JsonValueKind.True:
                case JsonValueKind.False:
                    return value.GetRawText();
                default:
                    return null;
            }
        }

        /// <summary>Some fields may be returned either as an object or as a JSON string: accept both.</summary>
        static JsonElement ObjectOrParsedJson(JsonElement holder, string name)
        {
            if (holder.ValueKind != JsonValueKind.Object) return default(JsonElement);
            JsonElement value;
            if (!holder.TryGetProperty(name, out value)) return default(JsonElement);
            return ParseIfJsonString(value);
        }

        static JsonElement ParseIfJsonString(JsonElement value)
        {
            if (value.ValueKind == JsonValueKind.String)
            {
                string s = value.GetString();
                if (s != null)
                {
                    string t = s.TrimStart();
                    if (t.StartsWith("{") || t.StartsWith("["))
                    {
                        try { using (JsonDocument d = JsonDocument.Parse(s)) return d.RootElement.Clone(); }
                        catch (JsonException) { return default(JsonElement); }
                    }
                }
            }
            return value;
        }

        static IEnumerable<JsonElement> PolicyEntries(JsonElement record)
        {
            JsonElement info = ObjectOrParsedJson(record, "PolicyMatchInfo");
            if (info.ValueKind == JsonValueKind.Undefined || info.ValueKind == JsonValueKind.Null)
                info = ObjectOrParsedJson(record, "AEPolicyMatchInfo");
            if (info.ValueKind == JsonValueKind.Object) { yield return info; yield break; }
            if (info.ValueKind == JsonValueKind.Array)
            {
                foreach (JsonElement item in info.EnumerateArray())
                {
                    JsonElement e = ParseIfJsonString(item);
                    if (e.ValueKind == JsonValueKind.Object) yield return e;
                }
                yield break;
            }
            // Fallback: some schemas expose PolicyId/RuleId at the top level of the record.
            yield return record;
        }

        static IEnumerable<string> ConditionTexts(JsonElement value, int depth)
        {
            if (depth > 12) yield break;
            value = ParseIfJsonString(value);
            switch (value.ValueKind)
            {
                case JsonValueKind.String:
                    yield return value.GetString() ?? string.Empty;
                    break;
                case JsonValueKind.Array:
                    foreach (JsonElement item in value.EnumerateArray())
                        foreach (string s in ConditionTexts(item, depth + 1)) yield return s;
                    break;
                case JsonValueKind.Object:
                    JsonElement condition;
                    if (value.TryGetProperty("Condition", out condition))
                        foreach (string s in ConditionTexts(condition, depth + 1)) yield return s;
                    break;
            }
        }

        static List<string> Strings(JsonElement holder, string name)
        {
            if (holder.ValueKind != JsonValueKind.Object) return null;
            JsonElement value;
            if (!holder.TryGetProperty(name, out value)) return null;
            value = ParseIfJsonString(value);
            var list = new List<string>();
            if (value.ValueKind == JsonValueKind.Array)
            {
                foreach (JsonElement item in value.EnumerateArray())
                    if (item.ValueKind == JsonValueKind.String && !string.IsNullOrWhiteSpace(item.GetString()))
                        list.Add(item.GetString().Trim());
                return list;
            }
            if (value.ValueKind == JsonValueKind.String)
            {
                foreach (string part in (value.GetString() ?? string.Empty).Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries))
                    if (!string.IsNullOrWhiteSpace(part)) list.Add(part.Trim());
                return list;
            }
            return null;
        }

        static bool TryParseCount(JsonElement value, out long n)
        {
            n = 0;
            if (value.ValueKind == JsonValueKind.Number) return value.TryGetInt64(out n) && n >= 0;
            if (value.ValueKind == JsonValueKind.String)
            {
                string s = (value.GetString() ?? string.Empty).Trim();
                return s.Length > 0 && s.All(char.IsDigit) && long.TryParse(s, NumberStyles.None, CultureInfo.InvariantCulture, out n);
            }
            return false;
        }

        public static bool TryParseTime(string text, out long unixMs)
        {
            unixMs = 0;
            if (string.IsNullOrWhiteSpace(text)) return false;
            DateTimeOffset parsed;
            if (!DateTimeOffset.TryParse(text, CultureInfo.InvariantCulture,
                DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal, out parsed)) return false;
            unixMs = parsed.ToUnixTimeMilliseconds();
            return true;
        }
    }

    // -------------------------------------------------------------------------
    // 2. Interval arithmetic
    // -------------------------------------------------------------------------

    public struct TimeRange
    {
        public long Start;   // Unix ms, included
        public long End;     // Unix ms, excluded
        public TimeRange(long start, long end) { Start = start; End = end; }
        public long Length { get { return End - Start; } }
        public override string ToString()
        {
            return DateTimeOffset.FromUnixTimeMilliseconds(Start).ToString("o") + " -> " +
                   DateTimeOffset.FromUnixTimeMilliseconds(End).ToString("o");
        }
    }

    public static class Coverage
    {
        /// <summary>Merges overlapping or adjacent ranges.</summary>
        public static List<TimeRange> Merge(IEnumerable<TimeRange> ranges)
        {
            var result = new List<TimeRange>();
            foreach (TimeRange r in ranges.Where(x => x.End > x.Start).OrderBy(x => x.Start))
            {
                if (result.Count > 0 && r.Start <= result[result.Count - 1].End)
                {
                    TimeRange last = result[result.Count - 1];
                    if (r.End > last.End) result[result.Count - 1] = new TimeRange(last.Start, r.End);
                }
                else result.Add(r);
            }
            return result;
        }

        /// <summary>Returns the parts of [start, end) that are NOT covered by the given ranges.</summary>
        public static List<TimeRange> Gaps(IEnumerable<TimeRange> covered, long start, long end)
        {
            var gaps = new List<TimeRange>();
            long cursor = start;
            foreach (TimeRange r in Merge(covered))
            {
                if (r.End <= cursor) continue;
                if (r.Start >= end) break;
                if (r.Start > cursor) gaps.Add(new TimeRange(cursor, Math.Min(r.Start, end)));
                cursor = Math.Max(cursor, r.End);
                if (cursor >= end) break;
            }
            if (cursor < end) gaps.Add(new TimeRange(cursor, end));
            return gaps;
        }

        /// <summary>Splits ranges at local midnights, then into pieces of at most maxHours.</summary>
        public static List<TimeRange> SplitIntoSlices(IEnumerable<TimeRange> ranges, TimeZoneInfo zone, int maxHours)
        {
            var slices = new List<TimeRange>();
            long maxMs = Math.Max(1, maxHours) * 3600000L;
            foreach (TimeRange range in ranges)
            {
                long cursor = range.Start;
                while (cursor < range.End)
                {
                    long nextMidnight = NextLocalMidnight(cursor, zone);
                    long stop = Math.Min(range.End, Math.Min(nextMidnight, cursor + maxMs));
                    slices.Add(new TimeRange(cursor, stop));
                    cursor = stop;
                }
            }
            return slices;
        }

        public static long NextLocalMidnight(long unixMs, TimeZoneInfo zone)
        {
            DateTime local = TimeZoneInfo.ConvertTime(DateTimeOffset.FromUnixTimeMilliseconds(unixMs), zone).DateTime;
            return LocalToUnixMs(local.Date.AddDays(1), zone);
        }

        public static long LocalToUnixMs(DateTime local, TimeZoneInfo zone)
        {
            DateTime unspecified = DateTime.SpecifyKind(local, DateTimeKind.Unspecified);
            // Local midnight can be invalid (spring DST gap) in a few zones: move forward until valid.
            while (zone.IsInvalidTime(unspecified)) unspecified = unspecified.AddMinutes(30);
            TimeSpan offset = zone.IsAmbiguousTime(unspecified)
                ? zone.GetAmbiguousTimeOffsets(unspecified).Max()
                : zone.GetUtcOffset(unspecified);
            return new DateTimeOffset(unspecified, offset).ToUnixTimeMilliseconds();
        }
    }

    // -------------------------------------------------------------------------
    // 3. SQLite store
    // -------------------------------------------------------------------------

    public sealed class PageIngestResult
    {
        public int Records;          // records in ResultData
        public int NewEvents;        // inserted in this call
        public int Duplicates;       // already in the database (same RecordIdentity)
        public int Quarantined;      // rejected by the normalizer (kept in the quarantine table)
        public int Conflicts;        // same MessageID seen with a different sender/subject/recipients
        public int WithoutCount;     // events without a reliable recipient count
        public long MinHappenedMs;
        public long MaxHappenedMs;
        public string Sha256;        // hash of ResultData, used to detect a repeated page
        public string PolicyName;
        public string RuleName;
        public double ParseMs, NormalizeMs, WriteMs, CommitMs;   // time spent in each part (diagnostics)
        public Dictionary<string, int> QuarantineReasons = new Dictionary<string, int>();
    }

    /// <summary>
    /// Carries the ResultData text of one page from PowerShell to the engine.
    /// Why a wrapper and not a string parameter: PowerShell 7 sends the arguments of every
    /// .NET method call to the antimalware scan interface (AMSI). A 2 MB string argument
    /// costs about 0.5 s per call; a property assignment followed by a call with this
    /// small object does not.
    /// </summary>
    public sealed class PageData
    {
        public string ResultData { get; set; }
        public override string ToString() { return "ResultData (" + (ResultData == null ? 0 : ResultData.Length) + " characters)"; }
    }

    public sealed class IntervalRow
    {
        public long IntervalId, RunId, StartMs, EndMs, StartedMs;
        public long? CompletedMs;
        public string Status;
        public long Pages, Records, NewEvents;
        public long? ServiceTotal;
        public string Error;
    }

    public sealed class StoreStatistics
    {
        public long FileBytes, Events, Messages, Runs, Quarantined, CompletedIntervals;
        public long? FirstEventMs, LastEventMs, FirstCoveredMs, LastCoveredMs, LastSuccessfulCollectionMs;
        public int SchemaVersion;
        public string CreatedByVersion;
    }

    public sealed class DayStatistics
    {
        public DateTime LocalDate;
        public long StartMs, EndMs, Events, Messages, CoveredMs, FinalCoveredMs;
    }

    public sealed class PurgeResult { public long Events, Messages, Intervals, Quarantine; }

    public sealed class DlpStore : IDisposable
    {
        public const int SchemaVersion = 1;
        readonly SqliteConnection _connection;
        readonly bool _readOnly;
        SqliteCommand _findMessage, _insertMessage, _insertEvent, _insertQuarantine;
        SqliteTransaction _tx;   // current transaction; every command created by Command() joins it
        public string DatabasePath { get; private set; }
        public bool KeepQuarantinedRaw { get; set; }

        public DlpStore(string path, bool readOnly, string toolVersion)
        {
            DatabasePath = Path.GetFullPath(path);
            _readOnly = readOnly;
            if (!readOnly) Directory.CreateDirectory(Path.GetDirectoryName(DatabasePath));
            var builder = new SqliteConnectionStringBuilder
            {
                DataSource = DatabasePath,
                Mode = readOnly ? SqliteOpenMode.ReadOnly : SqliteOpenMode.ReadWriteCreate,
                Pooling = false          // release the file as soon as the store is disposed
            };
            _connection = new SqliteConnection(builder.ToString());
            _connection.Open();
            Execute("PRAGMA busy_timeout=60000;");
            Execute("PRAGMA cache_size=-65536;");   // 64 MB page cache
            if (readOnly)
            {
                long version = Scalar<long>("PRAGMA user_version;");
                if (version != SchemaVersion)
                    throw new InvalidOperationException("The database schema version is " + version + "; this tool expects " + SchemaVersion + ".");
            }
            else Migrate(toolVersion);
            KeepQuarantinedRaw = true;
        }

        public void Dispose()
        {
            foreach (SqliteCommand c in new[] { _findMessage, _insertMessage, _insertEvent, _insertQuarantine })
                if (c != null) c.Dispose();
            if (!_readOnly)
            {
                try { Execute("PRAGMA optimize;"); Execute("PRAGMA wal_checkpoint(TRUNCATE);"); } catch (SqliteException) { }
            }
            _connection.Dispose();
        }

        // ---- Schema ---------------------------------------------------------

        void Migrate(string toolVersion)
        {
            long version = Scalar<long>("PRAGMA user_version;");
            if (version > SchemaVersion)
                throw new InvalidOperationException("The database was created by a newer version of the tool (schema " + version + ").");
            if (version == 0)
            {
                // auto_vacuum must be chosen before the first table is created.
                Execute("PRAGMA auto_vacuum=INCREMENTAL;");
                Execute("PRAGMA journal_mode=WAL;");
                _tx = _connection.BeginTransaction();
                try
                {
                    Execute(@"
CREATE TABLE metadata (
    key   TEXT PRIMARY KEY,
    value TEXT NOT NULL
);
-- One row per execution of the tool.
CREATE TABLE run (
    run_id       INTEGER PRIMARY KEY,
    started_ms   INTEGER NOT NULL,
    finished_ms  INTEGER,
    mode         TEXT    NOT NULL,
    status       TEXT    NOT NULL,
    account      TEXT,
    host         TEXT,
    tool_version TEXT,
    details      TEXT
);
-- DLP policy/rule pairs collected (today one; several are supported by the schema).
CREATE TABLE target (
    target_id   INTEGER PRIMARY KEY,
    policy_id   TEXT NOT NULL COLLATE NOCASE,
    rule_id     TEXT NOT NULL COLLATE NOCASE,
    policy_name TEXT,
    rule_name   TEXT,
    UNIQUE (policy_id, rule_id)
);
-- Time ranges requested from Activity Explorer. status = Running | Completed | Failed.
-- A Completed range whose paging finished is 'covered'.
CREATE TABLE collection_interval (
    interval_id   INTEGER PRIMARY KEY,
    run_id        INTEGER NOT NULL REFERENCES run(run_id),
    target_id     INTEGER NOT NULL REFERENCES target(target_id),
    start_ms      INTEGER NOT NULL,
    end_ms        INTEGER NOT NULL,
    status        TEXT    NOT NULL,
    started_ms    INTEGER NOT NULL,
    completed_ms  INTEGER,
    pages         INTEGER NOT NULL DEFAULT 0,
    records       INTEGER NOT NULL DEFAULT 0,
    new_events    INTEGER NOT NULL DEFAULT 0,
    service_total INTEGER,
    error         TEXT
);
CREATE INDEX ix_interval_target ON collection_interval (target_id, status, start_ms);
-- One row per unique MessageID.
CREATE TABLE message (
    message_key          INTEGER PRIMARY KEY,
    message_id           TEXT NOT NULL UNIQUE,
    sender               TEXT,
    subject              TEXT,
    recipients           TEXT,
    recipient_list_count INTEGER
);
-- One row per Activity Explorer event (RecordIdentity): re-collecting is idempotent.
CREATE TABLE dlp_event (
    record_identity TEXT    PRIMARY KEY,
    target_id       INTEGER NOT NULL REFERENCES target(target_id),
    message_key     INTEGER NOT NULL REFERENCES message(message_key),
    happened_ms     INTEGER NOT NULL,
    recipient_count INTEGER,
    count_source    TEXT    NOT NULL,
    interval_id     INTEGER,
    collected_ms    INTEGER NOT NULL
) WITHOUT ROWID;
CREATE INDEX ix_event_target_time ON dlp_event (target_id, happened_ms, message_key);
CREATE INDEX ix_event_message     ON dlp_event (message_key);
-- Records that could not be normalized. Never used in reports; kept for troubleshooting.
CREATE TABLE quarantine (
    quarantine_id   INTEGER PRIMARY KEY,
    interval_id     INTEGER,
    received_ms     INTEGER NOT NULL,
    record_identity TEXT,
    reason          TEXT NOT NULL,
    raw_json        TEXT
);
CREATE INDEX ix_quarantine_time ON quarantine (received_ms);
");
                    SetMetadata("created_utc", DateTimeOffset.UtcNow.ToString("o"));
                    SetMetadata("created_by_version", toolVersion ?? "");
                    Execute("PRAGMA user_version=" + SchemaVersion + ";");
                    _tx.Commit();
                }
                finally { _tx.Dispose(); _tx = null; }
            }
            Execute("PRAGMA journal_mode=WAL;");
            Execute("PRAGMA synchronous=NORMAL;");
            Execute("PRAGMA foreign_keys=ON;");
            SetMetadata("last_opened_by_version", toolVersion ?? "");
        }

        void SetMetadata(string key, string value)
        {
            using (SqliteCommand c = Command("INSERT INTO metadata(key, value) VALUES($k, $v) ON CONFLICT(key) DO UPDATE SET value = excluded.value;"))
            {
                c.Parameters.AddWithValue("$k", key);
                c.Parameters.AddWithValue("$v", value);
                c.ExecuteNonQuery();
            }
        }

        public string GetMetadata(string key)
        {
            using (SqliteCommand c = Command("SELECT value FROM metadata WHERE key = $k;"))
            {
                c.Parameters.AddWithValue("$k", key);
                object o = c.ExecuteScalar();
                return o == null || o is DBNull ? null : (string)o;
            }
        }

        // ---- Runs, targets, intervals --------------------------------------

        public long StartRun(string mode, string account, string host, string toolVersion, string details)
        {
            using (SqliteCommand c = Command("INSERT INTO run(started_ms, mode, status, account, host, tool_version, details) VALUES($s, $m, 'Running', $a, $h, $v, $d) RETURNING run_id;"))
            {
                c.Parameters.AddWithValue("$s", NowMs());
                c.Parameters.AddWithValue("$m", mode);
                c.Parameters.AddWithValue("$a", (object)account ?? DBNull.Value);
                c.Parameters.AddWithValue("$h", (object)host ?? DBNull.Value);
                c.Parameters.AddWithValue("$v", (object)toolVersion ?? DBNull.Value);
                c.Parameters.AddWithValue("$d", (object)details ?? DBNull.Value);
                return (long)c.ExecuteScalar();
            }
        }

        public void UpdateRunAccount(long runId, string account)
        {
            using (SqliteCommand c = Command("UPDATE run SET account = $a WHERE run_id = $r;"))
            {
                c.Parameters.AddWithValue("$a", (object)account ?? DBNull.Value);
                c.Parameters.AddWithValue("$r", runId);
                c.ExecuteNonQuery();
            }
        }

        public void FinishRun(long runId, string status, string details)
        {
            using (SqliteCommand c = Command("UPDATE run SET finished_ms = $f, status = $s, details = COALESCE($d, details) WHERE run_id = $r;"))
            {
                c.Parameters.AddWithValue("$f", NowMs());
                c.Parameters.AddWithValue("$s", status);
                c.Parameters.AddWithValue("$d", (object)details ?? DBNull.Value);
                c.Parameters.AddWithValue("$r", runId);
                c.ExecuteNonQuery();
            }
        }

        /// <summary>Intervals left 'Running' by an interrupted execution are marked as failed.</summary>
        public int CloseAbandonedWork()
        {
            int n;
            using (SqliteCommand c = Command("UPDATE collection_interval SET status = 'Failed', error = 'Interrupted (the previous execution stopped before the end of this range)' WHERE status = 'Running';"))
                n = c.ExecuteNonQuery();
            using (SqliteCommand c = Command("UPDATE run SET status = 'Interrupted' WHERE status = 'Running';"))
                c.ExecuteNonQuery();
            return n;
        }

        public long GetOrCreateTarget(string policyId, string ruleId, string policyName, string ruleName)
        {
            using (SqliteCommand c = Command(@"INSERT INTO target(policy_id, rule_id, policy_name, rule_name) VALUES($p, $r, $pn, $rn)
ON CONFLICT(policy_id, rule_id) DO UPDATE SET policy_name = COALESCE(excluded.policy_name, target.policy_name), rule_name = COALESCE(excluded.rule_name, target.rule_name)
RETURNING target_id;"))
            {
                c.Parameters.AddWithValue("$p", policyId);
                c.Parameters.AddWithValue("$r", ruleId);
                c.Parameters.AddWithValue("$pn", (object)policyName ?? DBNull.Value);
                c.Parameters.AddWithValue("$rn", (object)ruleName ?? DBNull.Value);
                return (long)c.ExecuteScalar();
            }
        }

        public long FindTarget(string policyId, string ruleId)
        {
            using (SqliteCommand c = Command("SELECT target_id FROM target WHERE policy_id = $p AND rule_id = $r;"))
            {
                c.Parameters.AddWithValue("$p", policyId);
                c.Parameters.AddWithValue("$r", ruleId);
                object o = c.ExecuteScalar();
                return o == null || o is DBNull ? 0 : (long)o;
            }
        }

        public long BeginInterval(long runId, long targetId, long startMs, long endMs)
        {
            using (SqliteCommand c = Command("INSERT INTO collection_interval(run_id, target_id, start_ms, end_ms, status, started_ms) VALUES($run, $t, $s, $e, 'Running', $now) RETURNING interval_id;"))
            {
                c.Parameters.AddWithValue("$run", runId);
                c.Parameters.AddWithValue("$t", targetId);
                c.Parameters.AddWithValue("$s", startMs);
                c.Parameters.AddWithValue("$e", endMs);
                c.Parameters.AddWithValue("$now", NowMs());
                return (long)c.ExecuteScalar();
            }
        }

        public void EndInterval(long intervalId, bool completed, long pages, long records, long newEvents, long? serviceTotal, string error)
        {
            using (SqliteCommand c = Command(@"UPDATE collection_interval SET status = $st, completed_ms = $c, pages = $p, records = $r, new_events = $n, service_total = $t, error = $e WHERE interval_id = $id;"))
            {
                c.Parameters.AddWithValue("$st", completed ? "Completed" : "Failed");
                c.Parameters.AddWithValue("$c", completed ? (object)NowMs() : DBNull.Value);
                c.Parameters.AddWithValue("$p", pages);
                c.Parameters.AddWithValue("$r", records);
                c.Parameters.AddWithValue("$n", newEvents);
                c.Parameters.AddWithValue("$t", serviceTotal.HasValue ? (object)serviceTotal.Value : DBNull.Value);
                c.Parameters.AddWithValue("$e", (object)error ?? DBNull.Value);
                c.Parameters.AddWithValue("$id", intervalId);
                c.ExecuteNonQuery();
            }
        }

        /// <summary>
        /// Completed ranges for a target. When settlingMs is greater than zero, the part of each
        /// range that was less than settlingMs old when it was collected is left out ("settled"
        /// part only): Activity Explorer can receive late events, so those recent hours are
        /// collected again at the next run.
        /// </summary>
        public List<TimeRange> GetCoveredRanges(long targetId, long settlingMs)
        {
            var list = new List<TimeRange>();
            using (SqliteCommand c = Command("SELECT start_ms, end_ms, completed_ms FROM collection_interval WHERE target_id = $t AND status = 'Completed' ORDER BY start_ms;"))
            {
                c.Parameters.AddWithValue("$t", targetId);
                using (SqliteDataReader r = c.ExecuteReader())
                    while (r.Read())
                    {
                        long start = r.GetInt64(0), end = r.GetInt64(1), completed = r.GetInt64(2);
                        if (settlingMs > 0) end = Math.Min(end, completed - settlingMs);
                        if (end > start) list.Add(new TimeRange(start, end));
                    }
            }
            return Coverage.Merge(list);
        }

        public List<IntervalRow> GetIntervals(long runId)
        {
            var list = new List<IntervalRow>();
            using (SqliteCommand c = Command("SELECT interval_id, run_id, start_ms, end_ms, started_ms, completed_ms, status, pages, records, new_events, service_total, error FROM collection_interval WHERE run_id = $r ORDER BY interval_id;"))
            {
                c.Parameters.AddWithValue("$r", runId);
                using (SqliteDataReader r = c.ExecuteReader())
                    while (r.Read())
                        list.Add(new IntervalRow
                        {
                            IntervalId = r.GetInt64(0), RunId = r.GetInt64(1), StartMs = r.GetInt64(2), EndMs = r.GetInt64(3),
                            StartedMs = r.GetInt64(4), CompletedMs = r.IsDBNull(5) ? (long?)null : r.GetInt64(5), Status = r.GetString(6),
                            Pages = r.GetInt64(7), Records = r.GetInt64(8), NewEvents = r.GetInt64(9),
                            ServiceTotal = r.IsDBNull(10) ? (long?)null : r.GetInt64(10), Error = r.IsDBNull(11) ? null : r.GetString(11)
                        });
            }
            return list;
        }

        // ---- Ingestion ------------------------------------------------------

        /// <summary>
        /// Parses one Activity Explorer page (the ResultData JSON array), normalizes every record
        /// and stores it in a single transaction. Safe to call again with the same data.
        /// </summary>
        public PageIngestResult IngestPage(long intervalId, long targetId, string policyId, string ruleId, PageData page)
        {
            if (page == null || page.ResultData == null) throw new ArgumentNullException("page");
            string resultDataJson = page.ResultData;
            var result = new PageIngestResult { MinHappenedMs = long.MaxValue, MaxHappenedMs = long.MinValue };
            var clock = System.Diagnostics.Stopwatch.StartNew();
            using (SHA256 sha = SHA256.Create())
                result.Sha256 = Convert.ToHexString(sha.ComputeHash(Encoding.UTF8.GetBytes(resultDataJson)));

            JsonDocument parsed;
            try { parsed = JsonDocument.Parse(resultDataJson); }
            catch (JsonException ex) { throw new FormatException("MalformedPage: ResultData is not valid JSON (" + ex.Message + ")."); }
            using (JsonDocument document = parsed)
            {
                if (document.RootElement.ValueKind != JsonValueKind.Array)
                    throw new FormatException("MalformedPage: ResultData is not a JSON array.");
                result.ParseMs = clock.Elapsed.TotalMilliseconds;
                long now = NowMs();
                PrepareIngestCommands();
                _tx = _connection.BeginTransaction();
                try
                {
                    foreach (SqliteCommand c in new[] { _findMessage, _insertMessage, _insertEvent, _insertQuarantine }) c.Transaction = _tx;
                    foreach (JsonElement record in document.RootElement.EnumerateArray())
                    {
                        result.Records++;
                        NormalizationResult n;
                        long before = clock.ElapsedTicks;
                        try { n = EventNormalizer.Normalize(record, policyId, ruleId); }
                        catch (Exception ex) { n = new NormalizationResult { Problem = "NormalizerException:" + ex.GetType().Name }; }
                        result.NormalizeMs += (clock.ElapsedTicks - before) * 1000.0 / System.Diagnostics.Stopwatch.Frequency;
                        if (n.Problem != null)
                        {
                            Quarantine(intervalId, now, record, n.Problem);
                            result.Quarantined++;
                            string reason = n.Problem.Split(':')[0];
                            int count;
                            result.QuarantineReasons[reason] = result.QuarantineReasons.TryGetValue(reason, out count) ? count + 1 : 1;
                            continue;
                        }
                        DlpEvent e = n.Event;
                        if (result.PolicyName == null) result.PolicyName = e.PolicyName;
                        if (result.RuleName == null) result.RuleName = e.RuleName;
                        if (!e.RecipientCount.HasValue) result.WithoutCount++;
                        result.MinHappenedMs = Math.Min(result.MinHappenedMs, e.HappenedMs);
                        result.MaxHappenedMs = Math.Max(result.MaxHappenedMs, e.HappenedMs);

                        long messageKey;
                        bool conflict;
                        UpsertMessage(e, out messageKey, out conflict);
                        if (conflict) result.Conflicts++;

                        _insertEvent.Parameters["$id"].Value = e.RecordIdentity;
                        _insertEvent.Parameters["$t"].Value = targetId;
                        _insertEvent.Parameters["$m"].Value = messageKey;
                        _insertEvent.Parameters["$h"].Value = e.HappenedMs;
                        _insertEvent.Parameters["$c"].Value = e.RecipientCount.HasValue ? (object)e.RecipientCount.Value : DBNull.Value;
                        _insertEvent.Parameters["$cs"].Value = e.CountSource;
                        _insertEvent.Parameters["$i"].Value = intervalId;
                        _insertEvent.Parameters["$now"].Value = now;
                        if (_insertEvent.ExecuteNonQuery() == 1) result.NewEvents++;
                        else result.Duplicates++;
                    }
                    double beforeCommit = clock.Elapsed.TotalMilliseconds;
                    _tx.Commit();
                    result.CommitMs = clock.Elapsed.TotalMilliseconds - beforeCommit;
                    result.WriteMs = beforeCommit - result.ParseMs - result.NormalizeMs;
                }
                finally
                {
                    _tx.Dispose(); _tx = null;
                    foreach (SqliteCommand c in new[] { _findMessage, _insertMessage, _insertEvent, _insertQuarantine }) c.Transaction = null;
                }
            }
            if (result.MinHappenedMs == long.MaxValue) { result.MinHappenedMs = 0; result.MaxHappenedMs = 0; }
            return result;
        }

        void PrepareIngestCommands()
        {
            if (_insertEvent != null) return;
            _findMessage = Command("SELECT message_key, sender, subject, recipients FROM message WHERE message_id = $mid;");
            _findMessage.Parameters.Add("$mid", SqliteType.Text);
            _findMessage.Prepare();
            _insertMessage = Command("INSERT INTO message(message_id, sender, subject, recipients, recipient_list_count) VALUES($mid, $s, $j, $r, $n) RETURNING message_key;");
            foreach (string p in new[] { "$mid", "$s", "$j", "$r" }) _insertMessage.Parameters.Add(p, SqliteType.Text);
            _insertMessage.Parameters.Add("$n", SqliteType.Integer);
            _insertMessage.Prepare();
            _insertEvent = Command(@"INSERT INTO dlp_event(record_identity, target_id, message_key, happened_ms, recipient_count, count_source, interval_id, collected_ms)
VALUES($id, $t, $m, $h, $c, $cs, $i, $now) ON CONFLICT(record_identity) DO NOTHING;");
            _insertEvent.Parameters.Add("$id", SqliteType.Text);
            foreach (string p in new[] { "$t", "$m", "$h", "$c", "$i", "$now" }) _insertEvent.Parameters.Add(p, SqliteType.Integer);
            _insertEvent.Parameters.Add("$cs", SqliteType.Text);
            _insertEvent.Prepare();
            _insertQuarantine = Command("INSERT INTO quarantine(interval_id, received_ms, record_identity, reason, raw_json) VALUES($i, $now, $id, $reason, $raw);");
            _insertQuarantine.Parameters.Add("$i", SqliteType.Integer);
            _insertQuarantine.Parameters.Add("$now", SqliteType.Integer);
            foreach (string p in new[] { "$id", "$reason", "$raw" }) _insertQuarantine.Parameters.Add(p, SqliteType.Text);
            _insertQuarantine.Prepare();
        }

        void UpsertMessage(DlpEvent e, out long messageKey, out bool conflict)
        {
            conflict = false;
            _findMessage.Parameters["$mid"].Value = e.MessageId;
            using (SqliteDataReader r = _findMessage.ExecuteReader())
            {
                if (r.Read())
                {
                    messageKey = r.GetInt64(0);
                    // The first stored values are kept; a different value is only counted.
                    conflict = !SameText(r.IsDBNull(1) ? null : r.GetString(1), e.Sender, true)
                            || !SameText(r.IsDBNull(2) ? null : r.GetString(2), e.Subject, false)
                            || !SameRecipients(r.IsDBNull(3) ? null : r.GetString(3), e.Recipients);
                    return;
                }
            }
            _insertMessage.Parameters["$mid"].Value = e.MessageId;
            _insertMessage.Parameters["$s"].Value = (object)e.Sender ?? DBNull.Value;
            _insertMessage.Parameters["$j"].Value = (object)e.Subject ?? DBNull.Value;
            _insertMessage.Parameters["$r"].Value = (object)e.Recipients ?? DBNull.Value;
            _insertMessage.Parameters["$n"].Value = e.RecipientListLength;
            messageKey = (long)_insertMessage.ExecuteScalar();
        }

        static bool SameText(string a, string b, bool ignoreCase)
        {
            return string.Equals(a ?? "", b ?? "", ignoreCase ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal);
        }

        static bool SameRecipients(string a, string b)
        {
            if (SameText(a, b, true)) return true;
            Func<string, string> canonical = s => string.Join(";", (s ?? "").Split(new[] { "; " }, StringSplitOptions.RemoveEmptyEntries)
                .Select(x => x.Trim().ToLowerInvariant()).OrderBy(x => x, StringComparer.Ordinal));
            return canonical(a) == canonical(b);
        }

        void Quarantine(long intervalId, long now, JsonElement record, string reason)
        {
            _insertQuarantine.Parameters["$i"].Value = intervalId;
            _insertQuarantine.Parameters["$now"].Value = now;
            _insertQuarantine.Parameters["$id"].Value = (object)EventNormalizer.Text(record, "RecordIdentity") ?? DBNull.Value;
            _insertQuarantine.Parameters["$reason"].Value = reason;
            _insertQuarantine.Parameters["$raw"].Value = KeepQuarantinedRaw ? (object)record.GetRawText() : DBNull.Value;
            _insertQuarantine.ExecuteNonQuery();
        }

        // ---- Statistics -----------------------------------------------------

        public StoreStatistics GetStatistics(long targetId)
        {
            var s = new StoreStatistics();
            s.FileBytes = FileSize(DatabasePath) + FileSize(DatabasePath + "-wal");
            s.SchemaVersion = (int)Scalar<long>("PRAGMA user_version;");
            s.CreatedByVersion = GetMetadata("created_by_version");
            s.Messages = Scalar<long>("SELECT COUNT(*) FROM message;");
            s.Runs = Scalar<long>("SELECT COUNT(*) FROM run;");
            s.Quarantined = Scalar<long>("SELECT COUNT(*) FROM quarantine;");
            using (SqliteCommand c = Command("SELECT COUNT(*), MIN(happened_ms), MAX(happened_ms) FROM dlp_event WHERE target_id = $t;"))
            {
                c.Parameters.AddWithValue("$t", targetId);
                using (SqliteDataReader r = c.ExecuteReader())
                    if (r.Read())
                    {
                        s.Events = r.GetInt64(0);
                        s.FirstEventMs = r.IsDBNull(1) ? (long?)null : r.GetInt64(1);
                        s.LastEventMs = r.IsDBNull(2) ? (long?)null : r.GetInt64(2);
                    }
            }
            using (SqliteCommand c = Command("SELECT COUNT(*), MIN(start_ms), MAX(end_ms), MAX(completed_ms) FROM collection_interval WHERE target_id = $t AND status = 'Completed';"))
            {
                c.Parameters.AddWithValue("$t", targetId);
                using (SqliteDataReader r = c.ExecuteReader())
                    if (r.Read())
                    {
                        s.CompletedIntervals = r.GetInt64(0);
                        s.FirstCoveredMs = r.IsDBNull(1) ? (long?)null : r.GetInt64(1);
                        s.LastCoveredMs = r.IsDBNull(2) ? (long?)null : r.GetInt64(2);
                        s.LastSuccessfulCollectionMs = r.IsDBNull(3) ? (long?)null : r.GetInt64(3);
                    }
            }
            return s;
        }

        public List<DayStatistics> GetDailyStatistics(long targetId, DateTime firstLocalDate, int days, TimeZoneInfo zone, long settlingMs)
        {
            List<TimeRange> covered = GetCoveredRanges(targetId, 0);
            List<TimeRange> final = GetCoveredRanges(targetId, settlingMs);
            var list = new List<DayStatistics>();
            using (SqliteCommand c = Command("SELECT COUNT(*), COUNT(DISTINCT message_key) FROM dlp_event WHERE target_id = $t AND happened_ms >= $s AND happened_ms < $e;"))
            {
                c.Parameters.Add("$t", SqliteType.Integer).Value = targetId;
                c.Parameters.Add("$s", SqliteType.Integer);
                c.Parameters.Add("$e", SqliteType.Integer);
                for (int i = 0; i < days; i++)
                {
                    DateTime day = firstLocalDate.Date.AddDays(i);
                    long start = Coverage.LocalToUnixMs(day, zone), end = Coverage.LocalToUnixMs(day.AddDays(1), zone);
                    c.Parameters["$s"].Value = start;
                    c.Parameters["$e"].Value = end;
                    var d = new DayStatistics { LocalDate = day, StartMs = start, EndMs = end };
                    using (SqliteDataReader r = c.ExecuteReader())
                        if (r.Read()) { d.Events = r.GetInt64(0); d.Messages = r.GetInt64(1); }
                    d.CoveredMs = (end - start) - Coverage.Gaps(covered, start, end).Sum(g => g.Length);
                    d.FinalCoveredMs = (end - start) - Coverage.Gaps(final, start, end).Sum(g => g.Length);
                    list.Add(d);
                }
            }
            return list;
        }

        // ---- Retention ------------------------------------------------------

        /// <summary>Deletes everything older than cutoffMs and trims coverage accordingly.</summary>
        public PurgeResult PurgeBefore(long cutoffMs)
        {
            var p = new PurgeResult();
            _tx = _connection.BeginTransaction();
            try
            {
                p.Events = NonQuery("DELETE FROM dlp_event WHERE happened_ms < $c;", cutoffMs);
                p.Messages = NonQuery("DELETE FROM message WHERE NOT EXISTS (SELECT 1 FROM dlp_event e WHERE e.message_key = message.message_key);", null);
                p.Intervals = NonQuery("DELETE FROM collection_interval WHERE end_ms <= $c;", cutoffMs);
                NonQuery("UPDATE collection_interval SET start_ms = $c WHERE start_ms < $c AND end_ms > $c;", cutoffMs);
                p.Quarantine = NonQuery("DELETE FROM quarantine WHERE received_ms < $c;", cutoffMs);
                _tx.Commit();
            }
            finally { _tx.Dispose(); _tx = null; }
            if (p.Events + p.Messages + p.Quarantine > 0) Execute("PRAGMA incremental_vacuum;");
            return p;
        }

        long NonQuery(string sql, long? cutoff)
        {
            using (SqliteCommand c = Command(sql))
            {
                if (cutoff.HasValue) c.Parameters.AddWithValue("$c", cutoff.Value);
                return c.ExecuteNonQuery();
            }
        }

        // ---- Report queries -------------------------------------------------

        const string MessageAggregate = @"
SELECT message_key, MIN(happened_ms) AS first_ms, COUNT(*) AS events,
       MAX(recipient_count) AS recipient_count, COUNT(DISTINCT recipient_count) AS distinct_counts
FROM dlp_event
WHERE target_id = $t AND happened_ms >= $s AND happened_ms < $e
GROUP BY message_key";

        /// <summary>
        /// Builds and writes the report inside one read transaction, so the row plan
        /// (first pass) and the written rows (second pass) see exactly the same data.
        /// </summary>
        public ReportResult WriteReport(long targetId, ReportRequest request)
        {
            _tx = _connection.BeginTransaction(true);
            try
            {
                var firstTimes = new List<long>();
                using (SqliteCommand c = Command("SELECT first_ms FROM (" + MessageAggregate + ") ORDER BY first_ms, message_key;"))
                {
                    AddRange(c, targetId, request.StartMs, request.EndMs);
                    using (SqliteDataReader r = c.ExecuteReader())
                        while (r.Read()) firstTimes.Add(r.GetInt64(0));
                }
                List<ReportPart> plan = ReportPlanner.Plan(firstTimes, request);
                using (SqliteCommand c = Command(@"SELECT m.message_id, a.first_ms, a.events, a.recipient_count, a.distinct_counts, m.sender, m.subject, m.recipients
FROM (" + MessageAggregate + @") a JOIN message m ON m.message_key = a.message_key
ORDER BY a.first_ms, a.message_key;"))
                {
                    AddRange(c, targetId, request.StartMs, request.EndMs);
                    using (SqliteDataReader r = c.ExecuteReader())
                    {
                        IEnumerable<ReportRow> rows = ReadRows(r);
                        ReportResult result = ReportWriter.Write(rows, plan, request);
                        _tx.Commit();
                        return result;
                    }
                }
            }
            finally { _tx.Dispose(); _tx = null; }
        }

        static IEnumerable<ReportRow> ReadRows(SqliteDataReader r)
        {
            while (r.Read())
            {
                yield return new ReportRow
                {
                    MessageId = r.GetString(0),
                    FirstMs = r.GetInt64(1),
                    Events = r.GetInt64(2),
                    RecipientCount = r.IsDBNull(3) ? (long?)null : r.GetInt64(3),
                    CountConflict = r.GetInt64(4) > 1,
                    Sender = r.IsDBNull(5) ? "" : r.GetString(5),
                    Subject = r.IsDBNull(6) ? "" : r.GetString(6),
                    Recipients = r.IsDBNull(7) ? "" : r.GetString(7)
                };
            }
        }

        static void AddRange(SqliteCommand c, long targetId, long start, long end)
        {
            c.Parameters.AddWithValue("$t", targetId);
            c.Parameters.AddWithValue("$s", start);
            c.Parameters.AddWithValue("$e", end);
        }

        // ---- Helpers --------------------------------------------------------

        SqliteCommand Command(string sql)
        {
            SqliteCommand c = _connection.CreateCommand();
            c.CommandText = sql;
            c.CommandTimeout = 0;
            c.Transaction = _tx;
            return c;
        }

        void Execute(string sql)
        {
            using (SqliteCommand c = Command(sql)) c.ExecuteNonQuery();
        }

        T Scalar<T>(string sql)
        {
            using (SqliteCommand c = Command(sql))
            {
                object o = c.ExecuteScalar();
                if (o == null || o is DBNull) return default(T);
                return (T)Convert.ChangeType(o, typeof(T), CultureInfo.InvariantCulture);
            }
        }

        static long FileSize(string path) { return File.Exists(path) ? new FileInfo(path).Length : 0; }
        public static long NowMs() { return DateTimeOffset.UtcNow.ToUnixTimeMilliseconds(); }
    }

    // -------------------------------------------------------------------------
    // 4. Report planning and writing
    // -------------------------------------------------------------------------

    public sealed class ReportRow
    {
        public string MessageId, Sender, Subject, Recipients;
        public long FirstMs, Events;
        public long? RecipientCount;
        public bool CountConflict;
    }

    public sealed class ReportRequest
    {
        public long StartMs, EndMs;                 // report period [start, end)
        public TimeZoneInfo Zone;
        public string TimeZoneLabel;                // shown in column headers, e.g. "Europe/Paris"
        public string SplitBy = "Week";             // Rows | Day | Week
        public int MaxRowsPerFile = 500000;
        public bool IncludeRecipientDetails = true;
        public bool WriteCsv = true, WriteHtml = true;
        public string CsvDelimiter = ";";
        public string OutputDirectory;
        public string FilePrefix = "PurviewDLP";
        public string HtmlTemplatePath;
        public string Title = "DLP messages with more than 25 recipients";
        public string PolicyName, RuleName, ToolVersion, RangeName, CoverageNote;
        public double CoveragePercent = 100;
        public int HtmlChunkRows = 20000;
    }

    public sealed class ReportPart
    {
        public int Index;                 // 1-based position in the whole report
        public long FirstRow, RowCount;   // position in the ordered message list
        public long StartMs, EndMs;       // period shown for this file
        public long GroupStartMs, GroupEndMs;   // day or week (or whole period) the file belongs to; used in the file name
        public string GroupLabel;         // e.g. "2026-09-01 to 2026-09-07"
        public int PartInGroup = 1, PartsInGroup = 1;
        public string BaseName;           // file name without extension
    }

    public sealed class ReportFile
    {
        public string Path, Kind;
        public long Rows, Bytes;
        public int PartIndex;
        public string Label;
    }

    public sealed class ReportResult
    {
        public long Messages, Events, Senders, CountConflicts, MessagesWithoutCount, RecipientTotal;
        public List<ReportFile> Files = new List<ReportFile>();
        public List<ReportPart> Parts = new List<ReportPart>();
    }

    public static class ReportPlanner
    {
        /// <summary>
        /// Decides how many files to write. There is no split while the report fits in
        /// MaxRowsPerFile rows. Above that, rows are grouped by local day or local week
        /// (Monday to Sunday), or simply cut every MaxRowsPerFile rows. A day or week that
        /// is still too large is cut into balanced numbered parts.
        /// </summary>
        public static List<ReportPart> Plan(IList<long> orderedFirstMs, ReportRequest q)
        {
            var parts = new List<ReportPart>();
            long total = orderedFirstMs.Count;
            int max = Math.Max(1, q.MaxRowsPerFile);
            if (total <= max)
            {
                parts.Add(new ReportPart { FirstRow = 0, RowCount = total, StartMs = q.StartMs, EndMs = q.EndMs, GroupStartMs = q.StartMs, GroupEndMs = q.EndMs, GroupLabel = PeriodLabel(q.StartMs, q.EndMs, q.Zone) });
            }
            else if (string.Equals(q.SplitBy, "Rows", StringComparison.OrdinalIgnoreCase))
            {
                int count = (int)((total + max - 1) / max);
                for (int i = 0; i < count; i++)
                {
                    long first = (long)i * max, rows = Math.Min(max, total - first);
                    long start = i == 0 ? q.StartMs : orderedFirstMs[(int)first];
                    long end = i == count - 1 ? q.EndMs : orderedFirstMs[(int)(first + rows)];
                    parts.Add(new ReportPart { FirstRow = first, RowCount = rows, StartMs = start, EndMs = end, GroupStartMs = q.StartMs, GroupEndMs = q.EndMs, GroupLabel = PeriodLabel(start, end, q.Zone), PartInGroup = i + 1, PartsInGroup = count });
                }
            }
            else
            {
                bool week = string.Equals(q.SplitBy, "Week", StringComparison.OrdinalIgnoreCase);
                long row = 0;
                while (row < total)
                {
                    DateTime groupStartLocal = GroupStart(orderedFirstMs[(int)row], q.Zone, week);
                    long groupEndMs = Coverage.LocalToUnixMs(groupStartLocal.AddDays(week ? 7 : 1), q.Zone);
                    long groupStartMs = Math.Max(q.StartMs, Coverage.LocalToUnixMs(groupStartLocal, q.Zone));
                    long clippedEnd = Math.Min(q.EndMs, groupEndMs);
                    long first = row;
                    while (row < total && orderedFirstMs[(int)row] < groupEndMs) row++;
                    long rows = row - first;
                    int pieces = (int)((rows + max - 1) / max);
                    long size = (rows + pieces - 1) / pieces;
                    string label = PeriodLabel(groupStartMs, clippedEnd, q.Zone);
                    for (int i = 0; i < pieces; i++)
                    {
                        long pf = first + i * size, pr = Math.Min(size, rows - i * size);
                        parts.Add(new ReportPart
                        {
                            FirstRow = pf, RowCount = pr,
                            StartMs = i == 0 ? groupStartMs : orderedFirstMs[(int)pf],
                            EndMs = i == pieces - 1 ? clippedEnd : orderedFirstMs[(int)(pf + pr)],
                            GroupStartMs = groupStartMs, GroupEndMs = clippedEnd,
                            GroupLabel = label, PartInGroup = i + 1, PartsInGroup = pieces
                        });
                    }
                }
            }
            for (int i = 0; i < parts.Count; i++)
            {
                ReportPart p = parts[i];
                p.Index = i + 1;
                string name = q.FilePrefix + "_" + FileLabel(p.GroupStartMs, p.GroupEndMs, q.Zone);
                if (parts.Count > 1 && string.Equals(q.SplitBy, "Rows", StringComparison.OrdinalIgnoreCase))
                    name = q.FilePrefix + "_" + FileLabel(q.StartMs, q.EndMs, q.Zone) + "_part" + p.PartInGroup.ToString("00") + "of" + p.PartsInGroup.ToString("00");
                else if (p.PartsInGroup > 1)
                    name += "_part" + p.PartInGroup + "of" + p.PartsInGroup;
                p.BaseName = name;
            }
            return parts;
        }

        static DateTime GroupStart(long unixMs, TimeZoneInfo zone, bool week)
        {
            DateTime local = TimeZoneInfo.ConvertTime(DateTimeOffset.FromUnixTimeMilliseconds(unixMs), zone).DateTime.Date;
            if (!week) return local;
            int shift = ((int)local.DayOfWeek + 6) % 7;   // Monday = 0
            return local.AddDays(-shift);
        }

        static DateTime Local(long unixMs, TimeZoneInfo zone)
        {
            return TimeZoneInfo.ConvertTime(DateTimeOffset.FromUnixTimeMilliseconds(unixMs), zone).DateTime;
        }

        /// <summary>Whole days are shown with an inclusive last day; other bounds show times.</summary>
        public static string PeriodLabel(long startMs, long endMs, TimeZoneInfo zone)
        {
            DateTime s = Local(startMs, zone), e = Local(endMs, zone);
            if (s.TimeOfDay == TimeSpan.Zero && e.TimeOfDay == TimeSpan.Zero)
            {
                DateTime last = e.AddDays(-1);
                return last <= s ? s.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
                                 : s.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) + " to " + last.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            }
            return s.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture) + " to " + e.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture);
        }

        public static string FileLabel(long startMs, long endMs, TimeZoneInfo zone)
        {
            return PeriodLabel(startMs, endMs, zone).Replace(" to ", "_to_").Replace(" ", "_").Replace(":", "");
        }
    }

    public static class ReportWriter
    {
        static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions { Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping };

        public static ReportResult Write(IEnumerable<ReportRow> rows, List<ReportPart> plan, ReportRequest q)
        {
            var result = new ReportResult { Parts = plan };
            Directory.CreateDirectory(q.OutputDirectory);
            string[] template = null;
            if (q.WriteHtml)
            {
                string text = File.ReadAllText(q.HtmlTemplatePath, Encoding.UTF8);
                int a = text.IndexOf("%%CHUNKS%%", StringComparison.Ordinal), b = text.IndexOf("%%META%%", StringComparison.Ordinal);
                if (a < 0 || b < a || a != text.LastIndexOf("%%CHUNKS%%", StringComparison.Ordinal) || b != text.LastIndexOf("%%META%%", StringComparison.Ordinal))
                    throw new InvalidDataException("The HTML template must contain %%CHUNKS%% then %%META%%, exactly once each: " + q.HtmlTemplatePath);
                template = new[] { text.Substring(0, a), text.Substring(a + 10, b - a - 10), text.Substring(b + 8) };
            }
            var allSenders = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            using (IEnumerator<ReportRow> e = rows.GetEnumerator())
            {
                foreach (ReportPart part in plan)
                {
                    CsvOut csv = q.WriteCsv ? new CsvOut(Path.Combine(q.OutputDirectory, part.BaseName + ".csv"), q) : null;
                    HtmlOut html = q.WriteHtml ? new HtmlOut(Path.Combine(q.OutputDirectory, part.BaseName + ".html"), q, template) : null;
                    try
                    {
                        for (long i = 0; i < part.RowCount; i++)
                        {
                            if (!e.MoveNext()) throw new InvalidOperationException("The database returned fewer rows than planned; the report is not written.");
                            ReportRow row = e.Current;
                            result.Messages++;
                            result.Events += row.Events;
                            if (row.CountConflict) result.CountConflicts++;
                            if (!row.RecipientCount.HasValue || row.CountConflict) result.MessagesWithoutCount++;
                            else result.RecipientTotal += row.RecipientCount.Value;
                            allSenders.Add(row.Sender ?? "");
                            if (csv != null) csv.Add(row);
                            if (html != null) html.Add(row);
                        }
                        if (csv != null) { csv.Close(); result.Files.Add(new ReportFile { Path = csv.FilePath, Kind = "CSV", Rows = part.RowCount, Bytes = new FileInfo(csv.FilePath).Length, PartIndex = part.Index, Label = part.GroupLabel }); }
                        if (html != null)
                        {
                            html.Close(part, plan);
                            result.Files.Add(new ReportFile { Path = html.FilePath, Kind = "HTML", Rows = part.RowCount, Bytes = new FileInfo(html.FilePath).Length, PartIndex = part.Index, Label = part.GroupLabel });
                        }
                    }
                    finally
                    {
                        if (csv != null) csv.Dispose();
                        if (html != null) html.Dispose();
                    }
                }
                if (e.MoveNext()) throw new InvalidOperationException("The database returned more rows than planned; the report is not written.");
            }
            result.Senders = allSenders.Count;
            return result;
        }

        public static string FormatLocal(long unixMs, TimeZoneInfo zone)
        {
            return TimeZoneInfo.ConvertTime(DateTimeOffset.FromUnixTimeMilliseconds(unixMs), zone).ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// CSV cell for Excel: quoted when needed. Text starting with = + - @ (or a control
        /// character) is prefixed with an apostrophe so Excel never evaluates it as a formula.
        /// </summary>
        public static string SafeCsv(string value, string delimiter)
        {
            if (string.IsNullOrEmpty(value)) return "";
            string v = value;
            char first = v[0];
            if (first == '=' || first == '+' || first == '-' || first == '@' || first == '\t' || first == '\r') v = "'" + v;
            bool quote = v.Contains(delimiter) || v.IndexOf('"') >= 0 || v.IndexOf('\n') >= 0 || v.IndexOf('\r') >= 0;
            return quote ? "\"" + v.Replace("\"", "\"\"") + "\"" : v;
        }

        // ---- CSV ------------------------------------------------------------

        sealed class CsvOut : IDisposable
        {
            readonly StreamWriter _w;
            readonly ReportRequest _q;
            public string FilePath;
            public CsvOut(string path, ReportRequest q)
            {
                FilePath = path;
                _q = q;
                _w = new StreamWriter(new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.Read, 1 << 16), new UTF8Encoding(true));
                var header = new List<string> { "Detection time (" + q.TimeZoneLabel + ")", "Sender" };
                if (q.IncludeRecipientDetails) header.Add("Recipients");
                header.AddRange(new[] { "Subject", "Recipient count", "Message ID" });
                _w.WriteLine(string.Join(q.CsvDelimiter, header.Select(h => SafeCsv(h, q.CsvDelimiter))));
            }
            public void Add(ReportRow r)
            {
                string d = _q.CsvDelimiter;
                _w.Write(FormatLocal(r.FirstMs, _q.Zone)); _w.Write(d);
                _w.Write(SafeCsv(r.Sender, d)); _w.Write(d);
                if (_q.IncludeRecipientDetails) { _w.Write(SafeCsv(r.Recipients, d)); _w.Write(d); }
                _w.Write(SafeCsv(r.Subject, d)); _w.Write(d);
                if (r.RecipientCount.HasValue && !r.CountConflict) _w.Write(r.RecipientCount.Value.ToString(CultureInfo.InvariantCulture));
                _w.Write(d);
                _w.WriteLine(SafeCsv(r.MessageId, d));
            }
            public void Close() { _w.Flush(); _w.Dispose(); }
            public void Dispose() { _w.Dispose(); }
        }

        // ---- HTML -----------------------------------------------------------
        // The HTML file is self-contained. Rows are written in compressed chunks
        // (JSON -> gzip -> base64) inside <script> blocks; the page decompresses them
        // in the browser. Senders, subjects and recipients are stored once in
        // dictionaries and referenced by number, which keeps large files small.

        sealed class HtmlOut : IDisposable
        {
            readonly StreamWriter _w;
            readonly ReportRequest _q;
            readonly string[] _template;
            public string FilePath;
            readonly Dictionary<string, int> _senders = new Dictionary<string, int>(StringComparer.Ordinal);
            readonly Dictionary<string, int> _subjects = new Dictionary<string, int>(StringComparer.Ordinal);
            readonly Dictionary<string, int> _people = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            readonly List<string> _newSenders = new List<string>(), _newSubjects = new List<string>(), _newPeople = new List<string>();
            readonly List<long> _t = new List<long>(), _n = new List<long>();
            readonly List<int> _s = new List<int>(), _j = new List<int>(), _rl = new List<int>(), _ri = new List<int>();
            readonly List<string> _m = new List<string>();
            long _rows, _chunks, _recipientTotal, _withoutCount;

            public HtmlOut(string path, ReportRequest q, string[] template)
            {
                FilePath = path;
                _q = q;
                _template = template;
                _w = new StreamWriter(new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.Read, 1 << 16), new UTF8Encoding(false));
                _w.Write(template[0]);
            }

            static int Index(Dictionary<string, int> map, List<string> added, string value)
            {
                value = value ?? "";
                int i;
                if (!map.TryGetValue(value, out i)) { i = map.Count; map.Add(value, i); added.Add(value); }
                return i;
            }

            public void Add(ReportRow r)
            {
                _rows++;
                // Local wall-clock seconds: the page formats them as UTC, so the browser's own
                // time zone never changes what is displayed.
                _t.Add((r.FirstMs + (long)_q.Zone.GetUtcOffset(DateTimeOffset.FromUnixTimeMilliseconds(r.FirstMs)).TotalMilliseconds) / 1000);
                _s.Add(Index(_senders, _newSenders, r.Sender));
                _j.Add(Index(_subjects, _newSubjects, r.Subject));
                bool hasCount = r.RecipientCount.HasValue && !r.CountConflict;
                _n.Add(hasCount ? r.RecipientCount.Value : -1);
                if (hasCount) _recipientTotal += r.RecipientCount.Value; else _withoutCount++;
                _m.Add(r.MessageId);
                if (_q.IncludeRecipientDetails)
                {
                    string[] people = string.IsNullOrEmpty(r.Recipients) ? new string[0] : r.Recipients.Split(new[] { "; " }, StringSplitOptions.RemoveEmptyEntries);
                    _rl.Add(people.Length);
                    foreach (string p in people) _ri.Add(Index(_people, _newPeople, p.Trim()));
                }
                if (_t.Count >= _q.HtmlChunkRows) Flush();
            }

            void Flush()
            {
                if (_t.Count == 0) return;
                var chunk = new Dictionary<string, object>
                {
                    { "ds", _newSenders }, { "dj", _newSubjects }, { "t", _t }, { "s", _s }, { "j", _j }, { "n", _n }, { "m", _m }
                };
                if (_q.IncludeRecipientDetails) { chunk["dr"] = _newPeople; chunk["rl"] = _rl; chunk["ri"] = _ri; }
                byte[] json = JsonSerializer.SerializeToUtf8Bytes(chunk, JsonOptions);
                using (var buffer = new MemoryStream())
                {
                    using (var gzip = new GZipStream(buffer, CompressionLevel.Optimal, true)) gzip.Write(json, 0, json.Length);
                    _w.Write("<script type=\"application/x-dlp-chunk\">");
                    _w.Write(Convert.ToBase64String(buffer.GetBuffer(), 0, (int)buffer.Length));
                    _w.Write("</script>\n");
                }
                _chunks++;
                foreach (var list in new List<string>[] { _newSenders, _newSubjects, _newPeople, _m }) list.Clear();
                _t.Clear(); _n.Clear(); _s.Clear(); _j.Clear(); _rl.Clear(); _ri.Clear();
            }

            public void Close(ReportPart part, List<ReportPart> plan)
            {
                Flush();
                var meta = new Dictionary<string, object>
                {
                    { "title", _q.Title },
                    { "policy", _q.PolicyName }, { "rule", _q.RuleName },
                    { "rangeName", _q.RangeName },
                    { "timeZone", _q.TimeZoneLabel },
                    { "periodStart", FormatLocal(_q.StartMs, _q.Zone) },
                    { "periodEnd", FormatLocal(_q.EndMs, _q.Zone) },
                    { "fileStart", FormatLocal(part.StartMs, _q.Zone) },
                    { "fileEnd", FormatLocal(part.EndMs, _q.Zone) },
                    { "fileLabel", part.GroupLabel },
                    { "partIndex", part.Index }, { "partCount", plan.Count },
                    { "partInGroup", part.PartInGroup }, { "partsInGroup", part.PartsInGroup },
                    { "rows", _rows }, { "chunks", _chunks },
                    { "recipientTotal", _recipientTotal }, { "withoutCount", _withoutCount },
                    { "includeRecipients", _q.IncludeRecipientDetails },
                    { "coveragePercent", Math.Round(_q.CoveragePercent, 2) },
                    { "coverageNote", _q.CoverageNote },
                    { "generated", DateTimeOffset.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) },
                    { "toolVersion", _q.ToolVersion },
                    { "siblings", plan.Select(p => new Dictionary<string, object> {
                        { "file", p.BaseName + ".html" }, { "label", p.GroupLabel + (p.PartsInGroup > 1 ? " (part " + p.PartInGroup + " of " + p.PartsInGroup + ")" : "") },
                        { "rows", p.RowCount }, { "index", p.Index } }).ToList() }
                };
                string json = JsonSerializer.Serialize(meta, JsonOptions).Replace("</", "<\\/");
                _w.Write(_template[1]);
                _w.Write(json);
                _w.Write(_template[2]);
                _w.Flush();
                _w.Dispose();
            }

            public void Dispose() { _w.Dispose(); }
        }
    }
}
