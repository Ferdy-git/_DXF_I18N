using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using Microsoft.Data.Sqlite;

namespace DxfI18N.App
{
    public static class Db
    {
        private static string? _dbPath;

        public static void SetDatabasePath(string path)
        {
            _dbPath = path;
        }

        public static string? GetDatabasePath() => _dbPath;

        private static SqliteConnection GetConn()
        {
            if (string.IsNullOrWhiteSpace(_dbPath))
                throw new InvalidOperationException("Percorso DB non impostato.");
            var cs = new SqliteConnectionStringBuilder
            {
                DataSource = _dbPath,
                Mode = SqliteOpenMode.ReadWrite,
                Cache = SqliteCacheMode.Default
            }.ToString();
            return new SqliteConnection(cs);
        }

        /// <summary>
        /// Verifica accesso al DB e applica piccole migrazioni automatiche se mancano colonne.
        /// </summary>
        public static bool QuickHealthCheck()
        {
            using var conn = GetConn();
            conn.Open();

            // ping
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "SELECT 1;";
                _ = cmd.ExecuteScalar();
            }

            // MIGRAZIONE: aggiungi AlignX/AlignY a TextOccurrences se mancanti
            EnsureColumn(conn, "TextOccurrences", "AlignX", "REAL");
            EnsureColumn(conn, "TextOccurrences", "AlignY", "REAL");

            return true;
        }

        private static void EnsureColumn(SqliteConnection conn, string table, string column, string type)
        {
            using var check = conn.CreateCommand();
            check.CommandText = $"PRAGMA table_info({table});";
            using var rd = check.ExecuteReader();
            bool exists = false;
            while (rd.Read())
            {
                var colName = rd.GetString(1);
                if (string.Equals(colName, column, StringComparison.OrdinalIgnoreCase))
                {
                    exists = true; break;
                }
            }
            if (!exists)
            {
                using var alter = conn.CreateCommand();
                alter.CommandText = $"ALTER TABLE {table} ADD COLUMN {column} {type};";
                alter.ExecuteNonQuery();
            }
        }

        // =====================================================================
        //                              LISTE BASE
        // =====================================================================

        public static List<string> GetCultures()
        {
            var list = new List<string>();
            using var conn = GetConn();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT Code FROM Cultures ORDER BY Code;";
            using var rd = cmd.ExecuteReader();
            while (rd.Read()) list.Add(rd.GetString(0));
            return list;
        }

        public static List<(string CultureCode, string LayerName)> GetCultureLayers()
        {
            var list = new List<(string, string)>();
            using var conn = GetConn();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT CultureCode, LayerName FROM LayerRules ORDER BY CultureCode;";
            using var rd = cmd.ExecuteReader();
            while (rd.Read()) list.Add((rd.GetString(0), rd.GetString(1)));
            return list;
        }

        public static List<(int id, string path, string importedAt)> GetDrawings()
        {
            var list = new List<(int, string, string)>();
            using var conn = GetConn();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT DrawingId, Path, IFNULL(ImportedAt,'') FROM Drawings ORDER BY DrawingId DESC;";
            using var rd = cmd.ExecuteReader();
            while (rd.Read()) list.Add((rd.GetInt32(0), rd.GetString(1), rd.GetString(2)));
            return list;
        }

        // =====================================================================
        //                         DRAWINGS & FILE-HASH
        // =====================================================================

        public static string ComputeFileHash(string filePath)
        {
            using var md5 = MD5.Create();
            using var fs = File.OpenRead(filePath);
            var hash = md5.ComputeHash(fs);
            return Convert.ToHexString(hash);
        }

        public static int UpsertDrawing(string path, string fileHash)
        {
            using var conn = GetConn();
            conn.Open();

            using (var check = conn.CreateCommand())
            {
                check.CommandText = "SELECT DrawingId FROM Drawings WHERE Path=$p;";
                check.Parameters.AddWithValue("$p", path);
                var existing = check.ExecuteScalar();
                if (existing is long l) return (int)l;
            }

            using (var ins = conn.CreateCommand())
            {
                ins.CommandText = @"INSERT INTO Drawings (Path, ImportedAt, FileHash)
                                    VALUES ($p, datetime('now'), $h);
                                    SELECT last_insert_rowid();";
                ins.Parameters.AddWithValue("$p", path);
                ins.Parameters.AddWithValue("$h", fileHash);
                var id = ins.ExecuteScalar();
                return Convert.ToInt32((long)id!);
            }
        }

        public static void UpdateDrawingHash(int drawingId, string fileHash)
        {
            using var conn = GetConn();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"UPDATE Drawings SET FileHash=$h, ImportedAt=datetime('now') WHERE DrawingId=$id;";
            cmd.Parameters.AddWithValue("$h", fileHash);
            cmd.Parameters.AddWithValue("$id", drawingId);
            cmd.ExecuteNonQuery();
        }

        public static void DeleteOccurrencesForDrawing(int drawingId)
        {
            using var conn = GetConn();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"DELETE FROM TextOccurrences WHERE DrawingId=$d;";
            cmd.Parameters.AddWithValue("$d", drawingId);
            cmd.ExecuteNonQuery();
        }

        // =====================================================================
        //                               TEXT KEYS
        // =====================================================================

        public static int GetOrCreateTextKey(
            string defaultText,
            string? normalized = null,
            string defaultCulture = "it-IT",
            string? context = null)
        {
            using var conn = GetConn();
            conn.Open();

            // match esatto
            using (var find = conn.CreateCommand())
            {
                find.CommandText = @"SELECT KeyId FROM TextKeys
                                     WHERE DefaultCulture=$c AND DefaultText=$t AND IFNULL(Context,'')=IFNULL($ctx,'')
                                     LIMIT 1;";
                find.Parameters.AddWithValue("$c", defaultCulture);
                find.Parameters.AddWithValue("$t", defaultText);
                find.Parameters.AddWithValue("$ctx", (object?)context ?? DBNull.Value);
                var id = find.ExecuteScalar();
                if (id is long l1) return (int)l1;
            }

            // match normalizzato
            if (!string.IsNullOrWhiteSpace(normalized))
            {
                using var findN = conn.CreateCommand();
                findN.CommandText = @"SELECT KeyId FROM TextKeys
                                      WHERE DefaultCulture=$c AND NormalizedDefaultText=$n AND IFNULL(Context,'')=IFNULL($ctx,'')
                                      LIMIT 1;";
                findN.Parameters.AddWithValue("$c", defaultCulture);
                findN.Parameters.AddWithValue("$n", normalized);
                findN.Parameters.AddWithValue("$ctx", (object?)context ?? DBNull.Value);
                var id2 = findN.ExecuteScalar();
                if (id2 is long l2) return (int)l2;
            }

            // crea
            using var ins = conn.CreateCommand();
            ins.CommandText = @"INSERT INTO TextKeys (DefaultText, DefaultCulture, Context, NormalizedDefaultText, CreatedAt)
                                VALUES ($t,$c,$ctx,$n,datetime('now'));
                                SELECT last_insert_rowid();";
            ins.Parameters.AddWithValue("$t", defaultText);
            ins.Parameters.AddWithValue("$c", defaultCulture);
            ins.Parameters.AddWithValue("$ctx", (object?)context ?? DBNull.Value);
            ins.Parameters.AddWithValue("$n", (object?)normalized ?? DBNull.Value);
            var newId = ins.ExecuteScalar();
            return Convert.ToInt32((long)newId!);
        }

        // =====================================================================
        //                            OCCURRENCES: INSERT
        // =====================================================================

        public static void InsertOccurrence(
            int drawingId,
            string entityHandle,
            string entityType,
            string layerOriginal,
            double posX,
            double posY,
            double posZ,
            double? rotation,
            double? height,
            string? style,
            double? widthFactor,
            string? attachment,
            double? wrapWidth,
            string originalRaw,
            string? originalNormalized,
            int? keyId,
            string? blockName = null,
            string? attribTag = null,
            double? alignX = null,
            double? alignY = null
        )
        {
            using var conn = GetConn();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
INSERT INTO TextOccurrences
 (DrawingId, EntityHandle, EntityType, LayerOriginal, PosX, PosY, PosZ, Rotation, Height, Style, WidthFactor, Attachment, WrapWidth, OriginalTextRaw, OriginalTextNormalized, KeyId, BlockName, AttribTag, AlignX, AlignY)
 VALUES
 ($d,$h,$t,$l,$x,$y,$z,$rot,$ht,$st,$wf,$att,$ww,$raw,$norm,$key,$blk,$tag,$ax,$ay);";
            cmd.Parameters.AddWithValue("$d", drawingId);
            cmd.Parameters.AddWithValue("$h", entityHandle);
            cmd.Parameters.AddWithValue("$t", entityType);
            cmd.Parameters.AddWithValue("$l", layerOriginal);
            cmd.Parameters.AddWithValue("$x", posX);
            cmd.Parameters.AddWithValue("$y", posY);
            cmd.Parameters.AddWithValue("$z", posZ);
            cmd.Parameters.AddWithValue("$rot", (object?)rotation ?? DBNull.Value);
            cmd.Parameters.AddWithValue("$ht", (object?)height ?? DBNull.Value);
            cmd.Parameters.AddWithValue("$st", (object?)style ?? DBNull.Value);
            cmd.Parameters.AddWithValue("$wf", (object?)widthFactor ?? DBNull.Value);
            cmd.Parameters.AddWithValue("$att", (object?)attachment ?? DBNull.Value);
            cmd.Parameters.AddWithValue("$ww", (object?)wrapWidth ?? DBNull.Value);
            cmd.Parameters.AddWithValue("$raw", originalRaw);
            cmd.Parameters.AddWithValue("$norm", (object?)originalNormalized ?? DBNull.Value);
            cmd.Parameters.AddWithValue("$key", (object?)keyId ?? DBNull.Value);
            cmd.Parameters.AddWithValue("$blk", (object?)blockName ?? DBNull.Value);
            cmd.Parameters.AddWithValue("$tag", (object?)attribTag ?? DBNull.Value);
            cmd.Parameters.AddWithValue("$ax", (object?)alignX ?? DBNull.Value);
            cmd.Parameters.AddWithValue("$ay", (object?)alignY ?? DBNull.Value);
            cmd.ExecuteNonQuery();
        }

        // =====================================================================
        //                              TRANSLATIONS
        // =====================================================================

        public static List<(int keyId, string defaultText)> GetMissingKeysForCulture(string cultureCode)
        {
            var list = new List<(int, string)>();
            using var conn = GetConn();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
SELECT KeyId, DefaultText
FROM v_MissingTranslations
WHERE CultureCode=$c
ORDER BY KeyId;";
            cmd.Parameters.AddWithValue("$c", cultureCode);
            using var rd = cmd.ExecuteReader();
            while (rd.Read()) list.Add((rd.GetInt32(0), rd.GetString(1)));
            return list;
        }

        public static List<(int keyId, string defaultText, string translated, string status, int warn, string notes)>
            GetTranslationsForCulture(string cultureCode, string? filter)
        {
            var list = new List<(int, string, string, string, int, string)>();
            using var conn = GetConn();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
SELECT tk.KeyId,
       tk.DefaultText,
       IFNULL(t.Translated,'')                    AS Translated,
       CASE WHEN t.KeyId IS NULL THEN 'Missing'
            ELSE IFNULL(t.Status,'Draft') END     AS Status,
       IFNULL(t.OverLengthWarn,0)                 AS Warn,
       IFNULL(t.Notes,'')                         AS Notes
FROM TextKeys tk
LEFT JOIN Translations t
       ON t.KeyId = tk.KeyId
      AND t.CultureCode = $c
WHERE ($f = '' OR
       UPPER(tk.DefaultText) LIKE UPPER($like) OR
       UPPER(IFNULL(t.Translated,'')) LIKE UPPER($like))
ORDER BY tk.KeyId;";
            var f = (filter ?? "").Trim();
            cmd.Parameters.AddWithValue("$c", cultureCode);
            cmd.Parameters.AddWithValue("$f", f);
            cmd.Parameters.AddWithValue("$like", $"%{f}%");
            using var rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                list.Add((
                    rd.GetInt32(0),
                    rd.GetString(1),
                    rd.GetString(2),
                    rd.GetString(3),
                    rd.GetInt32(4),
                    rd.GetString(5)
                ));
            }
            return list;
        }

        public static List<(int keyId, string defaultText, string culture, string translated, string status)>
            GetAllForCultureExport(string cultureCode)
        {
            var list = new List<(int, string, string, string, string)>();
            using var conn = GetConn();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
SELECT tk.KeyId, tk.DefaultText, $c AS Culture,
       IFNULL(t.Translated,'') AS Translated,
       IFNULL(t.Status,'Draft') AS Status
FROM TextKeys tk
LEFT JOIN Translations t
  ON t.KeyId = tk.KeyId AND t.CultureCode = $c
ORDER BY tk.KeyId;";
            cmd.Parameters.AddWithValue("$c", cultureCode);
            using var rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                list.Add((rd.GetInt32(0), rd.GetString(1), cultureCode,
                          rd.GetString(3), rd.GetString(4)));
            }
            return list;
        }

        public static void UpsertTranslation(int keyId, string cultureCode, string translated, string? status, bool overwriteApproved)
        {
            string normStatus = string.IsNullOrWhiteSpace(status) ? "Draft" : status.Trim();
            int len = translated?.Length ?? 0;
            int warn = 0;

            using var conn = GetConn();
            conn.Open();

            // calcola warn (±20%) contro DefaultText
            using (var getLen = conn.CreateCommand())
            {
                getLen.CommandText = "SELECT LENGTH(DefaultText) FROM TextKeys WHERE KeyId=$k;";
                getLen.Parameters.AddWithValue("$k", keyId);
                var baseLenObj = getLen.ExecuteScalar();
                int baseLen = baseLenObj is long l ? (int)l : 0;
                if (baseLen > 0)
                {
                    double diff = Math.Abs(len - baseLen) / (double)baseLen;
                    warn = diff > 0.20 ? 1 : 0;
                }
            }

            // esiste già?
            using (var chk = conn.CreateCommand())
            {
                chk.CommandText = "SELECT Status FROM Translations WHERE KeyId=$k AND CultureCode=$c;";
                chk.Parameters.AddWithValue("$k", keyId);
                chk.Parameters.AddWithValue("$c", cultureCode);
                var st = chk.ExecuteScalar() as string;

                if (st != null)
                {
                    if (st == "Approved" && !overwriteApproved) return;
                    using var up = conn.CreateCommand();
                    up.CommandText = @"UPDATE Translations
                                       SET Translated=$t, Status=$s, CharCount=$n, OverLengthWarn=$w, UpdatedAt=datetime('now')
                                       WHERE KeyId=$k AND CultureCode=$c;";
                    up.Parameters.AddWithValue("$t", translated ?? "");
                    up.Parameters.AddWithValue("$s", normStatus);
                    up.Parameters.AddWithValue("$n", len);
                    up.Parameters.AddWithValue("$w", warn);
                    up.Parameters.AddWithValue("$k", keyId);
                    up.Parameters.AddWithValue("$c", cultureCode);
                    up.ExecuteNonQuery();
                    return;
                }
            }

            using (var ins = conn.CreateCommand())
            {
                ins.CommandText = @"INSERT INTO Translations (KeyId, CultureCode, Translated, Status, Notes, CharCount, OverLengthWarn, UpdatedAt)
                                    VALUES ($k,$c,$t,$s,NULL,$n,$w,datetime('now'));";
                ins.Parameters.AddWithValue("$k", keyId);
                ins.Parameters.AddWithValue("$c", cultureCode);
                ins.Parameters.AddWithValue("$t", translated ?? "");
                ins.Parameters.AddWithValue("$s", normStatus);
                ins.Parameters.AddWithValue("$n", len);
                ins.Parameters.AddWithValue("$w", warn);
                ins.ExecuteNonQuery();
            }
        }

        public static void UpdateStatus(int keyId, string cultureCode, string newStatus)
        {
            if (string.IsNullOrWhiteSpace(newStatus)) newStatus = "Draft";
            using var conn = GetConn();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"UPDATE Translations
                                SET Status=$s, UpdatedAt=datetime('now')
                                WHERE KeyId=$k AND CultureCode=$c;";
            cmd.Parameters.AddWithValue("$s", newStatus);
            cmd.Parameters.AddWithValue("$k", keyId);
            cmd.Parameters.AddWithValue("$c", cultureCode);
            cmd.ExecuteNonQuery();
        }

        // =====================================================================
        //                          OCCURRENCES: READ
        // =====================================================================

        public class OccurrenceRow
        {
            public int DrawingId { get; set; }
            public string EntityType { get; set; } = "";
            public string LayerOriginal { get; set; } = "";
            public double PosX { get; set; }
            public double PosY { get; set; }
            public double PosZ { get; set; }
            public double? Rotation { get; set; }
            public double? Height { get; set; }
            public string? Style { get; set; }
            public string EntityHandle { get; set; } = "";
            public string? BlockName { get; set; }
            public string? AttribTag { get; set; }
            public int KeyId { get; set; }
            public string DefaultText { get; set; } = "";

            public double? WidthFactor { get; set; }
            public string? Attachment { get; set; }
            public double? WrapWidth { get; set; }

            public double? AlignX { get; set; }
            public double? AlignY { get; set; }
        }

        public static List<OccurrenceRow> GetOccurrencesForDrawing(int drawingId)
        {
            var list = new List<OccurrenceRow>();
            using var conn = GetConn();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
SELECT o.DrawingId, o.EntityType, o.LayerOriginal, o.PosX, o.PosY, o.PosZ,
       o.Rotation, o.Height, o.Style, o.EntityHandle, o.BlockName, o.AttribTag,
       o.KeyId, tk.DefaultText,
       o.WidthFactor, o.Attachment, o.WrapWidth,
       o.AlignX, o.AlignY
FROM TextOccurrences o
JOIN TextKeys tk ON tk.KeyId = o.KeyId
WHERE o.DrawingId = $d
ORDER BY o.OccurrenceId;";
            cmd.Parameters.AddWithValue("$d", drawingId);
            using var rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                var r = new OccurrenceRow
                {
                    DrawingId = rd.GetInt32(0),
                    EntityType = rd.GetString(1),
                    LayerOriginal = rd.GetString(2),
                    PosX = rd.GetDouble(3),
                    PosY = rd.GetDouble(4),
                    PosZ = rd.GetDouble(5),
                    Rotation = rd.IsDBNull(6) ? (double?)null : rd.GetDouble(6),
                    Height = rd.IsDBNull(7) ? (double?)null : rd.GetDouble(7),
                    Style = rd.IsDBNull(8) ? null : rd.GetString(8),
                    EntityHandle = rd.GetString(9),
                    BlockName = rd.IsDBNull(10) ? null : rd.GetString(10),
                    AttribTag = rd.IsDBNull(11) ? null : rd.GetString(11),
                    KeyId = rd.GetInt32(12),
                    DefaultText = rd.GetString(13),
                    WidthFactor = rd.IsDBNull(14) ? (double?)null : rd.GetDouble(14),
                    Attachment = rd.IsDBNull(15) ? null : rd.GetString(15),
                    WrapWidth = rd.IsDBNull(16) ? (double?)null : rd.GetDouble(16),
                    AlignX = rd.IsDBNull(17) ? (double?)null : rd.GetDouble(17),
                    AlignY = rd.IsDBNull(18) ? (double?)null : rd.GetDouble(18),
                };
                list.Add(r);
            }
            return list;
        }

        public static Dictionary<int, string> GetTranslationsMap(string cultureCode, bool onlyApproved)
        {
            var dict = new Dictionary<int, string>();
            using var conn = GetConn();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = onlyApproved
                ? @"SELECT KeyId, Translated FROM Translations WHERE CultureCode=$c AND Status='Approved';"
                : @"SELECT KeyId, Translated FROM Translations WHERE CultureCode=$c;";
            cmd.Parameters.AddWithValue("$c", cultureCode);
            using var rd = cmd.ExecuteReader();
            while (rd.Read()) dict[rd.GetInt32(0)] = rd.GetString(1);
            return dict;
        }
    }
}
