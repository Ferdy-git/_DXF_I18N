using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection; // reflection per AlignmentPoint netDxf
using System.Text.RegularExpressions;
using System.Windows.Forms;
using CsvHelper;

// ==== LIBRERIA PER L'ANTEPRIMA/IMPORT (come prima) ====
using netDxf;
using netDxf.Entities;
using netDxf.Tables;
// alias per evitare conflitto con System.Attribute
using DxfAttribute = netDxf.Entities.Attribute;

// ==== LIBRERIA PER LA GENERAZIONE (NUOVA, precisa) ====
using iDxf = IxMilia.Dxf;
using iEnt = IxMilia.Dxf.Entities;
using iTables = IxMilia.Dxf.Tables;

namespace DxfI18N.App
{
    public partial class Form1 : Form
    {
        private readonly List<string> _selectedDxfFiles = new();

        public Form1()
        {
            InitializeComponent();

            // Nascondi i controlli “storici” non necessari
            lblLangOut.Visible = false;
            cmbCultureGenerate.Visible = false;

            // ----- Importa DXF -----
            btnSelectDb.Click += (s, e) => OnSelectDb();
            btnPickDxf.Click += (s, e) => OnPickDxf();
            btnPreview.Click += (s, e) => OnPreview();
            btnApplyFilters.Click += (s, e) => OnApplyFilters();
            btnSelectAllPreview.Click += (s, e) => OnSelectAllPreview(true);
            btnDeselectAllPreview.Click += (s, e) => OnSelectAllPreview(false);
            btnImport.Click += (s, e) => OnImport();

            gridOccurrences.CellContentClick += (s, e) =>
            {
                if (e.RowIndex >= 0 && e.ColumnIndex == gridOccurrences.Columns["colInclude"].Index)
                    gridOccurrences.CommitEdit(DataGridViewDataErrorContexts.Commit);
            };

            // ----- Traduzioni -----
            cmbCultureTranslations.SelectedIndexChanged += (s, e) => LoadTranslationsGrid();
            txtFilter.TextChanged += (s, e) => LoadTranslationsGrid();
            btnExportCsv.Click += (s, e) => OnExportMissingCsv();
            btnImportCsv.Click += (s, e) => OnImportCsv();
            btnExportLanguage.Click += (s, e) => OnExportWholeLanguage();
            btnSetApproved.Click += (s, e) => OnSetStatusSelected("Approved");
            btnSetDraft.Click += (s, e) => OnSetStatusSelected("Draft");
            btnSetBlocked.Click += (s, e) => OnSetStatusSelected("Blocked");

            // ----- DXF Unico -----
            btnGenerateUnified.Click += (s, e) => OnGenerateUnified();

            ConfigureGridForManualResize(gridOccurrences);
            ConfigureGridForManualResize(gridTranslations);
            ConfigureGridForManualResize(gridDrawings);
        }

        private void ConfigureGridForManualResize(DataGridView g)
        {
            g.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            g.AllowUserToResizeColumns = true;
            g.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
        }

        private static double SafeToDouble(object? v) => v == null ? 0.0 : Convert.ToDouble(v);

        private static double? SafeToNullableDouble(object? v)
        {
            if (v == null) return null;
            var s = v.ToString();
            if (string.IsNullOrWhiteSpace(s)) return null;
            if (double.TryParse(s, out var d)) return d;
            try { return Convert.ToDouble(v); } catch { return null; }
        }

        private static string NormalizeDxfText(string raw, bool isMText)
        {
            if (string.IsNullOrEmpty(raw)) return string.Empty;
            var s = raw;
            if (isMText)
            {
                s = s.Replace("\\P", "\n");
                s = s.Replace("\\~", " ");
            }
            s = s.Replace("\r", "");
            var lines = s.Split('\n');
            for (int i = 0; i < lines.Length; i++)
                lines[i] = CompressSpaces(lines[i]);
            return string.Join("\n", lines).Trim();
        }

        private static string CompressSpaces(string line)
        {
            if (string.IsNullOrEmpty(line)) return string.Empty;
            var res = new System.Text.StringBuilder(line.Length);
            bool prevSpace = false;
            foreach (var ch in line)
            {
                bool isSpace = ch == ' ' || ch == '\t';
                if (isSpace) { if (!prevSpace) res.Append(' '); prevSpace = true; }
                else { res.Append(ch); prevSpace = false; }
            }
            return res.ToString().Trim();
        }

        // =====================================================================
        //                           TAB: Importa DXF
        // =====================================================================

        private void OnSelectDb()
        {
            using var dlg = new OpenFileDialog
            {
                Title = "Seleziona database SQLite",
                Filter = "SQLite DB (*.sqlite;*.db;*.db3)|*.sqlite;*.db;*.db3|Tutti i file (*.*)|*.*",
                CheckFileExists = true,
                Multiselect = false
            };
            if (dlg.ShowDialog(this) != DialogResult.OK) return;

            try
            {
                Db.SetDatabasePath(dlg.FileName);
                Db.QuickHealthCheck();
                tsslDbPath.Text = $"DB: {dlg.FileName}";
                LoadCulturesIntoCombos();
                LoadDrawingsGrid();
                LoadTranslationsGrid();
                MessageBox.Show(this, "Connessione al DB: COMPLETATA", "COMPLETATO",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                tsslDbPath.Text = "DB: (errore)";
                MessageBox.Show(this, $"Errore DB:\n{ex.Message}", "ERRORE",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void OnPickDxf()
        {
            using var dlg = new OpenFileDialog
            {
                Title = "Seleziona file DXF",
                Filter = "Drawing eXchange Format (*.dxf)|*.dxf|Tutti i file (*.*)|*.*",
                CheckFileExists = true,
                Multiselect = true
            };
            if (dlg.ShowDialog(this) != DialogResult.OK) return;

            _selectedDxfFiles.Clear();
            _selectedDxfFiles.AddRange(dlg.FileNames);

            var allLayers = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            int filesOk = 0, filesErr = 0;

            foreach (var file in _selectedDxfFiles)
            {
                try
                {
                    var dxf = DxfDocument.Load(file);
                    foreach (var lay in dxf.Layers)
                        allLayers.Add(lay.Name ?? "");
                    filesOk++;
                }
                catch { filesErr++; }
            }

            clbLayers.Items.Clear();
            foreach (var name in allLayers)
                clbLayers.Items.Add(name, true);

            lblStatsImport.Text = $"Selezionati: {filesOk} file ({filesErr} errori). Layer trovati: {allLayers.Count}.";
        }

        private void OnPreview()
        {
            if (_selectedDxfFiles.Count == 0)
            {
                MessageBox.Show(this, "Seleziona prima uno o più DXF.", "MANCANO PREREQUISITI",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var selectedLayers = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < clbLayers.Items.Count; i++)
                if (clbLayers.GetItemChecked(i))
                    selectedLayers.Add(clbLayers.Items[i]?.ToString() ?? "");

            bool includeText = chkText.Checked;
            bool includeMText = chkMText.Checked;
            bool includeAttrib = chkAttrib.Checked;

            gridOccurrences.Rows.Clear();
            int totalFiles = 0, totalOcc = 0, totalShown = 0;

            foreach (var file in _selectedDxfFiles)
            {
                try
                {
                    var dxf = DxfDocument.Load(file);

                    if (includeText && dxf.Entities?.Texts != null)
                    {
                        foreach (var t in dxf.Entities.Texts)
                        {
                            var layer = t.Layer?.Name ?? "";
                            if (selectedLayers.Count > 0 && !selectedLayers.Contains(layer)) continue;

                            var raw = t.Value ?? "";
                            var include = PassesFilters(NormalizeDxfText(raw, false));

                            AddPreviewRow(include, file, layer, "TEXT", raw,
                                t.Position.X, t.Position.Y, t.Rotation, t.Height, t.Style?.Name, t.Handle, null);

                            totalOcc++; if (include) totalShown++;
                        }
                    }

                    if (includeMText && dxf.Entities?.MTexts != null)
                    {
                        foreach (var m in dxf.Entities.MTexts)
                        {
                            var layer = m.Layer?.Name ?? "";
                            if (selectedLayers.Count > 0 && !selectedLayers.Contains(layer)) continue;

                            var raw = m.Value ?? "";
                            var include = PassesFilters(NormalizeDxfText(raw, true));

                            AddPreviewRow(include, file, layer, "MTEXT", raw,
                                m.Position.X, m.Position.Y, m.Rotation, m.Height, m.Style?.Name, m.Handle, null);

                            totalOcc++; if (include) totalShown++;
                        }
                    }

                    if (includeAttrib && dxf.Entities?.Inserts != null)
                    {
                        foreach (var ins in dxf.Entities.Inserts)
                        {
                            string blockName = ins.Block?.Name ?? "";
                            if (ins.Attributes == null) continue;

                            foreach (var a in ins.Attributes)
                            {
                                var layer = a.Layer?.Name ?? ins.Layer?.Name ?? "";
                                if (selectedLayers.Count > 0 && !selectedLayers.Contains(layer)) continue;

                                string tag = a.Tag ?? "";
                                string raw = a.Value ?? "";
                                bool include = PassesFilters(NormalizeDxfText(raw, false));

                                string handle = !string.IsNullOrEmpty(a.Handle) ? a.Handle : $"{ins.Handle}/{tag}";

                                double? aHeight = null, aRot = null;
                                try { aHeight = a.Height; } catch { }
                                try { aRot = a.Rotation; } catch { }

                                AddPreviewRow(include, file, layer, "ATTRIB", raw,
                                    a.Position.X, a.Position.Y, aRot, aHeight, a.Style?.Name, handle, null,
                                    blockName, tag);

                                totalOcc++; if (include) totalShown++;
                            }
                        }
                    }

                    totalFiles++;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, $"Errore anteprima '{file}':\n{ex.Message}", "ERRORE",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            lblStatsImport.Text = $"Anteprima: file OK {totalFiles}, occorrenze trovate {totalOcc} (mostrate {totalShown}).";
        }

        private void OnApplyFilters()
        {
            int selected = 0;
            foreach (DataGridViewRow row in gridOccurrences.Rows)
            {
                string tipo = row.Cells["colEntityType"].Value?.ToString() ?? "";
                string text = row.Cells["colOriginalText"].Value?.ToString() ?? "";
                bool isMText = string.Equals(tipo, "MTEXT", StringComparison.OrdinalIgnoreCase);
                bool include = PassesFilters(NormalizeDxfText(text, isMText));
                row.Cells["colInclude"].Value = include;
                if (include) selected++;
            }
            lblStatsImport.Text = $"Anteprima aggiornata: selezionate {selected} di {gridOccurrences.Rows.Count}.";
        }

        private void OnSelectAllPreview(bool selectAll)
        {
            foreach (DataGridViewRow row in gridOccurrences.Rows)
                row.Cells["colInclude"].Value = selectAll;
            lblStatsImport.Text = $"Selezionate {(selectAll ? gridOccurrences.Rows.Count : 0)} di {gridOccurrences.Rows.Count}.";
        }

        private bool PassesFilters(string text)
        {
            string s = text?.Trim() ?? "";
            if (s.Length == 0) return false;

            bool onlyLetters = chkOnlyLetters.Checked;
            int minLen = (int)numMinLen.Value;
            string pat = txtExcludePatterns.Text?.Trim() ?? "";

            if (s.Replace("\n", "").Length < minLen) return false;
            if (onlyLetters && !HasLetter(s)) return false;

            if (!string.IsNullOrWhiteSpace(pat))
            {
                var parts = pat.Split(';');
                foreach (var p in parts)
                {
                    var rx = p.Trim();
                    if (rx.Length == 0) continue;
                    try { if (Regex.IsMatch(s, rx)) return false; } catch { }
                }
            }
            return true;
        }

        private static bool HasLetter(string s)
        {
            foreach (var ch in s) if (char.IsLetter(ch)) return true;
            return false;
        }

        private void AddPreviewRow(
            bool include, string file, string layer, string tipo, string testo,
            double x, double y, double? rot, double? h, string? stile, string handle, int? keyId,
            string? blockName = null, string? tag = null)
        {
            int idx = gridOccurrences.Rows.Add();
            var r = gridOccurrences.Rows[idx];

            r.Cells["colInclude"].Value = include;
            r.Cells["colFile"].Value = file;
            r.Cells["colLayerOriginal"].Value = layer;
            r.Cells["colEntityType"].Value = tipo;
            r.Cells["colOriginalText"].Value = testo;
            r.Cells["colPosX"].Value = x;
            r.Cells["colPosY"].Value = y;
            r.Cells["colRotation"].Value = rot;
            r.Cells["colHeight"].Value = h;
            r.Cells["colStyle"].Value = stile;
            r.Cells["colEntityHandle"].Value = handle;
            r.Cells["colKeyId"].Value = keyId;
            r.Cells["colBlockName"].Value = blockName;
            r.Cells["colTag"].Value = tag;
        }

        private void OnImport()
        {
            var dbPath = Db.GetDatabasePath();
            if (string.IsNullOrWhiteSpace(dbPath) || !File.Exists(dbPath))
            {
                MessageBox.Show(this, "Seleziona il database (Scegli DB…).", "MANCANO PREREQUISITI",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (gridOccurrences.Rows.Count == 0)
            {
                MessageBox.Show(this, "Fai prima 'Anteprima' e seleziona le righe (✓).", "MANCANO PREREQUISITI",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var byFile = new Dictionary<string, List<DataGridViewRow>>(StringComparer.OrdinalIgnoreCase);
            foreach (DataGridViewRow r in gridOccurrences.Rows)
            {
                bool include = r.Cells["colInclude"].Value is bool b && b;
                if (!include) continue;
                string file = r.Cells["colFile"].Value?.ToString() ?? "";
                if (string.IsNullOrEmpty(file)) continue;
                if (!byFile.TryGetValue(file, out var list)) { list = new(); byFile[file] = list; }
                list.Add(r);
            }
            if (byFile.Count == 0)
            {
                MessageBox.Show(this, "Nessuna riga selezionata (✓).", "MANCANO PREREQUISITI",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            int filesOk = 0, occInserted = 0;

            foreach (var kv in byFile)
            {
                string file = kv.Key;
                var rows = kv.Value;

                try
                {
                    string hash = Db.ComputeFileHash(file);
                    int drawingId = Db.UpsertDrawing(file, hash);
                    Db.UpdateDrawingHash(drawingId, hash);
                    Db.DeleteOccurrencesForDrawing(drawingId);

                    var dxf = DxfDocument.Load(file);

                    var textByHandle = dxf.Entities.Texts?.Where(t => !string.IsNullOrEmpty(t.Handle))
                                         .ToDictionary(t => t.Handle!, t => t) ?? new Dictionary<string, Text>();
                    var mtextByHandle = dxf.Entities.MTexts?.Where(m => !string.IsNullOrEmpty(m.Handle))
                                         .ToDictionary(m => m.Handle!, m => m) ?? new Dictionary<string, MText>();

                    var attribByHandle = new Dictionary<string, DxfAttribute>();
                    if (dxf.Entities.Inserts != null)
                    {
                        foreach (var ins in dxf.Entities.Inserts)
                        {
                            if (ins.Attributes == null) continue;
                            foreach (var a in ins.Attributes)
                            {
                                if (!string.IsNullOrEmpty(a.Handle))
                                    attribByHandle[a.Handle] = a;
                            }
                        }
                    }

                    foreach (var r in rows)
                    {
                        string tipo = r.Cells["colEntityType"].Value?.ToString() ?? "";
                        string layer = r.Cells["colLayerOriginal"].Value?.ToString() ?? "";
                        string raw = r.Cells["colOriginalText"].Value?.ToString() ?? "";
                        string handle = r.Cells["colEntityHandle"].Value?.ToString() ?? "";
                        string? stile = r.Cells["colStyle"].Value?.ToString();
                        string? block = r.Cells["colBlockName"].Value?.ToString();
                        string? tag = r.Cells["colTag"].Value?.ToString();

                        double x = Math.Round(SafeToDouble(r.Cells["colPosX"]?.Value), 2);
                        double y = Math.Round(SafeToDouble(r.Cells["colPosY"]?.Value), 2);
                        double? rot = SafeToNullableDouble(r.Cells["colRotation"]?.Value);
                        double? hgt = SafeToNullableDouble(r.Cells["colHeight"]?.Value);
                        if (rot.HasValue) rot = Math.Round(rot.Value, 2);
                        if (hgt.HasValue) hgt = Math.Round(hgt.Value, 2);

                        double? widthFactor = null;
                        string? attachment = null;
                        double? wrapWidth = null;
                        double? alignX = null;
                        double? alignY = null;

                        if (string.Equals(tipo, "TEXT", StringComparison.OrdinalIgnoreCase))
                        {
                            if (!string.IsNullOrEmpty(handle) && textByHandle.TryGetValue(handle, out var t))
                            {
                                widthFactor = t.WidthFactor;
                                attachment = t.Alignment.ToString();

                                // Reflection: AlignmentPoint se disponibile
                                if (!string.Equals(attachment, "BaselineLeft", StringComparison.OrdinalIgnoreCase))
                                {
                                    try
                                    {
                                        var apProp = t.GetType().GetProperty("AlignmentPoint", BindingFlags.Public | BindingFlags.Instance);
                                        if (apProp != null)
                                        {
                                            var ap = apProp.GetValue(t);
                                            if (ap != null)
                                            {
                                                var xProp = ap.GetType().GetProperty("X");
                                                var yProp = ap.GetType().GetProperty("Y");
                                                if (xProp != null && yProp != null)
                                                {
                                                    alignX = Convert.ToDouble(xProp.GetValue(ap));
                                                    alignY = Convert.ToDouble(yProp.GetValue(ap));
                                                }
                                            }
                                        }
                                    }
                                    catch { }
                                }
                            }
                        }
                        else if (string.Equals(tipo, "MTEXT", StringComparison.OrdinalIgnoreCase))
                        {
                            if (!string.IsNullOrEmpty(handle) && mtextByHandle.TryGetValue(handle, out var m))
                            {
                                wrapWidth = m.RectangleWidth;
                                attachment = m.AttachmentPoint.ToString();
                            }
                        }
                        else if (string.Equals(tipo, "ATTRIB", StringComparison.OrdinalIgnoreCase))
                        {
                            if (!string.IsNullOrEmpty(handle) && attribByHandle.TryGetValue(handle, out var a))
                            {
                                try { widthFactor = a.WidthFactor; } catch { }
                                try { attachment = a.Alignment.ToString(); } catch { }

                                // Posizione world già buona in questa build
                                x = Math.Round(a.Position.X, 2);
                                y = Math.Round(a.Position.Y, 2);
                                alignX = x;
                                alignY = y;

                                // Rotazione/altezza effettive (ereditano dal blocco)
                                double aRot = 0.0; try { aRot = a.Rotation; } catch { }
                                var ins = a.Owner;
                                double insRot = 0.0; double sy = 1.0;
                                if (ins != null)
                                {
                                    try { insRot = ins.Rotation; } catch { }
                                    try { sy = ins.Scale.Y; } catch { }
                                }
                                rot = Math.Round((rot ?? 0.0) + aRot + insRot, 2);
                                if (hgt.HasValue) hgt = Math.Round(hgt.Value * sy, 2);
                            }
                        }

                        alignX ??= x;
                        alignY ??= y;

                        bool isM = string.Equals(tipo, "MTEXT", StringComparison.OrdinalIgnoreCase);
                        string norm = NormalizeDxfText(raw, isM);

                        int keyId = Db.GetOrCreateTextKey(defaultText: norm, normalized: norm);

                        Db.InsertOccurrence(
                            drawingId: drawingId,
                            entityHandle: handle,
                            entityType: tipo,
                            layerOriginal: layer,
                            posX: x, posY: y, posZ: 0,
                            rotation: rot, height: hgt, style: stile,
                            widthFactor: widthFactor, attachment: attachment, wrapWidth: wrapWidth,
                            originalRaw: raw, originalNormalized: norm,
                            keyId: keyId, blockName: block, attribTag: tag,
                            alignX: alignX, alignY: alignY
                        );

                        r.Cells["colKeyId"].Value = keyId;
                        occInserted++;
                    }

                    filesOk++;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, $"Errore import '{file}':\n{ex.Message}", "ERRORE",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            lblStatsImport.Text = $"Import selezione: file OK {filesOk}, occorrenze inserite {occInserted}.";
            LoadDrawingsGrid();
            LoadTranslationsGrid();
            MessageBox.Show(this,
                $"Import COMPLETATO.\nFile OK: {filesOk}\nOccorrenze inserite: {occInserted}",
                "COMPLETATO", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // =====================================================================
        //                           TAB: Traduzioni
        // =====================================================================

        private void LoadCulturesIntoCombos()
        {
            cmbCultureTranslations.Items.Clear();
            cmbCultureGenerate.Items.Clear();

            var dbPath = Db.GetDatabasePath();
            if (string.IsNullOrWhiteSpace(dbPath) || !File.Exists(dbPath)) return;

            var cultures = Db.GetCultures();
            foreach (var c in cultures)
            {
                cmbCultureTranslations.Items.Add(c);
                cmbCultureGenerate.Items.Add(c);
            }

            int idx = cmbCultureTranslations.Items.IndexOf("it-IT");
            cmbCultureTranslations.SelectedIndex = idx >= 0 ? idx : (cmbCultureTranslations.Items.Count > 0 ? 0 : -1);

            int idx2 = cmbCultureGenerate.Items.IndexOf("it-IT");
            cmbCultureGenerate.SelectedIndex = idx2 >= 0 ? idx2 : (cmbCultureGenerate.Items.Count > 0 ? 0 : -1);

            lblStatsTranslations.Text = cmbCultureTranslations.Items.Count > 0
                ? $"Lingue caricate: {cmbCultureTranslations.Items.Count}"
                : "Nessuna lingua caricata";
        }

        private void LoadTranslationsGrid()
        {
            gridTranslations.Rows.Clear();

            var culture = cmbCultureTranslations.SelectedItem?.ToString();
            if (string.IsNullOrWhiteSpace(culture)) return;

            var filter = txtFilter.Text ?? "";
            var rows = Db.GetTranslationsForCulture(culture, filter);

            int missing = 0, warn = 0;
            foreach (var r in rows)
            {
                gridTranslations.Rows.Add(
                    r.keyId, r.defaultText, r.translated, r.status,
                    r.warn == 1, r.notes
                );
                if (r.status == "Missing" || string.IsNullOrEmpty(r.translated)) missing++;
                if (r.warn == 1) warn++;
            }

            gridTranslations.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells);

            lblStatsTranslations.Text = $"Lingua {culture}: {rows.Count} voci (mancanti {missing}, warning {warn})";
        }

        private void OnExportMissingCsv()
        {
            var dbPath = Db.GetDatabasePath();
            if (string.IsNullOrWhiteSpace(dbPath) || !File.Exists(dbPath))
            {
                MessageBox.Show(this, "Seleziona il database (Scegli DB…).", "MANCANO PREREQUISITI",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var culture = cmbCultureTranslations.SelectedItem?.ToString();
            if (string.IsNullOrWhiteSpace(culture))
            {
                MessageBox.Show(this, "Seleziona una lingua.", "MANCANO PREREQUISITI",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var rows = Db.GetMissingKeysForCulture(culture);
            if (rows.Count == 0)
            {
                MessageBox.Show(this, $"Nessuna chiave mancante per {culture}.", "INFO",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using var sfd = new SaveFileDialog
            {
                Title = "Esporta CSV (mancanti)",
                Filter = "CSV (*.csv)|*.csv",
                FileName = $"translations_missing_{culture}_{DateTime.Now:yyyyMMdd}.csv",
                OverwritePrompt = true
            };
            if (sfd.ShowDialog(this) != DialogResult.OK) return;

            try
            {
                using var sw = new StreamWriter(sfd.FileName, false, System.Text.Encoding.UTF8);
                using var csv = new CsvWriter(sw, CultureInfo.InvariantCulture);
                csv.WriteField("KeyId"); csv.WriteField("DefaultText"); csv.WriteField("Culture"); csv.WriteField("Translated"); csv.WriteField("Status"); csv.NextRecord();
                foreach (var (keyId, defaultText) in rows)
                {
                    csv.WriteField(keyId);
                    csv.WriteField(defaultText);
                    csv.WriteField(culture);
                    csv.WriteField("");
                    csv.WriteField("");
                    csv.NextRecord();
                }
                csv.Flush(); sw.Flush();
                MessageBox.Show(this, $"Export COMPLETATO.\nRighe: {rows.Count}\n{sfd.FileName}",
                    "COMPLETATO", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"Errore export CSV:\n{ex.Message}", "ERRORE",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void OnImportCsv()
        {
            var dbPath = Db.GetDatabasePath();
            if (string.IsNullOrWhiteSpace(dbPath) || !File.Exists(dbPath))
            {
                MessageBox.Show(this, "Seleziona il database (Scegli DB…).", "MANCANO PREREQUISITI",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using var ofd = new OpenFileDialog
            {
                Title = "Importa CSV traduzioni",
                Filter = "CSV (*.csv)|*.csv|Tutti i file (*.*)|*.*",
                CheckFileExists = true,
                Multiselect = false
            };
            if (ofd.ShowDialog(this) != DialogResult.OK) return;

            int ok = 0, skip = 0, err = 0;
            try
            {
                using var sr = new StreamReader(ofd.FileName);
                using var csv = new CsvReader(sr, CultureInfo.InvariantCulture);

                while (csv.Read())
                {
                    try
                    {
                        int keyId = csv.GetField<int>("KeyId");
                        string culture = csv.TryGetField<string>("Culture", out var c) ? (c ?? "") : "";
                        string translated = csv.TryGetField<string>("Translated", out var t) ? (t ?? "") : "";
                        string status = csv.TryGetField<string>("Status", out var s) ? (s ?? "") : "";

                        if (string.IsNullOrWhiteSpace(culture))
                        {
                            culture = cmbCultureTranslations.SelectedItem?.ToString() ?? "";
                        }
                        if (string.IsNullOrWhiteSpace(culture))
                        {
                            skip++;
                            continue;
                        }

                        Db.UpsertTranslation(keyId, culture, translated, status, overwriteApproved: chkOverwriteApproved.Checked);
                        ok++;
                    }
                    catch
                    {
                        err++;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"Errore lettura CSV:\n{ex.Message}", "ERRORE",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            LoadTranslationsGrid();
            MessageBox.Show(this, $"Import CSV completato.\nOK: {ok}\nSaltati: {skip}\nErrori: {err}", "COMPLETATO",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void OnExportWholeLanguage()
        {
            var dbPath = Db.GetDatabasePath();
            if (string.IsNullOrWhiteSpace(dbPath) || !File.Exists(dbPath))
            {
                MessageBox.Show(this, "Seleziona il database (Scegli DB…).", "MANCANO PREREQUISITI",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            var culture = cmbCultureTranslations.SelectedItem?.ToString();
            if (string.IsNullOrWhiteSpace(culture))
            {
                MessageBox.Show(this, "Seleziona una lingua.", "MANCANO PREREQUISITI",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var rows = Db.GetAllForCultureExport(culture);
            if (rows.Count == 0)
            {
                MessageBox.Show(this, $"Nessuna chiave trovata per {culture}.", "INFO",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using var sfd = new SaveFileDialog
            {
                Title = "Esporta CSV (tutta la lingua)",
                Filter = "CSV (*.csv)|*.csv",
                FileName = $"translations_full_{culture}_{DateTime.Now:yyyyMMdd}.csv",
                OverwritePrompt = true
            };
            if (sfd.ShowDialog(this) != DialogResult.OK) return;

            try
            {
                using var sw = new StreamWriter(sfd.FileName, false, System.Text.Encoding.UTF8);
                using var csv = new CsvWriter(sw, CultureInfo.InvariantCulture);
                csv.WriteField("KeyId"); csv.WriteField("DefaultText"); csv.WriteField("Culture"); csv.WriteField("Translated"); csv.WriteField("Status"); csv.NextRecord();
                foreach (var r in rows)
                {
                    csv.WriteField(r.keyId);
                    csv.WriteField(r.defaultText);
                    csv.WriteField(r.culture);
                    csv.WriteField(r.translated);
                    csv.WriteField(r.status);
                    csv.NextRecord();
                }
                csv.Flush(); sw.Flush();
                MessageBox.Show(this, $"Export COMPLETATO.\nRighe: {rows.Count}\n{sfd.FileName}",
                    "COMPLETATO", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"Errore export CSV:\n{ex.Message}", "ERRORE",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void OnSetStatusSelected(string newStatus)
        {
            var culture = cmbCultureTranslations.SelectedItem?.ToString();
            if (string.IsNullOrWhiteSpace(culture))
            {
                MessageBox.Show(this, "Seleziona una lingua.", "MANCANO PREREQUISITI",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            int count = 0;
            foreach (DataGridViewRow r in gridTranslations.SelectedRows)
            {
                if (r.Cells["colKeyId2"].Value == null) continue;
                int keyId = Convert.ToInt32(r.Cells["colKeyId2"].Value);
                Db.UpdateStatus(keyId, culture, newStatus);
                count++;
            }

            LoadTranslationsGrid();
            MessageBox.Show(this, $"Stato impostato a {newStatus}: {count} righe", "COMPLETATO",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // =====================================================================
        //                           TAB: DXF Unico (IxMilia)
        // =====================================================================

        private void LoadDrawingsGrid()
        {
            gridDrawings.Rows.Clear();

            var dbPath = Db.GetDatabasePath();
            if (string.IsNullOrWhiteSpace(dbPath) || !File.Exists(dbPath)) return;

            var rows = Db.GetDrawings();
            foreach (var (id, path, importedAt) in rows)
                gridDrawings.Rows.Add(id, path, importedAt);
        }

        private void OnGenerateUnified()
        {
            // 1) Riga selezionata (o prima)
            DataGridViewRow? sel = gridDrawings.SelectedRows.Count > 0 ? gridDrawings.SelectedRows[0]
                                  : (gridDrawings.Rows.Count > 0 ? gridDrawings.Rows[0] : null);
            if (sel == null)
            {
                MessageBox.Show(this, "Nessun disegno in elenco. Importa prima un DXF.", "MANCANO PREREQUISITI",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            int drawingId = Convert.ToInt32(sel.Cells["colDrawingId"].Value);
            string sourcePath = sel.Cells["colPath"].Value?.ToString() ?? "";
            if (!File.Exists(sourcePath))
            {
                MessageBox.Show(this, "Il file sorgente non esiste più sul disco.", "ERRORE",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Conferma (prendiamo anche non-Approved se la spunta è disattivata)
            if (MessageBox.Show(this, "Hai controllato i Not Approved?", "Conferma",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No) return;

            // 2) File di uscita
            var defaultDir = @"C:\_ProjNoGit\_DXF_I18N\_DXF";
            try { Directory.CreateDirectory(defaultDir); } catch { }
            using var sfd = new SaveFileDialog
            {
                Title = "Salva DXF unico",
                Filter = "DXF (*.dxf)|*.dxf",
                InitialDirectory = defaultDir,
                FileName = Path.GetFileNameWithoutExtension(sourcePath) + "_multi.dxf",
                OverwritePrompt = true
            };
            if (sfd.ShowDialog(this) != DialogResult.OK) return;

            try
            {
                // 3) Carica il DXF sorgente (netDxf) e indicizza per handle
                var src = DxfDocument.Load(sourcePath);

                var textByHandle = src.Entities.Texts?
                    .Where(t => !string.IsNullOrEmpty(t.Handle))
                    .ToDictionary(t => t.Handle!, t => t)
                    ?? new Dictionary<string, Text>();

                var mtextByHandle = src.Entities.MTexts?
                    .Where(m => !string.IsNullOrEmpty(m.Handle))
                    .ToDictionary(m => m.Handle!, m => m)
                    ?? new Dictionary<string, MText>();

                // Attributi dentro gli INSERT: mappo handle -> Attribute
                var attribByHandle = new Dictionary<string, netDxf.Entities.Attribute>();
                if (src.Entities.Inserts != null)
                {
                    foreach (var ins in src.Entities.Inserts)
                    {
                        if (ins.Attributes == null) continue;
                        foreach (var a in ins.Attributes)
                        {
                            if (!string.IsNullOrEmpty(a.Handle))
                                attribByHandle[a.Handle] = a;
                        }
                    }
                }

                // 4) Assicura i layer di lingua
                var cultureLayers = Db.GetCultureLayers(); // (Code, LayerName)
                foreach (var (_, layerName) in cultureLayers)
                    if (!src.Layers.Contains(layerName))
                        src.Layers.Add(new Layer(layerName));

                // 5) Mappa traduzioni
                bool onlyApproved = chkOnlyApprovedGen.Checked;
                var cultureMaps = new Dictionary<string, Dictionary<int, string>>();
                foreach (var (code, _) in cultureLayers)
                    cultureMaps[code] = Db.GetTranslationsMap(code, onlyApproved);

                // 6) Occorrenze del disegno dal DB
                var occs = Db.GetOccurrencesForDrawing(drawingId);

                // 7) Overlay: per ogni occorrenza e lingua
                foreach (var occ in occs)
                {
                    foreach (var (code, layerName) in cultureLayers)
                    {
                        string textOut = cultureMaps[code].TryGetValue(occ.KeyId, out var tr) && !string.IsNullOrEmpty(tr)
                                         ? tr
                                         : occ.DefaultText;

                        var layerRef = src.Layers[layerName];

                        if (string.Equals(occ.EntityType, "TEXT", StringComparison.OrdinalIgnoreCase))
                        {
                            if (!string.IsNullOrEmpty(occ.EntityHandle) && textByHandle.TryGetValue(occ.EntityHandle, out var tOrig))
                            {
                                // Clona l’originale: geometria/stile/allineamento identici
                                var t = (Text)tOrig.Clone();
                                t.Value = textOut;
                                t.Layer = layerRef;
                                src.Entities.Add(t);
                            }
                            else
                            {
                                // Fallback (se handle mancante): crea TEXT con i valori salvati
                                var t = new Text(textOut, new netDxf.Vector2(occ.PosX, occ.PosY), occ.Height ?? 2.5)
                                {
                                    Rotation = occ.Rotation ?? 0.0,
                                    Layer = layerRef
                                };
                                if (!string.IsNullOrWhiteSpace(occ.Style) && src.TextStyles.Contains(occ.Style))
                                    t.Style = src.TextStyles[occ.Style];
                                if (!string.IsNullOrWhiteSpace(occ.Attachment) &&
                                    Enum.TryParse<TextAlignment>(occ.Attachment, out var ta))
                                    t.Alignment = ta;
                                if (occ.WidthFactor.HasValue && occ.WidthFactor.Value > 0)
                                    t.WidthFactor = occ.WidthFactor.Value;
                                src.Entities.Add(t);
                            }
                        }
                        else if (string.Equals(occ.EntityType, "MTEXT", StringComparison.OrdinalIgnoreCase))
                        {
                            if (!string.IsNullOrEmpty(occ.EntityHandle) && mtextByHandle.TryGetValue(occ.EntityHandle, out var mOrig))
                            {
                                var m = (MText)mOrig.Clone();
                                m.Value = (textOut ?? string.Empty).Replace("\r", "").Replace("\n", "\\P");
                                m.Layer = layerRef;
                                src.Entities.Add(m);
                            }
                            else
                            {
                                var m = new MText
                                {
                                    Value = (textOut ?? string.Empty).Replace("\r", "").Replace("\n", "\\P"),
                                    Position = new netDxf.Vector3(occ.PosX, occ.PosY, occ.PosZ),
                                    Height = occ.Height ?? 2.5,
                                    Rotation = occ.Rotation ?? 0.0,
                                    Layer = layerRef
                                };
                                if (!string.IsNullOrWhiteSpace(occ.Style) && src.TextStyles.Contains(occ.Style))
                                    m.Style = src.TextStyles[occ.Style];
                                if (!string.IsNullOrWhiteSpace(occ.Attachment) &&
                                    Enum.TryParse<MTextAttachmentPoint>(occ.Attachment, out var ap))
                                    m.AttachmentPoint = ap;
                                if (occ.WrapWidth.HasValue && occ.WrapWidth.Value > 0)
                                    m.RectangleWidth = occ.WrapWidth.Value;
                                src.Entities.Add(m);
                            }
                        }
                        else // ATTRIB: replica come TEXT con le stesse proprietà "world"
                        {
                            if (!string.IsNullOrEmpty(occ.EntityHandle) && attribByHandle.TryGetValue(occ.EntityHandle, out var a))
                            {
                                double h = 2.5; try { h = a.Height; } catch { }
                                double rot = 0.0; try { rot = a.Rotation; } catch { }
                                var t = new Text(textOut ?? string.Empty, new netDxf.Vector2(a.Position.X, a.Position.Y), h)
                                {
                                    Rotation = rot,
                                    Layer = layerRef
                                };
                                if (a.Style != null && src.TextStyles.Contains(a.Style.Name))
                                    t.Style = src.TextStyles[a.Style.Name];
                                try { t.Alignment = a.Alignment; } catch { }
                                try { t.WidthFactor = a.WidthFactor; } catch { }
                                src.Entities.Add(t);
                            }
                            else
                            {
                                var t = new Text(textOut ?? string.Empty, new netDxf.Vector2(occ.PosX, occ.PosY), occ.Height ?? 2.5)
                                {
                                    Rotation = occ.Rotation ?? 0.0,
                                    Layer = layerRef
                                };
                                if (!string.IsNullOrWhiteSpace(occ.Style) && src.TextStyles.Contains(occ.Style))
                                    t.Style = src.TextStyles[occ.Style];
                                if (!string.IsNullOrWhiteSpace(occ.Attachment) &&
                                    Enum.TryParse<TextAlignment>(occ.Attachment, out var ta))
                                    t.Alignment = ta;
                                if (occ.WidthFactor.HasValue && occ.WidthFactor.Value > 0)
                                    t.WidthFactor = occ.WidthFactor.Value;
                                src.Entities.Add(t);
                            }
                        }
                    }
                }

                // 8) Salva
                src.Save(sfd.FileName);
                lblGenerateInfo.Text = $"DXF unico creato: {sfd.FileName}";
                MessageBox.Show(this, "Generazione COMPLETATA.", "OK",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"Errore generazione DXF:\n{ex.Message}", "ERRORE",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
