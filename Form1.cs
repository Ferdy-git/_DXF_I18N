using CsvHelper;
using CsvHelper.Configuration;
// ==== LIBRERIA PER L'ANTEPRIMA/IMPORT (come prima) ====
using netDxf;
using netDxf.Entities;
using netDxf.Tables;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection; // reflection per AlignmentPoint netDxf
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
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

            // === DXF Unico: mostra i controlli lingua e aggiungi il bottone per singola lingua ===
            try
            {
                lblLangOut.Visible = true;           // era false
                cmbCultureGenerate.Visible = true;    // era false

                // Crea/riusa un bottone accanto a "Genera DXF unico"
                var found = this.Controls.Find("btnGenerateSingle", true);
                Button btnGenerateSingle;
                if (found.Length > 0)
                {
                    btnGenerateSingle = (Button)found[0];
                }
                else
                {
                    var parent = btnGenerateUnified.Parent ?? this;
                    btnGenerateSingle = new Button
                    {
                        Name = "btnGenerateSingle",
                        Text = "Genera DXF (lingua)…",
                        Size = btnGenerateUnified.Size,
                        Location = new System.Drawing.Point(btnGenerateUnified.Right + 10, btnGenerateUnified.Top),
                        Anchor = btnGenerateUnified.Anchor,
                        TabIndex = btnGenerateUnified.TabIndex + 1,
                    };
                    parent.Controls.Add(btnGenerateSingle);
                }

                // Hook all'handler esistente
                btnGenerateSingle.Click += (s, e) => OnGenerateSingle();
            }
            catch { /* se i controlli non ci sono in questa vista, ignora */ }

            // ----- Importa DXF -----
            btnSelectDb.Click += (s, e) => OnSelectDb();
            btnPickDxf.Click += (s, e) => OnPickDxf();
            btnPreview.Click += (s, e) => OnPreview();
            btnApplyFilters.Click += (s, e) => OnApplyFilters();
            btnSelectAllPreview.Click += (s, e) => OnSelectAllPreview(true);
            btnDeselectAllPreview.Click += (s, e) => OnSelectAllPreview(false);
            btnImport.Click += (s, e) => OnImport();
            btnCleanDb.Click += OnCleanDb_Click;
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

        // ----------------- NUOVI HELPER per entità dentro i blocchi -----------------

        // Trasforma un punto locale del blocco nelle coordinate world dell'INSERT
        // Costruisce la matrice di trasformazione dell'INSERT (Scala → Rotazione Z → Traslazione)
        //private static netDxf.Matrix3 InsertTransform(netDxf.Entities.Insert ins)
        //{
        //    var sx = ins.Scale.X == 0 ? 1.0 : ins.Scale.X;
        //    var sy = ins.Scale.Y == 0 ? 1.0 : ins.Scale.Y;
        //    var s = netDxf.Matrix3.Scale(sx, sy, 1.0);

        //    var r = netDxf.Matrix3.RotationZ(netDxf.MathHelper.DegToRad(ins.Rotation));
        //    var t = netDxf.Matrix3.Translation(ins.Position.X, ins.Position.Y, 0.0);

        //    // ordine: scala → ruota → trasla
        //    return t * r * s;
        //}

        // Clona un'entità testuale del blocco e la trasforma nello spazio mondo dell'INSERT
        // Restituisce null se non trovata.
        //private static netDxf.Entities.EntityObject? CloneAndTransformBlockText(
        //    netDxf.DxfDocument src, string insertHandle, string innerHandle, string textOut, netDxf.Tables.Layer layerRef)
        //{
        //    var ins = src.Entities.Inserts?.FirstOrDefault(i => string.Equals(i.Handle, insertHandle, StringComparison.OrdinalIgnoreCase));
        //    if (ins == null || ins.Block == null) return null;

        //    // cerca prima MTEXT poi TEXT
        //    var mInner = ins.Block.Entities.OfType<netDxf.Entities.MText>()
        //                  .FirstOrDefault(e => string.Equals(e.Handle, innerHandle, StringComparison.OrdinalIgnoreCase));
        //    if (mInner != null)
        //    {
        //        var clone = (netDxf.Entities.MText)mInner.Clone();
        //        clone.Layer = layerRef;
        //        clone.Value = (textOut ?? "").Replace("\r", "").Replace("\n", "\\P");
        //        // importantissimo: trasformazione completa dell'INSERT
        //        clone.TransformBy(InsertTransform(ins));
        //        return clone;
        //    }

        //    var tInner = ins.Block.Entities.OfType<netDxf.Entities.Text>()
        //                  .FirstOrDefault(e => string.Equals(e.Handle, innerHandle, StringComparison.OrdinalIgnoreCase));
        //    if (tInner != null)
        //    {
        //        var clone = (netDxf.Entities.Text)tInner.Clone();
        //        clone.Layer = layerRef;
        //        clone.Value = textOut ?? "";
        //        clone.TransformBy(InsertTransform(ins));
        //        return clone;
        //    }

        //    return null;
        //}
        // Clona l'entità (MTEXT/TEXT) dal BlockDefinition e la porta nello spazio world dell'INSERT
        // Clona l'entità (MTEXT/TEXT) dal BlockDefinition e la porta nello spazio world dell'INSERT
        private static netDxf.Entities.EntityObject? CloneAndTransformBlockText(
            netDxf.DxfDocument src, string insertHandle, string innerHandle, string textOut, netDxf.Tables.Layer layerRef)
        {
            var ins = src.Entities.Inserts?.FirstOrDefault(i => string.Equals(i.Handle, insertHandle, StringComparison.OrdinalIgnoreCase));
            if (ins == null || ins.Block == null) return null;

            // ---------- 1) MTEXT nel blocco ----------
            var mInner = ins.Block.Entities.OfType<netDxf.Entities.MText>()
                .FirstOrDefault(e => string.Equals(e.Handle, innerHandle, StringComparison.OrdinalIgnoreCase));
            if (mInner != null)
            {
                var clone = (netDxf.Entities.MText)mInner.Clone();
                clone.Layer = layerRef;
                clone.Value = (textOut ?? string.Empty).Replace("\r", "").Replace("\n", "\\P");

                // posizione world (Position è Vector3)
                var wp = TransformPoint2D(new netDxf.Vector2(clone.Position.X, clone.Position.Y), ins);
                clone.Position = new netDxf.Vector3(wp.X, wp.Y, clone.Position.Z);

                // rotazione e scala (altezza su Y, larghezza su X)
                clone.Rotation = clone.Rotation + ins.Rotation;
                double sy = ins.Scale.Y == 0 ? 1.0 : ins.Scale.Y;
                double sx = ins.Scale.X == 0 ? 1.0 : ins.Scale.X;
                clone.Height *= sy;
                if (clone.RectangleWidth > 0) clone.RectangleWidth *= sx;

                return clone;
            }

            // ---------- 2) TEXT nel blocco ----------
            var tInner = ins.Block.Entities.OfType<netDxf.Entities.Text>()
                .FirstOrDefault(e => string.Equals(e.Handle, innerHandle, StringComparison.OrdinalIgnoreCase));
            if (tInner != null)
            {
                var clone = (netDxf.Entities.Text)tInner.Clone();
                clone.Layer = layerRef;
                clone.Value = textOut ?? string.Empty;

                // --- Position: può essere Vector2 o Vector3 a seconda della versione ---
                var posProp = typeof(netDxf.Entities.Text).GetProperty("Position", BindingFlags.Public | BindingFlags.Instance);
                if (posProp != null)
                {
                    var posObj = posProp.GetValue(clone);
                    if (posObj is netDxf.Vector3 p3)
                    {
                        var wp = TransformPoint2D(new netDxf.Vector2(p3.X, p3.Y), ins);
                        posProp.SetValue(clone, new netDxf.Vector3(wp.X, wp.Y, p3.Z));
                    }
                    else if (posObj is netDxf.Vector2 p2)
                    {
                        var wp = TransformPoint2D(p2, ins);
                        posProp.SetValue(clone, new netDxf.Vector2(wp.X, wp.Y));
                    }
                }

                // rotazione/altezza
                clone.Rotation = clone.Rotation + ins.Rotation;
                double sy = ins.Scale.Y == 0 ? 1.0 : ins.Scale.Y;
                clone.Height *= sy;

                // --- AlignmentPoint: può essere Vector2 o Vector3 ---
                if (clone.Alignment != netDxf.Entities.TextAlignment.BaselineLeft)
                {
                    var apProp = typeof(netDxf.Entities.Text).GetProperty("AlignmentPoint", BindingFlags.Public | BindingFlags.Instance);
                    if (apProp != null)
                    {
                        var apObj = apProp.GetValue(clone);
                        if (apObj is netDxf.Vector3 ap3)
                        {
                            var apW = TransformPoint2D(new netDxf.Vector2(ap3.X, ap3.Y), ins);
                            apProp.SetValue(clone, new netDxf.Vector3(apW.X, apW.Y, ap3.Z));
                        }
                        else if (apObj is netDxf.Vector2 ap2)
                        {
                            var apW = TransformPoint2D(ap2, ins);
                            apProp.SetValue(clone, new netDxf.Vector2(apW.X, apW.Y));
                        }
                    }
                }

                return clone;
            }

            return null;
        }


        private static void EnsureTextAlignmentPoint(netDxf.Entities.Text t)
        {
            if (t == null) return;
            if (t.Alignment == netDxf.Entities.TextAlignment.BaselineLeft) return;

            // imposta anche i codici 11/21 = alignment point
            var apProp = t.GetType().GetProperty("AlignmentPoint", System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance);
            if (apProp != null)
            {
                var p = t.Position; // abbiamo già messo l'anchor qui
                apProp.SetValue(t, new netDxf.Vector2(p.X, p.Y));
            }

            // opzionale: normal sempre Z+
            try { t.Normal = new netDxf.Vector3(0, 0, 1); } catch { }
        }

        private static netDxf.Vector2 TransformPoint2D(netDxf.Vector2 p, netDxf.Entities.Insert ins)
        {
            double sx = ins.Scale.X, sy = ins.Scale.Y;
            double ang = (ins.Rotation) * Math.PI / 180.0;

            double x = p.X * sx;
            double y = p.Y * sy;

            double xr = x * Math.Cos(ang) - y * Math.Sin(ang);
            double yr = x * Math.Sin(ang) + y * Math.Cos(ang);

            return new netDxf.Vector2(xr + ins.Position.X, yr + ins.Position.Y);
        }

        // Layer "effettivo": se l'entità del blocco è su layer "0", eredita il layer dell'INSERT
        private static string EffectiveLayer(netDxf.Tables.Layer? entLayer, netDxf.Entities.Insert ins)
        {
            var l = entLayer?.Name ?? "";
            if (string.Equals(l, "0", StringComparison.OrdinalIgnoreCase) || string.IsNullOrWhiteSpace(l))
                l = ins.Layer?.Name ?? l;
            return l ?? "";
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
        private void OnCleanDb_Click(object? sender, EventArgs e) => OnCleanDb();

        /// <summary>
        /// Pulisce completamente i dati (Translations, Occurrences, Drawings, TextKeys).
        /// Mantiene le culture. Offre un backup del DB prima di procedere.
        /// </summary>
        private void OnCleanDb()
        {
            var dbPath = Db.GetDatabasePath();
            if (string.IsNullOrWhiteSpace(dbPath) || !File.Exists(dbPath))
            {
                MessageBox.Show(this, "Seleziona prima un database (Scegli DB…).", "MANCANO PREREQUISITI",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var res = MessageBox.Show(this,
                "Cancella TUTTI i dati: chiavi, occorrenze, disegni e traduzioni.\n" +
                "Le lingue (Cultures) restano.\n\nVuoi fare prima un backup del DB?",
                "Clean DB — Conferma",
                MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3);

            if (res == DialogResult.Cancel) return;

            if (res == DialogResult.Yes)
            {
                using var sfd = new SaveFileDialog
                {
                    Title = "Salva backup del database",
                    Filter = "SQLite DB (*.sqlite;*.db;*.db3)|*.sqlite;*.db;*.db3|Tutti i file (*.*)|*.*",
                    FileName = Path.GetFileNameWithoutExtension(dbPath) + $"_backup_{DateTime.Now:yyyyMMdd_HHmmss}.sqlite",
                    OverwritePrompt = true
                };
                if (sfd.ShowDialog(this) == DialogResult.OK)
                {
                    try { Db.BackupDatabase(sfd.FileName); }
                    catch (Exception ex)
                    {
                        MessageBox.Show(this, $"Backup non riuscito:\n{ex.Message}", "ERRORE",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
                else
                {
                    if (MessageBox.Show(this, "Backup annullato. Procedere comunque alla pulizia?",
                        "Conferma", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
                }
            }

            try
            {
                Db.CleanAllData(keepCultures: true);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"Errore durante la pulizia:\n{ex.Message}", "ERRORE",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Refresh UI
            gridOccurrences.Rows.Clear();
            gridDrawings.Rows.Clear();
            gridTranslations.Rows.Clear();
            lblStatsImport.Text = "DB pulito.";
            lblStatsTranslations.Text = "DB pulito.";
            LoadDrawingsGrid();
            LoadTranslationsGrid();

            MessageBox.Show(this, "Pulizia completata.\n(Cultures mantenute. File compattato.)", "COMPLETATO",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
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

                    // ---- TEXT top-level ----
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

                    // ---- MTEXT top-level ----
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

                    // ==== NOVITÀ: TEXT/MTEXT DENTRO I BLOCCHI (INSERT) ====
                    if ((includeText || includeMText) && dxf.Entities?.Inserts != null)
                    {
                        foreach (var ins in dxf.Entities.Inserts)
                        {
                            var block = ins.Block;
                            if (block == null) continue;

                            foreach (var ent in block.Entities)
                            {
                                if (includeText && ent is netDxf.Entities.Text tBlk)
                                {
                                    var wp = TransformPoint2D(new netDxf.Vector2(tBlk.Position.X, tBlk.Position.Y), ins);
                                    double wh = (tBlk.Height) * (ins.Scale.Y);
                                    double wrot = (tBlk.Rotation) + (ins.Rotation);

                                    string layerEff = EffectiveLayer(tBlk.Layer, ins);
                                    if (selectedLayers.Count > 0 && !selectedLayers.Contains(layerEff)) continue;

                                    string raw = tBlk.Value ?? "";
                                    bool include = PassesFilters(NormalizeDxfText(raw, false));
                                    string handle = $"{ins.Handle}|{tBlk.Handle}"; // handle composito

                                    AddPreviewRow(include, file, layerEff, "TEXT", raw,
                                        Math.Round(wp.X, 2), Math.Round(wp.Y, 2),
                                        Math.Round(wrot, 2), Math.Round(wh, 2),
                                        tBlk.Style?.Name, handle, null,
                                        block.Name, null);

                                    totalOcc++; if (include) totalShown++;
                                }
                                else if (includeMText && ent is netDxf.Entities.MText mBlk)
                                {
                                    var wp = TransformPoint2D(new netDxf.Vector2(mBlk.Position.X, mBlk.Position.Y), ins);
                                    double wh = (mBlk.Height) * (ins.Scale.Y);
                                    double wrot = (mBlk.Rotation) + (ins.Rotation);

                                    string layerEff = EffectiveLayer(mBlk.Layer, ins);
                                    if (selectedLayers.Count > 0 && !selectedLayers.Contains(layerEff)) continue;

                                    string raw = mBlk.Value ?? "";
                                    bool include = PassesFilters(NormalizeDxfText(raw, true));
                                    string handle = $"{ins.Handle}|{mBlk.Handle}"; // handle composito

                                    AddPreviewRow(include, file, layerEff, "MTEXT", raw,
                                        Math.Round(wp.X, 2), Math.Round(wp.Y, 2),
                                        Math.Round(wrot, 2), Math.Round(wh, 2),
                                        mBlk.Style?.Name, handle, null,
                                        block.Name, null);

                                    totalOcc++; if (include) totalShown++;
                                }
                            }
                        }
                    }

                    // ---- ATTRIB negli INSERT ----
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

                        // NEW: dati dell'INSERT (blocchi) — default null
                        string? ownerHandle = null;
                        double? ix = null, iy = null;
                        double? irot = null, sx = null, sy = null;

                        // Se proviene da un blocco (handle "INS|ENT"), salva parametri INSERT
                        // e, per MTEXT/TEXT, recupera anche le proprietà dall'entità interna al blocco
                        if ((string.Equals(tipo, "TEXT", StringComparison.OrdinalIgnoreCase) ||
                             string.Equals(tipo, "MTEXT", StringComparison.OrdinalIgnoreCase)) &&
                            handle.Contains("|"))
                        {
                            var parts = handle.Split('|');
                            if (parts.Length == 2)
                            {
                                string insHandle = parts[0];
                                string entHandle = parts[1];
                                ownerHandle = insHandle;

                                var ins = dxf.Entities.Inserts?.FirstOrDefault(i =>
                                    string.Equals(i.Handle, insHandle, StringComparison.OrdinalIgnoreCase));
                                if (ins != null)
                                {
                                    try { ix = ins.Position.X; } catch { }
                                    try { iy = ins.Position.Y; } catch { }
                                    try { irot = ins.Rotation; } catch { }
                                    try { sx = ins.Scale.X; } catch { }
                                    try { sy = ins.Scale.Y; } catch { }

                                    // >>> NOVITÀ: recupera l'entità interna dal BlockDefinition
                                    var blk = ins.Block;
                                    if (blk != null)
                                    {
                                        // MTEXT nel blocco
                                        if (string.Equals(tipo, "MTEXT", StringComparison.OrdinalIgnoreCase))
                                        {
                                            var mInner = blk.Entities.OfType<netDxf.Entities.MText>()
                                                .FirstOrDefault(e => string.Equals(e.Handle, entHandle, StringComparison.OrdinalIgnoreCase));
                                            if (mInner != null)
                                            {
                                                // Copiamo info che servono per replicare il wrapping
                                                try { wrapWidth = mInner.RectangleWidth; } catch { }
                                                try { attachment = mInner.AttachmentPoint.ToString(); } catch { }
                                                if (string.IsNullOrWhiteSpace(stile))
                                                    try { stile = mInner.Style?.Name; } catch { }
                                            }
                                        }
                                        // TEXT nel blocco (per completezza: allineamento/pitch)
                                        else if (string.Equals(tipo, "TEXT", StringComparison.OrdinalIgnoreCase))
                                        {
                                            var tInner = blk.Entities.OfType<netDxf.Entities.Text>()
                                                .FirstOrDefault(e => string.Equals(e.Handle, entHandle, StringComparison.OrdinalIgnoreCase));
                                            if (tInner != null)
                                            {
                                                try { widthFactor = tInner.WidthFactor; } catch { }
                                                try
                                                {
                                                    var alignStr = tInner.Alignment.ToString();
                                                    attachment = alignStr;
                                                    if (!string.Equals(alignStr, "BaselineLeft", StringComparison.OrdinalIgnoreCase))
                                                    {
                                                        // se hai bisogno del punto di allineamento, puoi calcolarlo qui
                                                        alignX ??= x; alignY ??= y;
                                                    }
                                                }
                                                catch { }
                                                if (string.IsNullOrWhiteSpace(stile))
                                                    try { stile = tInner.Style?.Name; } catch { }
                                            }
                                        }
                                    }
                                }
                            }
                        }


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

                                // Posizione "world"
                                x = Math.Round(a.Position.X, 2);
                                y = Math.Round(a.Position.Y, 2);
                                alignX = x;
                                alignY = y;

                                // Dati dell'INSERT (blocco)
                                var ins = a.Owner;
                                if (ins != null)
                                {
                                    try { ownerHandle = ins.Handle; } catch { }
                                    try { ix = ins.Position.X; } catch { }
                                    try { iy = ins.Position.Y; } catch { }
                                    try { irot = ins.Rotation; } catch { }
                                    try { sx = ins.Scale.X; } catch { }
                                    try { sy = ins.Scale.Y; } catch { }
                                }

                                // Rotazione/altezza effettive (ereditano dal blocco)
                                double aRot = 0.0; try { aRot = a.Rotation; } catch { }
                                double insRot = irot ?? 0.0;
                                double scaleY = sy ?? 1.0;

                                rot = Math.Round((rot ?? 0.0) + aRot + insRot, 2);
                                if (hgt.HasValue) hgt = Math.Round(hgt.Value * scaleY, 2);
                            }
                        }

                        // ----- COMUNE: normalizzazione testo + chiave + INSERT nel DB -----
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
                            alignX: alignX, alignY: alignY,
                            // NEW: dati dell'INSERT (blocco)
                            ownerInsertHandle: ownerHandle,
                            insertX: ix, insertY: iy,
                            insertRotation: irot,
                            insertScaleX: sx, insertScaleY: sy
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

        /// <summary>
        /// Importa un CSV di traduzioni (UTF-8), rileva automaticamente il separatore (, oppure ;) e
        /// aggiorna/crea righe in Translations. Rispetta "Consenti sovrascrittura Approved".
        /// </summary>
        // Restituisce il primo campo disponibile tra più possibili nomi di colonna (case-insensitive).
        private static string GetFieldAny(CsvReader csv, params string[] names)
        {
            foreach (var n in names)
            {
                if (csv.TryGetField<string>(n, out var v) && v != null)
                    return v;
                // prova anche varianti senza spazi / case-insensitive
                var noSpace = n.Replace(" ", "");
                if (csv.TryGetField<string>(noSpace, out v) && v != null)
                    return v;
            }
            return "";
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
            string? firstErr = null;

            try
            {
                // Heuristica separatore: se header ha ';' e non ',', usa ';'
                string? firstLine = null;
                try { firstLine = System.IO.File.ReadLines(ofd.FileName).FirstOrDefault(); } catch { }
                bool useSemicolon = firstLine != null && firstLine.IndexOf(';') >= 0 && firstLine.IndexOf(',') < 0;

                var cfg = new CsvHelper.Configuration.CsvConfiguration(System.Globalization.CultureInfo.InvariantCulture)
                {
                    Delimiter = useSemicolon ? ";" : ",",
                    IgnoreBlankLines = true,
                    TrimOptions = CsvHelper.Configuration.TrimOptions.Trim,
                    BadDataFound = null,
                    MissingFieldFound = null,
                };

                // UTF-8 + BOM detection (per caratteri speciali)
                using var sr = new System.IO.StreamReader(
                    ofd.FileName,
                    new System.Text.UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
                    detectEncodingFromByteOrderMarks: true
                );
                using var csv = new CsvHelper.CsvReader(sr, cfg);

                // Leggi intestazione
                if (!csv.Read()) throw new Exception("CSV vuoto.");
                csv.ReadHeader();

                while (csv.Read())
                {
                    try
                    {
                        // Requisito minimo: KeyId + Translated
                        if (!csv.TryGetField<int>("KeyId", out var keyId))
                        {
                            err++; if (firstErr == null) firstErr = "Colonna 'KeyId' assente o non numerica.";
                            continue;
                        }

                        // Translated: accetta sia "Translated" che "Translation"
                        string translated = GetFieldAny(csv, "Translated", "Translation", "TranslatedText");

                        // Culture: accetta "Culture", "CultureCode", "Culture code"; se manca usa la combo
                        string culture = GetFieldAny(csv, "Culture", "CultureCode", "Culture code");
                        if (string.IsNullOrWhiteSpace(culture))
                            culture = cmbCultureTranslations.SelectedItem?.ToString() ?? "";

                        // Status è facoltativo (default: Draft gestito in Upsert o lasciato vuoto)
                        string status = GetFieldAny(csv, "Status", "State");

                        if (string.IsNullOrWhiteSpace(culture))
                        {
                            // Non sappiamo dove importare → salto
                            skip++;
                            continue;
                        }

                        Db.UpsertTranslation(keyId, culture, translated, status, overwriteApproved: chkOverwriteApproved.Checked);
                        ok++;
                    }
                    catch (Exception rowEx)
                    {
                        err++; if (firstErr == null) firstErr = rowEx.Message;
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

            var extra = err > 0 && !string.IsNullOrEmpty(firstErr) ? $"\nPrimo errore: {firstErr}" : "";
            MessageBox.Show(this, $"Import CSV completato.\nOK: {ok}\nSaltati: {skip}\nErrori: {err}{extra}", "COMPLETATO",
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

        /// <summary>
        /// Genera un DXF "multi-lingua": duplica per ogni lingua i testi (TEXT/MTEXT/ATTRIB) su layer dedicati.
        /// Mantiene la geometria originale e aggiunge solo gli overlay di testo per ciascuna lingua attiva.
        /// </summary>
        private void OnGenerateUnified()
        {
            // [BEGIN] OnGenerateUnified
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
                        // === NEW: gestione occorrenze da blocco con clone+transform
                        if (!string.IsNullOrEmpty(occ.EntityHandle) && occ.EntityHandle.Contains("|"))
                        {
                            var parts = occ.EntityHandle.Split('|');
                            if (parts.Length == 2)
                            {
                                var cloned = CloneAndTransformBlockText(src, parts[0], parts[1], textOut ?? "", layerRef);
                                if (cloned != null)
                                {
                                    src.Entities.Add(cloned);
                                    continue;
                                }
                            }
                        }


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
                                EnsureTextAlignmentPoint(t);
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
                        else // ATTRIB
                        {
                            // Dati "world" già salvati in import
                            double px = occ.PosX, py = occ.PosY;
                            double ax = occ.AlignX ?? px, ay = occ.AlignY ?? py;

                            double outHeight = occ.Height ?? 2.5;
                            double outRotation = occ.Rotation ?? 0.0;

                            string? styleName = occ.Style;
                            double? widthFactorOut = occ.WidthFactor;

                            // Decodifica allineamento dal DB
                            TextAlignment? alignOut = null;
                            if (!string.IsNullOrWhiteSpace(occ.Attachment) &&
                                Enum.TryParse<TextAlignment>(occ.Attachment, out var ta))
                                alignOut = ta;

                            // Se non è BaselineLeft, usa il punto di allineamento come Position
                            var anchor = (alignOut.HasValue && alignOut.Value != TextAlignment.BaselineLeft)
                                         ? new netDxf.Vector2(ax, ay)
                                         : new netDxf.Vector2(px, py);

                            var t = new Text(textOut ?? string.Empty, anchor, outHeight)
                            {
                                Rotation = outRotation,
                                Layer = layerRef
                            };

                            if (!string.IsNullOrWhiteSpace(styleName) && src.TextStyles.Contains(styleName))
                                t.Style = src.TextStyles[styleName];

                            if (alignOut.HasValue)
                                t.Alignment = alignOut.Value;   // niente AlignmentPoint: questa versione non lo espone

                            if (widthFactorOut.HasValue && widthFactorOut.Value > 0)
                                t.WidthFactor = widthFactorOut.Value;
                            EnsureTextAlignmentPoint(t);
                            src.Entities.Add(t);
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

            // [END] OnGenerateUnified
        }

        /// <summary>
        /// Genera un DXF *solo* per la lingua selezionata in combo: crea overlay su un layer di lingua
        /// e salva con suffisso _<BCP47>.dxf (es. _pl-PL.dxf). Rispetta l'opzione "solo Approved".
        /// </summary>
        private void OnGenerateSingle()
        {
            // [BEGIN] OnGenerateSingle
            var culture = cmbCultureGenerate.SelectedItem?.ToString();
            if (string.IsNullOrWhiteSpace(culture))
            {
                MessageBox.Show(this, "Seleziona una lingua nella combo.", "MANCANO PREREQUISITI",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

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

            if (MessageBox.Show(this, "Hai controllato i Not Approved?", "Conferma",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No) return;

            var defaultDir = @"C:\_ProjNoGit\_DXF_I18N\_DXF";
            try { Directory.CreateDirectory(defaultDir); } catch { }
            using var sfd = new SaveFileDialog
            {
                Title = $"Salva DXF ({culture})",
                Filter = "DXF (*.dxf)|*.dxf",
                InitialDirectory = defaultDir,
                FileName = Path.GetFileNameWithoutExtension(sourcePath) + $"_{culture}.dxf",
                OverwritePrompt = true
            };
            if (sfd.ShowDialog(this) != DialogResult.OK) return;

            try
            {
                var src = netDxf.DxfDocument.Load(sourcePath);

                var textByHandle = src.Entities.Texts?
                    .Where(t => !string.IsNullOrEmpty(t.Handle))
                    .ToDictionary(t => t.Handle!, t => t)
                    ?? new Dictionary<string, netDxf.Entities.Text>();

                var mtextByHandle = src.Entities.MTexts?
                    .Where(m => !string.IsNullOrEmpty(m.Handle))
                    .ToDictionary(m => m.Handle!, m => m)
                    ?? new Dictionary<string, netDxf.Entities.MText>();

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

                var cultureLayers = Db.GetCultureLayers(); // (CultureCode, LayerName)
                string layerName = cultureLayers
                    .FirstOrDefault(x => string.Equals(x.CultureCode, culture, StringComparison.OrdinalIgnoreCase))
                    .LayerName ?? culture;

                if (!src.Layers.Contains(layerName))
                    src.Layers.Add(new netDxf.Tables.Layer(layerName));
                var layerRef = src.Layers[layerName];

                bool onlyApproved = chkOnlyApprovedGen.Checked;
                var map = Db.GetTranslationsMap(culture, onlyApproved);
                var occs = Db.GetOccurrencesForDrawing(drawingId);

                foreach (var occ in occs)
                {
                    string textOut = map.TryGetValue(occ.KeyId, out var tr) && !string.IsNullOrEmpty(tr)
                                     ? tr
                                     : occ.DefaultText;
                    // === NEW: se l'occorrenza proviene da un blocco, prova "clone + transform"
                    if (!string.IsNullOrEmpty(occ.EntityHandle) && occ.EntityHandle.Contains("|"))
                    {
                        var parts = occ.EntityHandle.Split('|');       // parts[0] = HANDLE INSERT, parts[1] = HANDLE entità nel blocco
                        if (parts.Length == 2)
                        {
                            var cloned = CloneAndTransformBlockText(src, parts[0], parts[1], textOut ?? "", layerRef);
                            if (cloned != null)
                            {
                                src.Entities.Add(cloned);
                                continue; // passa alla prossima occorrenza: abbiamo già inserito il testo
                            }
                        }
                    }

                    if (string.Equals(occ.EntityType, "TEXT", StringComparison.OrdinalIgnoreCase))
                    {
                        if (!string.IsNullOrEmpty(occ.EntityHandle) && textByHandle.TryGetValue(occ.EntityHandle, out var tOrig))
                        {
                            var t = (netDxf.Entities.Text)tOrig.Clone();
                            t.Value = textOut;
                            t.Layer = layerRef;
                            src.Entities.Add(t);
                        }
                        else
                        {
                            var anchor = new netDxf.Vector2(occ.PosX, occ.PosY);

                            if (!string.IsNullOrWhiteSpace(occ.Attachment) &&
                                Enum.TryParse<netDxf.Entities.TextAlignment>(occ.Attachment, out var ta) &&
                                ta != netDxf.Entities.TextAlignment.BaselineLeft &&
                                occ.AlignX.HasValue && occ.AlignY.HasValue)
                            {
                                anchor = new netDxf.Vector2(occ.AlignX.Value, occ.AlignY.Value);
                            }

                            var t = new netDxf.Entities.Text(textOut, anchor, occ.Height ?? 2.5)
                            {
                                Rotation = occ.Rotation ?? 0.0,
                                Layer = layerRef
                            };
                            if (!string.IsNullOrWhiteSpace(occ.Style) && src.TextStyles.Contains(occ.Style))
                                t.Style = src.TextStyles[occ.Style];
                            if (!string.IsNullOrWhiteSpace(occ.Attachment) &&
                                Enum.TryParse<netDxf.Entities.TextAlignment>(occ.Attachment, out var ta2))
                                t.Alignment = ta2;
                            if (occ.WidthFactor.HasValue && occ.WidthFactor.Value > 0)
                                t.WidthFactor = occ.WidthFactor.Value;
                            EnsureTextAlignmentPoint(t);
                            src.Entities.Add(t);
                        }
                    }
                    else if (string.Equals(occ.EntityType, "MTEXT", StringComparison.OrdinalIgnoreCase))
                    {
                        if (!string.IsNullOrEmpty(occ.EntityHandle) && mtextByHandle.TryGetValue(occ.EntityHandle, out var mOrig))
                        {
                            var m = (netDxf.Entities.MText)mOrig.Clone();
                            m.Value = (textOut ?? string.Empty).Replace("\r", "").Replace("\n", "\\P");
                            m.Layer = layerRef;
                            src.Entities.Add(m);
                        }
                        else
                        {
                            var m = new netDxf.Entities.MText
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
                                Enum.TryParse<netDxf.Entities.MTextAttachmentPoint>(occ.Attachment, out var ap))
                                m.AttachmentPoint = ap;
                            if (occ.WrapWidth.HasValue && occ.WrapWidth.Value > 0)
                                m.RectangleWidth = occ.WrapWidth.Value;
                            src.Entities.Add(m);
                        }
                    }
                    else // ATTRIB → TEXT coerente
                    {
                        double px = occ.PosX, py = occ.PosY;
                        double ax = occ.AlignX ?? px, ay = occ.AlignY ?? py;

                        double outHeight = occ.Height ?? 2.5;
                        double outRotation = occ.Rotation ?? 0.0;

                        netDxf.Entities.TextAlignment? alignOut = null;
                        if (!string.IsNullOrWhiteSpace(occ.Attachment) &&
                            Enum.TryParse<netDxf.Entities.TextAlignment>(occ.Attachment, out var ta))
                            alignOut = ta;

                        var anchor = (alignOut.HasValue && alignOut.Value != netDxf.Entities.TextAlignment.BaselineLeft)
                                     ? new netDxf.Vector2(ax, ay)
                                     : new netDxf.Vector2(px, py);

                        var t = new netDxf.Entities.Text(textOut ?? string.Empty, anchor, outHeight)
                        {
                            Rotation = outRotation,
                            Layer = layerRef
                        };
                        if (!string.IsNullOrWhiteSpace(occ.Style) && src.TextStyles.Contains(occ.Style))
                            t.Style = src.TextStyles[occ.Style];
                        if (alignOut.HasValue)
                            t.Alignment = alignOut.Value;
                        if (occ.WidthFactor.HasValue && occ.WidthFactor.Value > 0)
                            t.WidthFactor = occ.WidthFactor.Value;
                        EnsureTextAlignmentPoint(t);
                        src.Entities.Add(t);
                    }
                }

                src.Save(sfd.FileName);
                lblGenerateInfo.Text = $"DXF ({culture}) creato: {sfd.FileName}";
                MessageBox.Show(this, "Generazione COMPLETATA.", "OK",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"Errore generazione DXF:\n{ex.Message}", "ERRORE",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            // [END] OnGenerateSingle
        }
    }
}
