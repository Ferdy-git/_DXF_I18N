using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;

namespace DxfI18N.App
{
    partial class Form1
    {
        private IContainer components = null;

        // === Barra di stato ===
        private StatusStrip statusStrip;
        private ToolStripStatusLabel tsslDbPath;

        // === Tab principali ===
        private TabControl tabControlMain;
        private TabPage tabImport;
        private TabPage tabTranslations;
        private TabPage tabUnified;

        // === IMPORTA DXF (UI) ===
        private Panel panelTopImport;
        private Button btnSelectDb;
        private Button btnPickDxf;
        private Button btnPreview;
        private Button btnApplyFilters;
        private Button btnSelectAllPreview;
        private Button btnDeselectAllPreview;
        private Button btnImport;

        private CheckedListBox clbLayers;
        private CheckBox chkText;
        private CheckBox chkMText;
        private CheckBox chkAttrib;

        private CheckBox chkOnlyLetters;
        private NumericUpDown numMinLen;
        private TextBox txtExcludePatterns;

        private Label lblStatsImport;
        private DataGridView gridOccurrences;
        private DataGridViewCheckBoxColumn colInclude;
        private DataGridViewTextBoxColumn colFile;
        private DataGridViewTextBoxColumn colLayerOriginal;
        private DataGridViewTextBoxColumn colEntityType;
        private DataGridViewTextBoxColumn colOriginalText;
        private DataGridViewTextBoxColumn colPosX;
        private DataGridViewTextBoxColumn colPosY;
        private DataGridViewTextBoxColumn colRotation;
        private DataGridViewTextBoxColumn colHeight;
        private DataGridViewTextBoxColumn colStyle;
        private DataGridViewTextBoxColumn colEntityHandle;
        private DataGridViewTextBoxColumn colKeyId;
        private DataGridViewTextBoxColumn colBlockName;
        private DataGridViewTextBoxColumn colTag;

        // === TRADUZIONI (UI) ===
        private Panel panelTopTranslations;
        private Label lblLang;
        private ComboBox cmbCultureTranslations;
        private Label lblFiltro;
        private TextBox txtFilter;

        private Button btnExportCsv;        // Mancanti
        private Button btnImportCsv;        // Import
        private Button btnExportLanguage;   // Estrai lingua (tutte)

        private Button btnSetApproved;
        private Button btnSetDraft;
        private Button btnSetBlocked;

        private CheckBox chkOverwriteApproved;
        private Label lblStatsTranslations;
        private DataGridView gridTranslations;
        private DataGridViewTextBoxColumn colKeyId2;
        private DataGridViewTextBoxColumn colDefaultText;
        private DataGridViewTextBoxColumn colTranslated;
        private DataGridViewTextBoxColumn colStatus;
        private DataGridViewCheckBoxColumn colOverLengthWarn;
        private DataGridViewTextBoxColumn colNotes;

        // === DXF UNICO (UI) ===
        private Panel panelTopUnified;
        private Label lblLangOut;
        private ComboBox cmbCultureGenerate;
        private CheckBox chkOnlyApprovedGen;
        private Button btnGenerateUnified;
        private Label lblGenerateInfo;
        private DataGridView gridDrawings;
        private DataGridViewTextBoxColumn colDrawingId;
        private DataGridViewTextBoxColumn colPath;
        private DataGridViewTextBoxColumn colImportedAt;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
                components.Dispose();
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            components = new Container();

            // ====== Barra di stato ======
            statusStrip = new StatusStrip();
            tsslDbPath = new ToolStripStatusLabel { Spring = true, Text = "DB: (nessuno)", TextAlign = ContentAlignment.MiddleLeft };
            statusStrip.Items.AddRange(new ToolStripItem[] { tsslDbPath });
            statusStrip.Dock = DockStyle.Bottom;

            // ====== TabControl ======
            tabControlMain = new TabControl { Dock = DockStyle.Fill };
            tabImport = new TabPage { Text = "Importa DXF" };
            tabTranslations = new TabPage { Text = "Traduzioni" };
            tabUnified = new TabPage { Text = "DXF Unico" };
            tabControlMain.TabPages.AddRange(new[] { tabImport, tabTranslations, tabUnified });

            // ====== IMPORTA DXF (pannello alto) ======
            panelTopImport = new Panel { Dock = DockStyle.Top, Height = 160 };

            btnSelectDb = new Button { Name = "btnSelectDb", Text = "Scegli DB…", Location = new Point(10, 10), Size = new Size(100, 30) };
            btnPickDxf = new Button { Name = "btnPickDxf", Text = "Seleziona DXF…", Location = new Point(120, 10), Size = new Size(120, 30) };
            btnPreview = new Button { Name = "btnPreview", Text = "Anteprima", Location = new Point(250, 10), Size = new Size(100, 30) };
            btnApplyFilters = new Button { Name = "btnApplyFilters", Text = "Applica filtri", Location = new Point(360, 10), Size = new Size(110, 30) };
            btnSelectAllPreview = new Button { Name = "btnSelectAllPreview", Text = "Seleziona tutti", Location = new Point(480, 10), Size = new Size(110, 30) };
            btnDeselectAllPreview = new Button { Name = "btnDeselectAllPreview", Text = "Deseleziona tutti", Location = new Point(600, 10), Size = new Size(120, 30) };
            btnImport = new Button { Name = "btnImport", Text = "Importa selezionati", Location = new Point(730, 10), Size = new Size(140, 30) };

            clbLayers = new CheckedListBox { Name = "clbLayers", Location = new Point(10, 50), Size = new Size(260, 95), CheckOnClick = true };

            chkText = new CheckBox { Name = "chkText", Text = "TEXT", Checked = true, AutoSize = true, Location = new Point(280, 55) };
            chkMText = new CheckBox { Name = "chkMText", Text = "MTEXT", Checked = true, AutoSize = true, Location = new Point(280, 80) };
            chkAttrib = new CheckBox { Name = "chkAttrib", Text = "ATTRIB", Checked = false, AutoSize = true, Location = new Point(280, 105) };

            chkOnlyLetters = new CheckBox { Name = "chkOnlyLetters", Text = "Solo testi con lettere", Checked = true, AutoSize = true, Location = new Point(410, 55) };
            numMinLen = new NumericUpDown { Name = "numMinLen", Minimum = 0, Maximum = 100, Value = 2, Size = new Size(60, 23), Location = new Point(410, 80) };
            var lblMin = new Label { Text = "Lungh. minima", AutoSize = true, Location = new Point(480, 82) };

            var lblPat = new Label { Text = "Escludi pattern (regex; separati da ;)", AutoSize = true, Location = new Point(580, 55) };
            txtExcludePatterns = new TextBox { Name = "txtExcludePatterns", Location = new Point(580, 80), Size = new Size(300, 23) };
            txtExcludePatterns.PlaceholderText = @"^\d+$;^\+?\d+V$;^J\d+$;^\d+mm$;^(BK|BN|RD|OG|YE|GN|BU|VT|GY|WH|PK|GNYE)$";

            lblStatsImport = new Label { Name = "lblStatsImport", Text = "—", AutoSize = true, Location = new Point(900, 15), Anchor = AnchorStyles.Top | AnchorStyles.Right };

            panelTopImport.Controls.AddRange(new Control[] {
                btnSelectDb, btnPickDxf, btnPreview, btnApplyFilters, btnSelectAllPreview, btnDeselectAllPreview, btnImport,
                clbLayers, chkText, chkMText, chkAttrib, chkOnlyLetters, numMinLen, lblMin, lblPat, txtExcludePatterns, lblStatsImport
            });

            // ====== IMPORTA DXF (griglia) ======
            gridOccurrences = new DataGridView
            {
                Name = "gridOccurrences",
                Dock = DockStyle.Fill,
                ReadOnly = false,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                SelectionMode = DataGridViewSelectionMode.CellSelect,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
                AllowUserToResizeColumns = true
            };
            gridOccurrences.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;

            colInclude = new DataGridViewCheckBoxColumn { Name = "colInclude", HeaderText = "✓", Width = 40 };
            colFile = new DataGridViewTextBoxColumn { Name = "colFile", HeaderText = "File", Width = 300 };
            colLayerOriginal = new DataGridViewTextBoxColumn { Name = "colLayerOriginal", HeaderText = "Layer" };
            colEntityType = new DataGridViewTextBoxColumn { Name = "colEntityType", HeaderText = "Tipo" };
            colOriginalText = new DataGridViewTextBoxColumn { Name = "colOriginalText", HeaderText = "Testo", Width = 300 };
            colPosX = new DataGridViewTextBoxColumn { Name = "colPosX", HeaderText = "PosX", DefaultCellStyle = new DataGridViewCellStyle { Format = "N2" } };
            colPosY = new DataGridViewTextBoxColumn { Name = "colPosY", HeaderText = "PosY", DefaultCellStyle = new DataGridViewCellStyle { Format = "N2" } };
            colRotation = new DataGridViewTextBoxColumn { Name = "colRotation", HeaderText = "Rotazione", DefaultCellStyle = new DataGridViewCellStyle { Format = "N2" } };
            colHeight = new DataGridViewTextBoxColumn { Name = "colHeight", HeaderText = "Altezza", DefaultCellStyle = new DataGridViewCellStyle { Format = "N2" } };
            colStyle = new DataGridViewTextBoxColumn { Name = "colStyle", HeaderText = "Stile" };
            colEntityHandle = new DataGridViewTextBoxColumn { Name = "colEntityHandle", HeaderText = "Handle", Width = 120 };
            colKeyId = new DataGridViewTextBoxColumn { Name = "colKeyId", HeaderText = "KeyId", Width = 70 };
            colBlockName = new DataGridViewTextBoxColumn { Name = "colBlockName", HeaderText = "Blocco" };
            colTag = new DataGridViewTextBoxColumn { Name = "colTag", HeaderText = "Tag" };

            gridOccurrences.Columns.AddRange(new DataGridViewColumn[] {
                colInclude, colFile, colLayerOriginal, colEntityType, colOriginalText,
                colPosX, colPosY, colRotation, colHeight, colStyle, colEntityHandle, colKeyId, colBlockName, colTag
            });

            tabImport.Controls.Add(gridOccurrences);
            tabImport.Controls.Add(panelTopImport);

            // ====== TRADUZIONI (pannello alto) ======
            panelTopTranslations = new Panel { Dock = DockStyle.Top, Height = 86 };

            lblLang = new Label { Text = "Lingua:", AutoSize = true, Location = new Point(10, 12) };
            cmbCultureTranslations = new ComboBox { Name = "cmbCultureTranslations", DropDownStyle = ComboBoxStyle.DropDownList, Location = new Point(65, 8), Width = 140 };

            lblFiltro = new Label { Text = "Filtro:", AutoSize = true, Location = new Point(220, 12) };
            txtFilter = new TextBox { Name = "txtFilter", Location = new Point(260, 8), Width = 220 };

            btnExportCsv = new Button { Name = "btnExportCsv", Text = "Esporta mancanti", Size = new Size(130, 28), Location = new Point(10, 46) };
            btnImportCsv = new Button { Name = "btnImportCsv", Text = "Importa CSV", Size = new Size(110, 28), Location = new Point(150, 46) };
            btnExportLanguage = new Button { Name = "btnExportLanguage", Text = "Estrai lingua", Size = new Size(110, 28), Location = new Point(270, 46) };

            btnSetApproved = new Button { Name = "btnSetApproved", Text = "Approved", Size = new Size(90, 28), Location = new Point(390, 46) };
            btnSetDraft = new Button { Name = "btnSetDraft", Text = "Draft", Size = new Size(80, 28), Location = new Point(485, 46) };
            btnSetBlocked = new Button { Name = "btnSetBlocked", Text = "Blocked", Size = new Size(90, 28), Location = new Point(570, 46) };

            chkOverwriteApproved = new CheckBox { Name = "chkOverwriteApproved", Text = "Consenti sovrascrittura Approved (import)", AutoSize = true, Location = new Point(670, 50) };

            lblStatsTranslations = new Label { Name = "lblStatsTranslations", Text = "—", AutoSize = true, Location = new Point(900, 12), Anchor = AnchorStyles.Top | AnchorStyles.Right };

            panelTopTranslations.Controls.AddRange(new Control[] {
                lblLang, cmbCultureTranslations, lblFiltro, txtFilter,
                btnExportCsv, btnImportCsv, btnExportLanguage,
                btnSetApproved, btnSetDraft, btnSetBlocked,
                chkOverwriteApproved, lblStatsTranslations
            });

            // ====== TRADUZIONI (griglia) ======
            gridTranslations = new DataGridView
            {
                Name = "gridTranslations",
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
                AllowUserToResizeColumns = true
            };
            gridTranslations.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;

            colKeyId2 = new DataGridViewTextBoxColumn { Name = "colKeyId2", HeaderText = "KeyId", Width = 70 };
            colDefaultText = new DataGridViewTextBoxColumn { Name = "colDefaultText", HeaderText = "DefaultText", Width = 320 };
            colTranslated = new DataGridViewTextBoxColumn { Name = "colTranslated", HeaderText = "Translated", Width = 320 };
            colStatus = new DataGridViewTextBoxColumn { Name = "colStatus", HeaderText = "Status", Width = 90 };
            colOverLengthWarn = new DataGridViewCheckBoxColumn { Name = "colOverLengthWarn", HeaderText = "Warn", Width = 60 };
            colNotes = new DataGridViewTextBoxColumn { Name = "colNotes", HeaderText = "Note", Width = 200 };

            gridTranslations.Columns.AddRange(new DataGridViewColumn[] {
                colKeyId2, colDefaultText, colTranslated, colStatus, colOverLengthWarn, colNotes
            });

            tabTranslations.Controls.Add(gridTranslations);
            tabTranslations.Controls.Add(panelTopTranslations);

            // ====== DXF UNICO (pannello alto) ======
            panelTopUnified = new Panel { Dock = DockStyle.Top, Height = 86 };

            lblLangOut = new Label { Text = "Lingua attiva:", AutoSize = true, Location = new Point(10, 12) };
            cmbCultureGenerate = new ComboBox { Name = "cmbCultureGenerate", DropDownStyle = ComboBoxStyle.DropDownList, Location = new Point(100, 8), Width = 140 };

            chkOnlyApprovedGen = new CheckBox { Name = "chkOnlyApprovedGen", Text = "Usa solo traduzioni Approved", AutoSize = true, Location = new Point(260, 10) };
            btnGenerateUnified = new Button { Name = "btnGenerateUnified", Text = "Genera DXF unico…", Size = new Size(160, 28), Location = new Point(10, 46) };

            lblGenerateInfo = new Label { Name = "lblGenerateInfo", Text = "—", AutoSize = true, Location = new Point(180, 50) };

            panelTopUnified.Controls.AddRange(new Control[] {
                lblLangOut, cmbCultureGenerate, chkOnlyApprovedGen, btnGenerateUnified, lblGenerateInfo
            });

            // ====== DXF UNICO (griglia) ======
            gridDrawings = new DataGridView
            {
                Name = "gridDrawings",
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
                AllowUserToResizeColumns = true
            };
            gridDrawings.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;

            colDrawingId = new DataGridViewTextBoxColumn { Name = "colDrawingId", HeaderText = "Id", Width = 60 };
            colPath = new DataGridViewTextBoxColumn { Name = "colPath", HeaderText = "Path", Width = 650 };
            colImportedAt = new DataGridViewTextBoxColumn { Name = "colImportedAt", HeaderText = "ImportedAt", Width = 140 };

            gridDrawings.Columns.AddRange(new DataGridViewColumn[] {
                colDrawingId, colPath, colImportedAt
            });

            tabUnified.Controls.Add(gridDrawings);
            tabUnified.Controls.Add(panelTopUnified);

            // ====== Form ======
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1200, 780);
            Font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point);
            MinimumSize = new Size(1100, 700);
            StartPosition = FormStartPosition.CenterScreen;
            Text = "DXF Multilingua";

            Controls.Add(tabControlMain);
            Controls.Add(statusStrip);

            ResumeLayout(false);
            PerformLayout();
        }
    }
}
