using System.Data;
using System.Data.Common;
using System.Text;
using Advantage.Data.Provider;

namespace PdcGui;

public class MainForm : Form
{
    private MenuStrip menuStrip = null!;
    private StatusStrip statusStrip = null!;
    private ToolStripStatusLabel statusLabel = null!;
    private ToolStripStatusLabel userLabel = null!;
    private ToolStripStatusLabel clockLabel = null!;
    private ToolStripStatusLabel adsLabel = null!;
    private DataGridView mainGrid = null!;
    private System.Windows.Forms.Timer clockTimer = null!;
    private string currentUser = "UNKNOWN";
    private bool adsAvailable = false;

    public MainForm(string userName)
    {
        currentUser = userName.ToUpper();
        InitializeUI();
    }

    private void InitializeUI()
    {
        // --- Window ---
        Text = "AVX PDC — Production Data Collection System";
        Size = new Size(1024, 700);
        StartPosition = FormStartPosition.CenterScreen;
        Icon = SystemIcons.Application;
        Font = new Font("Segoe UI", 10f);

        // --- Menu ---
        menuStrip = new MenuStrip();

        // Batch Operations
        var mnuBatch = new ToolStripMenuItem("&Batch Operations");
        mnuBatch.DropDownItems.AddRange(new ToolStripItem[]
        {
            MakeMenuItem("Batch &Arrival", Keys.F2, OnBatchArrive),
            MakeMenuItem("&Start Stage", Keys.F3, OnBatchEnter),
            MakeMenuItem("Stage &Completion", Keys.F4, OnBatchLeave),
            new ToolStripSeparator(),
            MakeMenuItem("Ca&ncel Completion", Keys.F5, OnCancelFinish),
            MakeMenuItem("&Reprint Completion", Keys.F6, OnReprintBN),
            new ToolStripSeparator(),
            MakeMenuItem("Batch &Location", Keys.None, OnBatchLocation),
        });

        // Queries
        var mnuQueries = new ToolStripMenuItem("&Queries");
        mnuQueries.DropDownItems.AddRange(new ToolStripItem[]
        {
            MakeMenuItem("WIP by &Process", Keys.None, OnStockInProc),
            MakeMenuItem("WIP by &Workstation", Keys.None, OnStockInWork),
            MakeMenuItem("Batch &Path Query", Keys.None, OnBatchPath),
            new ToolStripSeparator(),
            MakeMenuItem("&Shipments Query", Keys.None, OnShipQuery),
        });

        // Shipments
        var mnuShipments = new ToolStripMenuItem("&Shipments");
        mnuShipments.DropDownItems.AddRange(new ToolStripItem[]
        {
            MakeMenuItem("Batch &Arrival for Outside Prod.", Keys.None, OnCzhehSends),
            MakeMenuItem("Proforma &Invoice Printing", Keys.None, OnCzhehDeliv),
        });

        // Labels
        var mnuLabels = new ToolStripMenuItem("&Labels");
        mnuLabels.DropDownItems.AddRange(new ToolStripItem[]
        {
            MakeMenuItem("&Packing Labels", Keys.None, OnPackingLabels),
            MakeMenuItem("P&roduction Labels", Keys.None, OnProductionLabels),
        });

        // Reports
        var mnuReports = new ToolStripMenuItem("&Reports");
        mnuReports.DropDownItems.AddRange(new ToolStripItem[]
        {
            MakeMenuItem("&Yield Analysis", Keys.None, OnYieldReport),
            MakeMenuItem("Production &Area Stages", Keys.None, OnProdArea),
        });

        // Tools
        var mnuTools = new ToolStripMenuItem("&Tools");
        mnuTools.DropDownItems.AddRange(new ToolStripItem[]
        {
            MakeMenuItem("&Open DBF Table...", Keys.Control | Keys.O, OnOpenDbf),
            new ToolStripSeparator(),
            MakeMenuItem("&Settings...", Keys.None, OnSettings),
        });

        // Help
        var mnuHelp = new ToolStripMenuItem("&Help");
        mnuHelp.DropDownItems.AddRange(new ToolStripItem[]
        {
            MakeMenuItem("&Keyboard Shortcuts", Keys.F1, OnHelpKeys),
            new ToolStripSeparator(),
            MakeMenuItem("&About PDC", Keys.None, OnAbout),
        });

        menuStrip.Items.AddRange(new ToolStripItem[]
        {
            mnuBatch, mnuQueries, mnuShipments, mnuLabels, mnuReports, mnuTools, mnuHelp
        });

        // --- Status Bar ---
        statusStrip = new StatusStrip();
        statusLabel = new ToolStripStatusLabel("Ready")
        {
            Spring = true,
            TextAlign = ContentAlignment.MiddleLeft
        };
        userLabel = new ToolStripStatusLabel($"User: {currentUser}")
        {
            BorderSides = ToolStripStatusLabelBorderSides.Left,
        };
        clockLabel = new ToolStripStatusLabel(DateTime.Now.ToString("HH:mm:ss"))
        {
            BorderSides = ToolStripStatusLabelBorderSides.Left,
        };
        adsLabel = new ToolStripStatusLabel("ADS: checking...")
        {
            BorderSides = ToolStripStatusLabelBorderSides.Left,
        };
        statusStrip.Items.AddRange(new ToolStripItem[] { statusLabel, adsLabel, userLabel, clockLabel });

        // --- Clock Timer ---
        clockTimer = new System.Windows.Forms.Timer { Interval = 1000 };
        clockTimer.Tick += (s, e) => clockLabel.Text = DateTime.Now.ToString("HH:mm:ss");
        clockTimer.Start();

        // --- Main Grid ---
        mainGrid = new DataGridView
        {
            Dock = DockStyle.Fill,
            ReadOnly = true,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            BackgroundColor = SystemColors.Window,
            BorderStyle = BorderStyle.None,
        };

        // --- Layout ---
        Controls.Add(mainGrid);
        Controls.Add(statusStrip);
        Controls.Add(menuStrip);
        MainMenuStrip = menuStrip;

        // --- Welcome ---
        SetStatus($"Welcome, {currentUser}. PDC GUI ready.");

        // --- Check ADS availability ---
        CheckAdsAvailability();
    }

    private void CheckAdsAvailability()
    {
        try
        {
            // Try to load ADS — just instantiate, no connection needed
            var testConn = new AdsConnection();
            testConn.Dispose();
            adsAvailable = true;
            adsLabel.Text = "ADS: Local";
            adsLabel.ForeColor = Color.DarkGreen;
        }
        catch
        {
            adsAvailable = false;
            adsLabel.Text = "ADS: N/A (using DbfReader)";
            adsLabel.ForeColor = Color.DarkRed;
        }
    }

    // ===== Helper =====
    private ToolStripMenuItem MakeMenuItem(string text, Keys shortcut, EventHandler handler)
    {
        var item = new ToolStripMenuItem(text);
        if (shortcut != Keys.None) item.ShortcutKeys = shortcut;
        item.Click += handler;
        return item;
    }

    private void SetStatus(string msg)
    {
        statusLabel.Text = msg;
    }

    private void ShowNotImplemented(string feature)
    {
        MessageBox.Show(
            $"'{feature}' will be implemented in the next phase.\n\nThis is the GUI skeleton — business logic from Harbour PDC will be connected here.",
            "Not Yet Implemented",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information);
    }

    // ===== Open DBF — ADS primary, DbfDataReader fallback =====
    private void OnOpenDbf(object? sender, EventArgs e)
    {
        using var dlg = new OpenFileDialog
        {
            Title = "Open DBF Table",
            Filter = "dBASE/DAT files (*.dbf;*.dat)|*.dbf;*.dat|All files (*.*)|*.*",
            InitialDirectory = @"C:\Users\AVXUser\PDC-clean\BMS"
        };

        if (dlg.ShowDialog() != DialogResult.OK) return;

        try
        {
            SetStatus($"Loading {Path.GetFileName(dlg.FileName)}...");
            Cursor = Cursors.WaitCursor;

            DataTable dt;
            int count;
            string engine;

            if (adsAvailable)
            {
                (dt, count) = LoadViaAds(dlg.FileName);
                engine = "ADS";
            }
            else
            {
                (dt, count) = LoadViaDbfReader(dlg.FileName);
                engine = "DbfReader";
            }

            mainGrid.DataSource = dt;
            Text = $"AVX PDC — {Path.GetFileName(dlg.FileName)} ({count} records)";
            SetStatus($"[{engine}] Loaded {Path.GetFileName(dlg.FileName)}: {count} records, {dt.Columns.Count} columns");
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error reading DBF:\n{ex.Message}\n\n{ex.GetType().Name}", "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            SetStatus("Error loading file");
        }
        finally
        {
            Cursor = Cursors.Default;
        }
    }

    private (DataTable dt, int count) LoadViaAds(string filePath)
    {
        var dt = new DataTable();
        string dir = Path.GetDirectoryName(filePath) ?? ".";
        string file = Path.GetFileName(filePath);

        string connStr = $"Data Source={dir};ServerType=LOCAL;TableType=CDX;LockMode=COMPATIBLE;CharType=OEM;TrimTrailingSpaces=TRUE;";
        using var conn = new AdsConnection(connStr);
        conn.Open();

        using var cmd = new AdsCommand($"SELECT * FROM \"{file}\"", conn);
        using var adapter = new AdsDataAdapter(cmd);

        int count = adapter.Fill(dt);
        return (dt, count);
    }

    private (DataTable dt, int count) LoadViaDbfReader(string filePath)
    {
        var dt = new DataTable();
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        var options = new DbfDataReader.DbfDataReaderOptions
        {
            SkipDeletedRecords = true,
            Encoding = Encoding.GetEncoding(862)
        };
        using var dbfReader = new DbfDataReader.DbfDataReader(filePath, options);

        for (int i = 0; i < dbfReader.FieldCount; i++)
        {
            dt.Columns.Add(new DataColumn(dbfReader.GetName(i), typeof(string)));
        }

        int count = 0;
        while (dbfReader.Read() && count < 10000)
        {
            var row = dt.NewRow();
            for (int i = 0; i < dbfReader.FieldCount; i++)
            {
                try
                {
                    row[i] = dbfReader.IsDBNull(i) ? DBNull.Value : dbfReader.GetValue(i)?.ToString() ?? "";
                }
                catch
                {
                    row[i] = DBNull.Value;
                }
            }
            dt.Rows.Add(row);
            count++;
        }

        return (dt, count);
    }

    // ===== Batch Operations (stubs) =====
    private void OnBatchArrive(object? sender, EventArgs e) => ShowNotImplemented("Batch Arrival (BNarrive)");
    private void OnBatchEnter(object? sender, EventArgs e) => ShowNotImplemented("Start Stage (BNenter)");
    private void OnBatchLeave(object? sender, EventArgs e) => ShowNotImplemented("Stage Completion (BNleave)");
    private void OnCancelFinish(object? sender, EventArgs e) => ShowNotImplemented("Cancel Completion");
    private void OnReprintBN(object? sender, EventArgs e) => ShowNotImplemented("Reprint Completion");
    private void OnBatchLocation(object? sender, EventArgs e) => ShowNotImplemented("Batch Location");

    // ===== Queries (stubs) =====
    private void OnStockInProc(object? sender, EventArgs e) => ShowNotImplemented("WIP by Process (StockInProc)");
    private void OnStockInWork(object? sender, EventArgs e) => ShowNotImplemented("WIP by Workstation (StockInWork)");
    private void OnBatchPath(object? sender, EventArgs e) => ShowNotImplemented("Batch Path Query (BNpath)");
    private void OnShipQuery(object? sender, EventArgs e) => ShowNotImplemented("Shipments Query (ShpQuery)");

    // ===== Shipments (stubs) =====
    private void OnCzhehSends(object? sender, EventArgs e) => ShowNotImplemented("Batch Arrival for Outside Prod.");
    private void OnCzhehDeliv(object? sender, EventArgs e) => ShowNotImplemented("Proforma Invoice Printing");

    // ===== Labels (stubs) =====
    private void OnPackingLabels(object? sender, EventArgs e) => ShowNotImplemented("Packing Labels");
    private void OnProductionLabels(object? sender, EventArgs e) => ShowNotImplemented("Production Labels");

    // ===== Reports (stubs) =====
    private void OnYieldReport(object? sender, EventArgs e) => ShowNotImplemented("Yield Analysis (RPQC36V1)");
    private void OnProdArea(object? sender, EventArgs e) => ShowNotImplemented("Production Area Stages");

    // ===== Tools =====
    private void OnSettings(object? sender, EventArgs e) => ShowNotImplemented("Settings");

    // ===== Help =====
    private void OnHelpKeys(object? sender, EventArgs e)
    {
        string help = """
            PDC GUI Keyboard Shortcuts
            ═════════════════════════════
            F1          — This help
            F2          — Batch Arrival
            F3          — Start Stage
            F4          — Stage Completion
            F5          — Cancel Completion
            F6          — Reprint Completion
            Ctrl+O      — Open DBF Table

            In DataGrid:
            ↑↓          — Navigate rows
            Ctrl+Home   — First record
            Ctrl+End    — Last record
            Ctrl+F      — Find (future)
            """;

        MessageBox.Show(help, "Keyboard Shortcuts", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private void OnAbout(object? sender, EventArgs e)
    {
        MessageBox.Show(
            "AVX PDC — Production Data Collection System\n\n" +
            "C# WinForms GUI Edition\n" +
            "Based on Clipper 5.3 → Harbour 3.2 codebase\n\n" +
            "Kyocera AVX Components\n" +
            $".NET {Environment.Version}",
            "About PDC",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information);
    }
}
