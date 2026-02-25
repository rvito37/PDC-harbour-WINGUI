using System.Data;
using Advantage.Data.Provider;

namespace PdcGui;

/// <summary>
/// WIP by Process query screen — migrated from Harbour STOKPROC.PRG
/// Shows all batches currently at a selected process, with totals.
/// Read-only view. Data source: D_LINE.DBF + C_PROC.DBF via ADS.
/// </summary>
public class StockInProcForm : Form
{
    private readonly string dataDir;

    // Controls
    private ComboBox cmbProcess = null!;
    private Label lblProcessName = null!;
    private Label lblIndexInfo = null!;
    private DataGridView grid = null!;
    private StatusStrip statusStrip = null!;
    private ToolStripStatusLabel statusLabel = null!;
    private ToolStripStatusLabel countLabel = null!;
    private Panel topPanel = null!;
    private Button btnQuery = null!;
    private Button btnTotals = null!;
    private Button btnSearch = null!;
    // Data
    private DataTable? dtProcesses;
    private DataTable? dtResult;
    private string currentProcessId = "";
    private int totalRecords = 0;

    // Sort indices matching Harbour CDX tags
    private readonly string[][] sortModes = new[]
    {
        new[] { "ISLACK",   "Prior+Slack+Purp+Stat+B/N",           "B_prior, B_dprom, B_purp, B_stat, B_id" },
        new[] { "NWISLACK", "Type+Line+Size+Prior+Slack+Purp",     "B_type, pLine_id, Size_id, B_prior, B_dprom, B_purp" },
    };
    private readonly string[][] sortModes190 = new[]
    {
        new[] { "QCISLACK", "Stat+Prior+Slack+Purp+B/N",           "B_stat, B_prior, B_dprom, B_purp, B_id" },
        new[] { "ISLACK",   "Prior+Slack+Purp+Stat+B/N",           "B_prior, B_dprom, B_stat, B_purp, B_id" },
        new[] { "NWISLACK", "Type+Line+Size+Prior+Slack+Purp",     "B_type, pLine_id, Size_id, B_prior, B_dprom, B_purp" },
    };
    private int currentSortIndex = 0;

    public StockInProcForm(string dataDirectory)
    {
        dataDir = dataDirectory;
        InitializeUI();
        LoadProcessList();
    }

    private void InitializeUI()
    {
        Text = "WIP by Process — Process WIP Query";
        Size = new Size(1280, 750);
        StartPosition = FormStartPosition.CenterParent;
        Font = new Font("Segoe UI", 9.5f);
        KeyPreview = true;
        KeyDown += OnFormKeyDown;

        // === Top Panel ===
        topPanel = new Panel
        {
            Dock = DockStyle.Top,
            Height = 50,
            Padding = new Padding(8, 8, 8, 4),
            BackColor = Color.FromArgb(240, 240, 250),
        };

        var lblProc = new Label
        {
            Text = "Process:",
            AutoSize = true,
            Location = new Point(10, 15),
            Font = new Font("Segoe UI", 10f, FontStyle.Bold),
        };

        cmbProcess = new ComboBox
        {
            Location = new Point(85, 12),
            Width = 90,
            DropDownStyle = ComboBoxStyle.DropDown,
            Font = new Font("Consolas", 11f),
        };
        cmbProcess.KeyDown += (s, e) =>
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                ExecuteQuery();
            }
        };

        lblProcessName = new Label
        {
            Text = "",
            AutoSize = true,
            Location = new Point(185, 15),
            Font = new Font("Segoe UI", 10f),
            ForeColor = Color.DarkBlue,
        };

        btnQuery = new Button
        {
            Text = "Show WIP",
            Location = new Point(420, 10),
            Size = new Size(90, 30),
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(0, 120, 215),
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9f, FontStyle.Bold),
        };
        btnQuery.Click += (s, e) => ExecuteQuery();

        btnSearch = new Button
        {
            Text = "Find Batch (F4)",
            Location = new Point(520, 10),
            Size = new Size(110, 30),
            FlatStyle = FlatStyle.Flat,
            Enabled = false,
        };
        btnSearch.Click += (s, e) => SearchBatch();

        btnTotals = new Button
        {
            Text = "Totals (Ctrl+T)",
            Location = new Point(640, 10),
            Size = new Size(110, 30),
            FlatStyle = FlatStyle.Flat,
            Enabled = false,
        };
        btnTotals.Click += (s, e) => ShowTotals();

        lblIndexInfo = new Label
        {
            Text = "",
            AutoSize = true,
            Location = new Point(760, 15),
            Font = new Font("Consolas", 9f),
            ForeColor = Color.DarkGreen,
        };

        topPanel.Controls.AddRange(new Control[]
        {
            lblProc, cmbProcess, lblProcessName, btnQuery,
            btnSearch, btnTotals, lblIndexInfo
        });

        // === Grid ===
        grid = new DataGridView
        {
            Dock = DockStyle.Fill,
            ReadOnly = true,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            BackgroundColor = SystemColors.Window,
            BorderStyle = BorderStyle.None,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
            ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
            {
                Font = new Font("Segoe UI", 9f, FontStyle.Bold),
                BackColor = Color.FromArgb(0, 51, 102),
                ForeColor = Color.White,
                Alignment = DataGridViewContentAlignment.MiddleCenter,
            },
            EnableHeadersVisualStyles = false,
        };
        grid.DoubleBuffered_();
        grid.CellFormatting += OnCellFormatting;
        grid.ColumnHeaderMouseClick += OnColumnHeaderClick;

        // === Status Bar ===
        statusStrip = new StatusStrip();
        statusLabel = new ToolStripStatusLabel("Select a process and click 'Show WIP'")
        {
            Spring = true,
            TextAlign = ContentAlignment.MiddleLeft,
        };
        countLabel = new ToolStripStatusLabel("Records: 0")
        {
            BorderSides = ToolStripStatusLabelBorderSides.Left,
        };
        statusStrip.Items.AddRange(new ToolStripItem[] { statusLabel, countLabel });

        // === Layout ===
        Controls.Add(grid);
        Controls.Add(topPanel);
        Controls.Add(statusStrip);
    }

    // === Load process list from C_PROC ===
    private void LoadProcessList()
    {
        try
        {
            string connStr = $"Data Source={dataDir};ServerType=LOCAL;TableType=CDX;LockMode=COMPATIBLE;CharType=OEM;TrimTrailingSpaces=TRUE;";
            using var conn = new AdsConnection(connStr);
            conn.Open();

            string sql = "SELECT proc_id, proc_nme FROM \"C_PROC.DBF\" ORDER BY proc_id";
            using var cmd = new AdsCommand(sql, conn);
            using var adapter = new AdsDataAdapter(cmd);
            dtProcesses = new DataTable();
            adapter.Fill(dtProcesses);

            cmbProcess.Items.Clear();
            foreach (DataRow row in dtProcesses.Rows)
            {
                string procId = row["proc_id"]?.ToString()?.Trim() ?? "";
                if (!string.IsNullOrWhiteSpace(procId))
                    cmbProcess.Items.Add(procId);
            }

            statusLabel.Text = $"Loaded {cmbProcess.Items.Count} processes. Enter process ID and press Enter or click 'Show WIP'.";
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error loading process list:\n{ex.Message}", "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            statusLabel.Text = "Error loading processes";
        }
    }

    // === Main query: load WIP for selected process ===
    private void ExecuteQuery()
    {
        string procId = cmbProcess.Text.Trim();
        if (string.IsNullOrWhiteSpace(procId))
        {
            MessageBox.Show("Please enter a Process ID.", "Input Required",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            cmbProcess.Focus();
            return;
        }

        // Pad to 5 chars like Harbour PICTURE "999.9"
        procId = procId.PadRight(5);

        try
        {
            Cursor = Cursors.WaitCursor;
            statusLabel.Text = $"Querying WIP for process {procId.Trim()}...";
            Application.DoEvents();

            // Look up process name
            string procName = LookupProcessName(procId);
            if (string.IsNullOrEmpty(procName))
            {
                MessageBox.Show($"Process '{procId.Trim()}' not found in C_PROC.", "Not Found",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                Cursor = Cursors.Default;
                return;
            }
            lblProcessName.Text = procName;
            currentProcessId = procId;
            currentSortIndex = 0;

            // Query D_LINE
            LoadWipData(procId);

            // Enable buttons
            btnSearch.Enabled = true;
            btnTotals.Enabled = true;

            // Set index label
            UpdateIndexLabel();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error querying WIP:\n{ex.Message}\n\n{ex.StackTrace}", "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            statusLabel.Text = "Error";
        }
        finally
        {
            Cursor = Cursors.Default;
        }
    }

    private string LookupProcessName(string procId)
    {
        if (dtProcesses == null) return "";
        foreach (DataRow row in dtProcesses.Rows)
        {
            string id = row["proc_id"]?.ToString()?.PadRight(5) ?? "";
            if (id == procId)
                return row["proc_nme"]?.ToString()?.Trim() ?? "";
        }
        return "";
    }

    private void LoadWipData(string procId)
    {
        string connStr = $"Data Source={dataDir};ServerType=LOCAL;TableType=CDX;LockMode=COMPATIBLE;CharType=OEM;TrimTrailingSpaces=TRUE;";
        using var conn = new AdsConnection(connStr);
        conn.Open();

        // Count total records for this process
        using var countCmd = new AdsCommand(
            $"SELECT COUNT(*) FROM \"D_LINE.DBF\" WHERE CpProc_id = '{procId}'", conn);
        totalRecords = (int)countCmd.ExecuteScalar();

        if (totalRecords == 0)
        {
            MessageBox.Show("No records found for this process!", "Empty",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            grid.DataSource = null;
            countLabel.Text = "Records: 0";
            statusLabel.Text = "No records found";
            return;
        }

        // Harbour original: browse only shows active batches (not C=Completed, D=Deleted)
        string sql = $@"SELECT
            B_id       AS [Batch],
            B_purp     AS [Pu],
            B_stat     AS [St],
            pLine_id   AS [P_Line],
            B_type     AS [Bt],
            Size_id    AS [Size],
            Value_id   AS [Val],
            Cp_bQtyW   AS [Wafers],
            Cp_bQtyS   AS [Strips],
            Cp_bQtyP   AS [Pcs],
            Cp_dArr    AS [Arrive_Date],
            B_prior    AS [Prior],
            B_wkprom   AS [WkProm],
            B_dprom    AS [Promise_Date],
            Pp_widFin  AS [Delivered_By],
            Cp_widArr  AS [Received_By],
            B_remark   AS [Comments],
            Cp_dSta    AS [Started],
            CpProc_id  AS [Proc_ID]
            FROM ""D_LINE.DBF""
            WHERE CpProc_id = '{procId}'
            AND UPPER(B_stat) NOT IN ('C','D')";

        // Apply sort order
        var sortModeArray = (procId.Trim() == "190.0") ? sortModes190 : sortModes;
        if (currentSortIndex < sortModeArray.Length)
            sql += $" ORDER BY {sortModeArray[currentSortIndex][2]}";

        using var cmd = new AdsCommand(sql, conn);
        using var adapter = new AdsDataAdapter(cmd);
        dtResult = new DataTable();
        adapter.Fill(dtResult);

        // Add computed columns
        AddComputedColumns();

        // Bind to grid
        grid.DataSource = null;
        grid.DataSource = dtResult;

        // Configure columns visibility and order
        ConfigureGridColumns();

        countLabel.Text = $"Records: {dtResult.Rows.Count} of {totalRecords} total";
        statusLabel.Text = $"WIP for process {procId.Trim()} — {lblProcessName.Text} — {dtResult.Rows.Count} active batches";
    }

    private void AddComputedColumns()
    {
        if (dtResult == null) return;

        // Row number
        var colRow = new DataColumn("#", typeof(int));
        dtResult.Columns.Add(colRow);

        // Days in Process = Today - Cp_dArr
        var colDays = new DataColumn("Days_In_Proc", typeof(int));
        dtResult.Columns.Add(colDays);

        // Slack = B_dprom - Today (simplified; Harbour version subtracts lead time too)
        var colSlack = new DataColumn("Slack", typeof(int));
        dtResult.Columns.Add(colSlack);

        // Priority display = B_prior + B_wkprom
        var colPriority = new DataColumn("Priority", typeof(string));
        dtResult.Columns.Add(colPriority);

        int rowNum = 0;
        foreach (DataRow row in dtResult.Rows)
        {
            rowNum++;
            row["#"] = rowNum;

            // Days in process
            if (row["Arrive_Date"] != DBNull.Value && row["Arrive_Date"] is DateTime arrDate)
            {
                row["Days_In_Proc"] = (DateTime.Today - arrDate).Days;
            }

            // Slack
            if (row["Promise_Date"] != DBNull.Value && row["Promise_Date"] is DateTime promDate)
            {
                row["Slack"] = (promDate - DateTime.Today).Days;
            }

            // Priority
            string prior = row["Prior"]?.ToString()?.Trim() ?? "";
            string wkprom = row["WkProm"]?.ToString()?.Trim() ?? "";
            row["Priority"] = prior + wkprom;
        }
    }

    private void ConfigureGridColumns()
    {
        if (grid.Columns.Count == 0) return;

        // Define display order matching Harbour English layout
        var displayOrder = new[]
        {
            ("#",             40),
            ("Batch",         75),
            ("Pu",            30),
            ("St",            30),
            ("P_Line",        55),
            ("Bt",            30),
            ("Size",          50),
            ("Val",           40),
            ("Wafers",        60),
            ("Strips",        65),
            ("Pcs",           85),
            ("Days_In_Proc",  55),
            ("Priority",      60),
            ("Slack",         50),
            ("Arrive_Date",   90),
            ("Delivered_By",  85),
            ("Received_By",   85),
            ("Comments",     180),
        };

        // Hide helper columns
        string[] hidden = { "Prior", "WkProm", "Promise_Date", "Started", "Proc_ID" };
        foreach (var col in hidden)
        {
            if (grid.Columns.Contains(col))
                grid.Columns[col]!.Visible = false;
        }

        int order = 0;
        foreach (var (name, width) in displayOrder)
        {
            if (grid.Columns.Contains(name))
            {
                grid.Columns[name]!.DisplayIndex = order;
                grid.Columns[name]!.Width = width;
                order++;

                // Right-align numeric columns
                if (name is "Wafers" or "Strips" or "Pcs" or "Days_In_Proc" or "Slack" or "#")
                {
                    grid.Columns[name]!.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                }
            }
        }
    }

    // === Cell formatting: grey out started batches (Cp_dSta not empty) ===
    private void OnCellFormatting(object? sender, DataGridViewCellFormattingEventArgs e)
    {
        if (dtResult == null || e.RowIndex < 0 || e.RowIndex >= dtResult.Rows.Count) return;

        var row = dtResult.Rows[e.RowIndex];
        var started = row["Started"];

        // Harbour: colorBlock checks EMPTY(D_line->Cp_dSta) — started batches get dimmed
        if (started != DBNull.Value && started is DateTime dt && dt != DateTime.MinValue)
        {
            e.CellStyle!.ForeColor = Color.Gray;
            e.CellStyle!.BackColor = Color.FromArgb(245, 245, 245);
        }

        // Color negative slack red
        if (grid.Columns[e.ColumnIndex].Name == "Slack" && e.Value is int slack && slack < 0)
        {
            e.CellStyle!.ForeColor = Color.Red;
            e.CellStyle!.Font = new Font(grid.Font, FontStyle.Bold);
        }
    }

    // === Ctrl+S style sort toggle via column header click ===
    private void OnColumnHeaderClick(object? sender, DataGridViewCellMouseEventArgs e)
    {
        // Any header click cycles to next sort mode (like Alt+S in Harbour)
        CycleSortOrder();
    }

    private void CycleSortOrder()
    {
        if (string.IsNullOrEmpty(currentProcessId) || dtResult == null) return;

        var sortModeArray = (currentProcessId.Trim() == "190.0") ? sortModes190 : sortModes;
        currentSortIndex = (currentSortIndex + 1) % sortModeArray.Length;

        // Reload with new sort
        try
        {
            Cursor = Cursors.WaitCursor;
            LoadWipData(currentProcessId);
            UpdateIndexLabel();
        }
        finally
        {
            Cursor = Cursors.Default;
        }
    }

    private void UpdateIndexLabel()
    {
        var sortModeArray = (currentProcessId.Trim() == "190.0") ? sortModes190 : sortModes;
        if (currentSortIndex < sortModeArray.Length)
            lblIndexInfo.Text = $"Index: {sortModeArray[currentSortIndex][1]}";
    }

    // === Search batch by ID (F4) ===
    private void SearchBatch()
    {
        if (dtResult == null || dtResult.Rows.Count == 0) return;

        string input = "";
        using var dlg = new Form
        {
            Text = "Find Batch",
            Size = new Size(320, 150),
            StartPosition = FormStartPosition.CenterParent,
            FormBorderStyle = FormBorderStyle.FixedDialog,
            MaximizeBox = false,
            MinimizeBox = false,
        };

        var lbl = new Label { Text = "Enter Batch Number (6 digits):", Location = new Point(15, 15), AutoSize = true };
        var txt = new TextBox { Location = new Point(15, 45), Width = 120, MaxLength = 6, Font = new Font("Consolas", 12f) };
        var btnOk = new Button { Text = "Find", Location = new Point(150, 43), Size = new Size(70, 30), DialogResult = DialogResult.OK };
        var btnCancel = new Button { Text = "Cancel", Location = new Point(225, 43), Size = new Size(70, 30), DialogResult = DialogResult.Cancel };

        dlg.Controls.AddRange(new Control[] { lbl, txt, btnOk, btnCancel });
        dlg.AcceptButton = btnOk;
        dlg.CancelButton = btnCancel;

        if (dlg.ShowDialog(this) == DialogResult.OK)
        {
            input = txt.Text.Trim().PadLeft(6);
        }
        else return;

        // Search in grid
        for (int i = 0; i < dtResult.Rows.Count; i++)
        {
            string bId = dtResult.Rows[i]["Batch"]?.ToString()?.Trim() ?? "";
            if (bId == input.Trim())
            {
                grid.ClearSelection();
                grid.Rows[i].Selected = true;
                grid.FirstDisplayedScrollingRowIndex = i;
                statusLabel.Text = $"Found batch {input.Trim()} at row {i + 1}";
                return;
            }
        }

        MessageBox.Show($"Batch {input.Trim()} not found in this process!", "Not Found",
            MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    // === Totals (Ctrl+T) — sum Wafers, Strips, Pcs ===
    private void ShowTotals()
    {
        if (dtResult == null || dtResult.Rows.Count == 0) return;

        long sumW = 0, sumS = 0, sumP = 0;
        int count = 0;

        foreach (DataRow row in dtResult.Rows)
        {
            count++;
            if (row["Wafers"] != DBNull.Value) sumW += Convert.ToInt64(row["Wafers"]);
            if (row["Strips"] != DBNull.Value) sumS += Convert.ToInt64(row["Strips"]);
            if (row["Pcs"] != DBNull.Value)    sumP += Convert.ToInt64(row["Pcs"]);
        }

        MessageBox.Show(
            $"Process: {currentProcessId.Trim()} — {lblProcessName.Text}\n" +
            $"━━━━━━━━━━━━━━━━━━━━━━━━\n" +
            $"  Batches:  {count:N0}\n" +
            $"  Wafers:   {sumW:N0}\n" +
            $"  Strips:   {sumS:N0}\n" +
            $"  Pieces:   {sumP:N0}\n",
            "WIP Totals",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information);
    }

    // === Keyboard shortcuts ===
    private void OnFormKeyDown(object? sender, KeyEventArgs e)
    {
        switch (e.KeyCode)
        {
            case Keys.Escape:
                Close();
                e.Handled = true;
                break;
            case Keys.F4:
                SearchBatch();
                e.Handled = true;
                break;
            case Keys.T when e.Control:
                ShowTotals();
                e.Handled = true;
                break;
            case Keys.S when e.Alt:
                CycleSortOrder();
                e.Handled = true;
                break;
            case Keys.Enter when cmbProcess.Focused:
                ExecuteQuery();
                e.Handled = true;
                break;
        }
    }
}

// Extension to enable DoubleBuffered on DataGridView
static class DataGridViewExtensions
{
    public static DataGridView DoubleBuffered_(this DataGridView dgv)
    {
        typeof(DataGridView)
            .GetProperty("DoubleBuffered", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)
            ?.SetValue(dgv, true);
        return dgv;
    }
}
