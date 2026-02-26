using System.Data;
using Advantage.Data.Provider;

namespace PdcGui;

/// <summary>
/// WIP by Process query screen — migrated from Clipper STOKPROC.PRG
/// Uses ADS SQL with CDX index field ordering, matching original Clipper sort/filter behavior.
/// Slack = B_dprom - (Today + LTime_NoRoute) — uses m_linemv + c_leadt for lead times.
/// Read-only view. Data source: D_LINE.DBF + C_PROC.DBF + m_linemv.DBF + c_leadt.DBF via ADS.
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
    private bool hasLeadTimeData = false; // true if m_linemv + c_leadt tables available

    // Sort modes matching Clipper CDX index tags
    // [0]=tag name (display), [1]=display label, [2]=SQL ORDER BY clause
    // CDX indexes have FOR !(b_stat$"CDMT") — replicated via UPPER(B_stat) NOT IN ('C','D','M','T')
    // ISLACK:   stored field 'islack' in DBF — ORDER BY islack
    // QCISLACK: stored field 'qcislack' in DBF — ORDER BY qcislack
    // NWISLACK: computed key cpproc_id+ptype_id+pline_id+size_id+b_prior+STR(slack)+purp
    // NWISLACK CDX key: cpproc_id+ptype_id+pline_id+size_id+b_prior
    //   +STR(100000+(b_dprom-ExpFindate),6)
    //   +IF(b_purp=CHR(55),CHR(50),IF(b_purp=CHR(54),CHR(51),b_purp))
    // CHR(55)='7'→CHR(50)='2', CHR(54)='6'→CHR(51)='3' — remaps purp for sort order
    // ADS SQL: use CASE instead of IF/CHR, (b_dprom - ExpFinDate) for slack sort
    private readonly string[][] sortModes = new[]
    {
        new[] { "ISLACK",   "Prior+Slack+Purp+Stat+B/N", "islack" },
        new[] { "NWISLACK", "Type+Line+Size+Prior+Slack+Purp",
                "ptype_id, pline_id, size_id, b_prior, (b_dprom - ExpFinDate), " +
                "CASE WHEN b_purp='7' THEN '2' WHEN b_purp='6' THEN '3' ELSE b_purp END, b_id" },
    };
    private readonly string[][] sortModes190 = new[]
    {
        new[] { "QCISLACK", "Stat+Prior+Slack+Purp+B/N", "qcislack" },
        new[] { "ISLACK",   "Prior+Slack+Purp+Stat+B/N", "islack" },
        new[] { "NWISLACK", "Type+Line+Size+Prior+Slack+Purp",
                "ptype_id, pline_id, size_id, b_prior, (b_dprom - ExpFinDate), " +
                "CASE WHEN b_purp='7' THEN '2' WHEN b_purp='6' THEN '3' ELSE b_purp END, b_id" },
    };
    // Process 001.0 (Metal) — single sort mode: IMSLACK
    // Clipper: aIndex := {"IMSLACK"} — label "Stat+Prior+Slack+Purp+B/N(METAL)"
    private readonly string[][] sortModes001 = new[]
    {
        new[] { "IMSLACK", "Stat+Prior+Slack+Purp+B/N(METAL)", "imslack" },
    };
    private int currentSortIndex = 0;

    /// <summary>
    /// Select sort mode array based on process ID (Clipper: STOKPROC.PRG lines 34-43).
    /// 190.0 → QCISLACK/ISLACK/NWISLACK, 001.0 → IMSLACK only, other → ISLACK/NWISLACK
    /// </summary>
    private string[][] GetSortModes(string procId)
    {
        string p = procId.Trim();
        if (p == "190.0") return sortModes190;
        if (p == "001.0") return sortModes001;
        return sortModes;
    }

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
            Width = 280,
            DropDownWidth = 320,
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
            Location = new Point(375, 15),
            Font = new Font("Segoe UI", 10f),
            ForeColor = Color.DarkBlue,
        };

        btnQuery = new Button
        {
            Text = "Show WIP",
            Location = new Point(600, 10),
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
            Location = new Point(700, 10),
            Size = new Size(110, 30),
            FlatStyle = FlatStyle.Flat,
            Enabled = false,
        };
        btnSearch.Click += (s, e) => SearchBatch();

        btnTotals = new Button
        {
            Text = "Totals (Ctrl+T)",
            Location = new Point(820, 10),
            Size = new Size(110, 30),
            FlatStyle = FlatStyle.Flat,
            Enabled = false,
        };
        btnTotals.Click += (s, e) => ShowTotals();

        lblIndexInfo = new Label
        {
            Text = "",
            AutoSize = true,
            Location = new Point(940, 15),
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
                string procName = row["proc_nme"]?.ToString()?.Trim() ?? "";
                if (!string.IsNullOrWhiteSpace(procId))
                    cmbProcess.Items.Add(string.IsNullOrEmpty(procName) ? procId : $"{procId} - {procName}");
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
        // Extract ID from "ID - Description" format in ComboBox
        string rawText = cmbProcess.Text.Trim();
        string procId = rawText.Contains(" - ") ? rawText.Split(" - ")[0].Trim() : rawText;
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

            // Query D_LINE via index seek
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

    /// <summary>
    /// Load WIP data via ADS SQL, replicating CDX index behavior:
    ///   ordsetfocus("ISLACK")   →  ORDER BY islack  (stored field = CDX key)
    ///   ordsetfocus("QCISLACK") →  ORDER BY qcislack
    ///   ordsetfocus("NWISLACK") →  ORDER BY (computed expression matching CDX key)
    ///   dbseek(cProc)           →  WHERE CpProc_id = :proc
    ///   FOR !(b_stat$"CDMT")    →  AND UPPER(B_stat) NOT IN ('C','D','M','T')
    /// Note: AdsExtendedReader ISAM Seek/SetRange fails with Error 7038 on conditional
    /// CDX indexes, so we use SQL with the stored index fields for equivalent results.
    /// </summary>
    private void LoadWipData(string procId)
    {
        string connStr = $"Data Source={dataDir};ServerType=LOCAL;TableType=CDX;LockMode=COMPATIBLE;CharType=OEM;TrimTrailingSpaces=TRUE;";
        using var conn = new AdsConnection(connStr);
        conn.Open();

        var sortModeArray = GetSortModes(procId);
        string indexTag = sortModeArray[currentSortIndex][0];
        string orderBy = sortModeArray[currentSortIndex][2];

        // SQL query: filter by process + exclude CDMT statuses (matching CDX FOR condition)
        // ORDER BY the index field for correct sort order
        string sql = $"SELECT * FROM \"D_LINE.DBF\" WHERE CpProc_id = '{procId}' " +
                     $"AND UPPER(B_stat) NOT IN ('C','D','M','T') ORDER BY {orderBy}";

        using var cmd = new AdsCommand(sql, conn);
        using var reader = cmd.ExecuteReader();

        dtResult = BuildDataTable();
        int rowNum = 0;

        // Collect batch IDs first for bulk lead time lookup
        var batchRows = new List<(DataRow row, string batchId, DateTime? promDate)>();

        while (reader.Read())
        {
            rowNum++;
            var row = dtResult.NewRow();
            row["#"] = rowNum;
            string batchId = reader["B_id"]?.ToString()?.Trim() ?? "";
            row["Batch"] = batchId;
            row["Pu"] = reader["B_purp"]?.ToString()?.Trim() ?? "";
            row["St"] = reader["B_stat"]?.ToString()?.Trim() ?? "";
            row["P_Line"] = reader["pLine_id"]?.ToString()?.Trim() ?? "";
            row["Bt"] = reader["B_type"]?.ToString()?.Trim() ?? "";
            row["Size"] = reader["Size_id"]?.ToString()?.Trim() ?? "";
            try { row["Val"] = Convert.ToDecimal(reader["Value_id"]); }
            catch { row["Val"] = 0m; }

            // Numeric fields
            row["Wafers"] = GetInt(reader, "Cp_bQtyW");
            row["Strips"] = GetInt(reader, "Cp_bQtyS");
            row["Pcs"] = GetInt(reader, "Cp_bQtyP");

            // Date fields
            var arrDate = GetDate(reader, "Cp_dArr");
            row["Arrive_Date"] = arrDate.HasValue ? arrDate.Value : DBNull.Value;
            if (arrDate.HasValue)
                row["Days_In_Proc"] = (DateTime.Today - arrDate.Value).Days;

            var promDate = GetDate(reader, "B_dprom");
            // Slack will be calculated after lead time lookup (below)

            // Priority = B_prior + B_wkprom (concatenated like Clipper)
            string prior = reader["B_prior"]?.ToString()?.Trim() ?? "";
            string wkprom = reader["B_wkprom"]?.ToString()?.Trim() ?? "";
            row["Priority"] = prior + wkprom;

            row["Delivered_By"] = reader["Pp_widFin"]?.ToString()?.Trim() ?? "";
            row["Received_By"] = reader["Cp_widArr"]?.ToString()?.Trim() ?? "";
            row["Comments"] = reader["B_remark"]?.ToString()?.Trim() ?? "";

            // Hidden columns for formatting
            var startedDate = GetDate(reader, "Cp_dSta");
            row["Started"] = startedDate.HasValue ? startedDate.Value : DBNull.Value;

            dtResult.Rows.Add(row);
            batchRows.Add((row, batchId, promDate));
        }

        // === Calculate Slack with LTime_NoRoute (Clipper: DELIVSCH.PRG) ===
        hasLeadTimeData = SlackCalculator.CalculateSlack(conn, batchRows);

        if (dtResult.Rows.Count == 0)
        {
            MessageBox.Show("No active records found for this process!", "Empty",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            grid.DataSource = null;
            countLabel.Text = "Records: 0";
            statusLabel.Text = "No records found";
            return;
        }

        // Bind to grid
        grid.DataSource = null;
        grid.DataSource = dtResult;
        ConfigureGridColumns();

        countLabel.Text = $"Records: {dtResult.Rows.Count}";
        string ltimeInfo = hasLeadTimeData ? "" : " [Slack: no lead time data]";
        statusLabel.Text = $"WIP for process {procId.Trim()} — {lblProcessName.Text} — {dtResult.Rows.Count} active batches [{indexTag}]{ltimeInfo}";
    }

    private static DataTable BuildDataTable()
    {
        var dt = new DataTable();
        dt.Columns.Add("#", typeof(int));
        dt.Columns.Add("Batch", typeof(string));
        dt.Columns.Add("Pu", typeof(string));
        dt.Columns.Add("St", typeof(string));
        dt.Columns.Add("P_Line", typeof(string));
        dt.Columns.Add("Bt", typeof(string));
        dt.Columns.Add("Size", typeof(string));
        dt.Columns.Add("Val", typeof(decimal));
        dt.Columns.Add("Wafers", typeof(int));
        dt.Columns.Add("Strips", typeof(int));
        dt.Columns.Add("Pcs", typeof(int));
        dt.Columns.Add("Arrive_Date", typeof(DateTime));
        dt.Columns.Add("Days_In_Proc", typeof(int));
        dt.Columns.Add("Priority", typeof(string));
        dt.Columns.Add("Slack", typeof(int));
        dt.Columns.Add("Delivered_By", typeof(string));
        dt.Columns.Add("Received_By", typeof(string));
        dt.Columns.Add("Comments", typeof(string));
        dt.Columns.Add("Started", typeof(DateTime)); // hidden, for color coding
        return dt;
    }

    private static int GetInt(System.Data.IDataReader reader, string field)
    {
        try
        {
            var val = reader[field];
            if (val == null || val == DBNull.Value) return 0;
            return Convert.ToInt32(val);
        }
        catch { return 0; }
    }

    private static DateTime? GetDate(System.Data.IDataReader reader, string field)
    {
        try
        {
            var val = reader[field];
            if (val == null || val == DBNull.Value) return null;
            if (val is DateTime dt && dt != DateTime.MinValue) return dt;
            return null;
        }
        catch { return null; }
    }

    private void ConfigureGridColumns()
    {
        if (grid.Columns.Count == 0) return;

        // Define display order matching Clipper English layout
        var displayOrder = new[]
        {
            ("#",             40),
            ("Batch",         75),
            ("Pu",            30),
            ("St",            30),
            ("P_Line",        55),
            ("Bt",            30),
            ("Size",          50),
            ("Val",           80),
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
        if (grid.Columns.Contains("Started"))
            grid.Columns["Started"]!.Visible = false;

        int order = 0;
        foreach (var (name, width) in displayOrder)
        {
            if (grid.Columns.Contains(name))
            {
                grid.Columns[name]!.DisplayIndex = order;
                grid.Columns[name]!.Width = width;
                order++;

                if (name is "Wafers" or "Strips" or "Pcs" or "Days_In_Proc" or "Slack" or "#" or "Val")
                    grid.Columns[name]!.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            }
        }

        // Number formatting matching Harbour Transform patterns
        if (grid.Columns.Contains("Val"))
            grid.Columns["Val"]!.DefaultCellStyle.Format = "N3";     // Clipper: Value_id (decimal 9,3)
        if (grid.Columns.Contains("Strips"))
            grid.Columns["Strips"]!.DefaultCellStyle.Format = "N0";  // Clipper: Transform(,"99,999")
        if (grid.Columns.Contains("Pcs"))
            grid.Columns["Pcs"]!.DefaultCellStyle.Format = "N0";     // Clipper: Transform(,"9,999,999")
    }

    // === Cell formatting: green background for started batches (Cp_dSta not empty) ===
    // Clipper: colorSpec "W+/B,W+/R,B/Bg", colorBlock {||Iif(EMPTY(Cp_dSta),{1,2},{3,2})}
    // B/Bg = Black text on Green background for started batches
    private void OnCellFormatting(object? sender, DataGridViewCellFormattingEventArgs e)
    {
        if (dtResult == null || e.RowIndex < 0 || e.RowIndex >= dtResult.Rows.Count) return;

        var row = dtResult.Rows[e.RowIndex];
        var started = row["Started"];

        // Clipper: colorBlock checks EMPTY(D_line->Cp_dSta) — started batches get green background
        if (started != DBNull.Value && started is DateTime dt && dt != DateTime.MinValue)
        {
            e.CellStyle!.ForeColor = Color.Black;
            e.CellStyle!.BackColor = Color.FromArgb(144, 238, 144); // LightGreen — Clipper B/Bg
        }

        // Color negative slack red
        if (grid.Columns[e.ColumnIndex].Name == "Slack" && e.Value is int slack && slack < 0)
        {
            e.CellStyle!.ForeColor = Color.Red;
            e.CellStyle!.Font = new Font(grid.Font, FontStyle.Bold);
        }
    }

    // === Alt+S sort toggle via column header click ===
    private void OnColumnHeaderClick(object? sender, DataGridViewCellMouseEventArgs e)
    {
        CycleSortOrder();
    }

    private void CycleSortOrder()
    {
        if (string.IsNullOrEmpty(currentProcessId) || dtResult == null) return;

        var sortModeArray = GetSortModes(currentProcessId);
        currentSortIndex = (currentSortIndex + 1) % sortModeArray.Length;

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
        var sortModeArray = GetSortModes(currentProcessId);
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
            input = txt.Text.Trim().PadLeft(6);
        else
            return;

        // Search in grid data
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
