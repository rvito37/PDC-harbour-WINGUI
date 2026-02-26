using System.Data;
using Advantage.Data.Provider;

namespace PdcGui;

/// <summary>
/// WIP by Workstation query screen — migrated from Clipper STOKWORK.PRG
/// Uses ADS SQL, matching original Clipper filter/sort behavior.
/// Supports individual workstations (CPWKSTN_ID) and GRU0-GRU7 process groups.
/// Read-only view. Data source: D_LINE.DBF + C_PROC.DBF + C_WKSTN.DBF via ADS.
/// </summary>
public class StockInWorkForm : Form
{
    private readonly string dataDir;

    // UI controls
    private Panel topPanel = null!;
    private ComboBox cmbWorkstation = null!;
    private Label lblWorkstationName = null!;
    private Button btnQuery = null!;
    private DataGridView grid = null!;
    private Label countLabel = null!;
    private Label statusLabel = null!;

    // Data
    private DataTable? dtWorkstations;
    private DataTable? dtResult;
    private Dictionary<string, string>? procNameCache;
    private string currentWorkstationId = "";
    private bool hasLeadTimeData = false;

    // GRU group definitions (Clipper STOKWORK.PRG lines 48-67, 114-137)
    // Each group maps to a set of process IDs from CDX index FOR conditions (indexes.csv)
    private static readonly Dictionary<string, (string name, string[] procs)> GruGroups = new()
    {
        ["GRU0"] = ("Substrate Cleaning", new[] { "200.0", "300.0", "400.0", "500.0", "600.0" }),
        ["GRU1"] = ("Lito + Neg + Pos", Array.Empty<string>()), // special handling: left(3) + right(1)
        ["GRU2"] = ("Polymid", new[] { "420.0","423.0","426.0","429.0","520.0","523.0","526.0","538.0","639.0","644.0","626.0","523.1","523.2","538.3","538.4","626.1" }),
        ["GRU3"] = ("Chemical", new[] { "203.0","303.0","325.0","413.0","462.0","417.0","437.0","419.0","506.0","507.0","508.0","506.1","507.1","508.1","442.0","433.0","403.0","438.0","501.2","529.0","608.0","608.2","608.1","647.0","609.0","655.0","654.0","631.0","638.0","275.3","632.0","625.0","661.0","617.0","621.0" }),
        ["GRU4"] = ("Electroplating", new[] { "324.0","430.0","436.0","532.0","541.0","432.0","441.0","530.0","531.0","536.0","431.0","341.0","649.0","653.0","624.0","629.0","323.0","461.0","462.0","532.6","543.1","541.1","541.3","137.1","137.2","137.0" }),
        ["GRU5"] = ("Cure", new[] { "421.0","425.0","428.0","439.0","446.0","521.0","525.0","527.0","539.0","645.0","627.0","648.0","422.0","424.0","427.0","444.0","522.0","524.0","528.0","544.0","404.0","566.1","164.4","541.2","579.1","525.1","577.3","539.1","539.4","627.1" }),
        ["GRU6"] = ("Measurements", new[] { "212.0","212.1","215.0","216.0","311.0","213.0","443.0","445.0","405.0","406.0","512.0","513.0","515.0","516.0","409.0","414.0","622.0","650.0","571.6","571.0","622.1","650.1" }),
        ["GRU7"] = ("ET-6", new[] { "305.0","307.0","422.0","424.0","522.0","524.0","533.0","533.1","544.0","545.0","559.0","603.0","605.0","605.1","605.2","606.0","607.0","607.1","613.0","616.0","628.0","631.0","637.0","646.0","660.0","662.0","415.0","528.4","657.0","407.0","434.0","524.2","643.0","667.0","528.0" }),
    };

    // GRU1 special: left(CPPROC_ID,3) $ "224_227_..." AND right(cpproc_id,1) $ "01234"
    private static readonly string[] Gru1Prefixes = new[]
    {
        "224","227","221","231","217","310","320","323","315","317",
        "461","411","416","418","435","509","511","505","514","517",
        "518","519","610","612","659","615","618","619","652","630",
        "633","326","510","543","537","312","623"
    };
    private static readonly string Gru1Suffixes = "01234";

    public StockInWorkForm(string dataDirectory)
    {
        dataDir = dataDirectory;
        InitializeUI();
        LoadWorkstationList();
    }

    private void InitializeUI()
    {
        Text = "WIP by Workstation — Workstation WIP Query";
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

        var lblWkstn = new Label
        {
            Text = "Workstation:",
            AutoSize = true,
            Location = new Point(10, 15),
            Font = new Font("Segoe UI", 10f, FontStyle.Bold),
        };

        cmbWorkstation = new ComboBox
        {
            Location = new Point(115, 12),
            Width = 90,
            DropDownStyle = ComboBoxStyle.DropDown,
            Font = new Font("Consolas", 11f),
            // CharacterCasing not available in WinForms ComboBox; use TextChanged to uppercase
        };
        cmbWorkstation.KeyDown += (s, e) =>
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                ExecuteQuery();
            }
        };
        // Clipper PICTURE "!!!!" — force uppercase input
        cmbWorkstation.TextChanged += (s, e) =>
        {
            if (cmbWorkstation.Text != cmbWorkstation.Text.ToUpper())
            {
                int pos = cmbWorkstation.SelectionStart;
                cmbWorkstation.Text = cmbWorkstation.Text.ToUpper();
                cmbWorkstation.SelectionStart = pos;
            }
        };

        lblWorkstationName = new Label
        {
            Text = "",
            AutoSize = true,
            Location = new Point(215, 15),
            Font = new Font("Segoe UI", 10f),
            ForeColor = Color.DarkBlue,
        };

        btnQuery = new Button
        {
            Text = "Show WIP",
            Location = new Point(500, 10),
            Size = new Size(90, 30),
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(0, 120, 215),
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9f, FontStyle.Bold),
        };
        btnQuery.Click += (s, e) => ExecuteQuery();

        topPanel.Controls.AddRange(new Control[] { lblWkstn, cmbWorkstation, lblWorkstationName, btnQuery });
        Controls.Add(topPanel);

        // === Grid ===
        grid = new DataGridView
        {
            Dock = DockStyle.Fill,
            ReadOnly = true,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            AllowUserToResizeRows = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
            RowHeadersVisible = false,
            BackgroundColor = Color.White,
            BorderStyle = BorderStyle.None,
            Font = new Font("Consolas", 9.5f),
            // DoubleBuffered set via reflection for smoother rendering
        };
        grid.CellFormatting += OnCellFormatting;
        // Enable DoubleBuffered via reflection (protected property)
        typeof(DataGridView).InvokeMember("DoubleBuffered",
            System.Reflection.BindingFlags.SetProperty | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic,
            null, grid, new object[] { true });
        Controls.Add(grid);

        // === Status Bar ===
        var statusPanel = new Panel
        {
            Dock = DockStyle.Bottom,
            Height = 25,
            BackColor = Color.FromArgb(230, 230, 240),
        };
        countLabel = new Label
        {
            Text = "Records: 0",
            AutoSize = true,
            Location = new Point(10, 4),
            Font = new Font("Segoe UI", 9f),
        };
        statusLabel = new Label
        {
            Text = "Select a workstation and press Enter or Show WIP",
            AutoSize = true,
            Location = new Point(200, 4),
            Font = new Font("Segoe UI", 9f),
            ForeColor = Color.DarkSlateGray,
        };
        statusPanel.Controls.AddRange(new Control[] { countLabel, statusLabel });
        Controls.Add(statusPanel);
    }

    private void OnFormKeyDown(object? sender, KeyEventArgs e)
    {
        if (e.KeyCode == Keys.Escape)
        {
            Close();
            e.Handled = true;
        }
    }

    // === Load workstation list from C_WKSTN.DBF ===
    private void LoadWorkstationList()
    {
        try
        {
            string connStr = $"Data Source={dataDir};ServerType=LOCAL;TableType=CDX;LockMode=COMPATIBLE;CharType=OEM;TrimTrailingSpaces=TRUE;";
            using var conn = new AdsConnection(connStr);
            conn.Open();

            string sql = "SELECT wkstn_id, wkstn_nme FROM \"C_WKSTN.DBF\" ORDER BY wkstn_id";
            using var cmd = new AdsCommand(sql, conn);
            using var adapter = new AdsDataAdapter(cmd);
            dtWorkstations = new DataTable();
            adapter.Fill(dtWorkstations);

            cmbWorkstation.Items.Clear();
            foreach (DataRow row in dtWorkstations.Rows)
            {
                string id = row["wkstn_id"]?.ToString()?.Trim() ?? "";
                if (!string.IsNullOrEmpty(id))
                    cmbWorkstation.Items.Add(id);
            }

            // Add GRU groups at the end
            cmbWorkstation.Items.Add("---");
            foreach (var gru in GruGroups)
                cmbWorkstation.Items.Add($"{gru.Key}");

            if (cmbWorkstation.Items.Count > 0)
                cmbWorkstation.SelectedIndex = 0;
        }
        catch (Exception ex)
        {
            statusLabel.Text = $"Cannot load workstations: {ex.Message}";
        }
    }

    // === Load process name cache for JOIN ===
    private void LoadProcNames(AdsConnection conn)
    {
        if (procNameCache != null) return;
        procNameCache = new Dictionary<string, string>();
        try
        {
            string sql = "SELECT proc_id, PROC_NMH FROM \"C_PROC.DBF\"";
            using var cmd = new AdsCommand(sql, conn);
            using var reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                string id = reader["proc_id"]?.ToString() ?? "";
                string name = reader["PROC_NMH"]?.ToString()?.Trim() ?? "";
                procNameCache.TryAdd(id, name);
            }
        }
        catch { }
    }

    // === Main query execution ===
    private void ExecuteQuery()
    {
        string wkstnId = cmbWorkstation.Text.Trim().ToUpper();
        if (string.IsNullOrWhiteSpace(wkstnId) || wkstnId == "---")
        {
            MessageBox.Show("Please enter a Workstation ID.", "Input Required",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            cmbWorkstation.Focus();
            return;
        }

        try
        {
            Cursor = Cursors.WaitCursor;
            statusLabel.Text = $"Querying WIP for workstation {wkstnId}...";
            Application.DoEvents();

            // Validate workstation and get display name
            string wkstnName = ValidateWorkstation(wkstnId);
            if (wkstnName == null!)
            {
                Cursor = Cursors.Default;
                return;
            }
            lblWorkstationName.Text = wkstnName;
            currentWorkstationId = wkstnId;

            // Query D_LINE
            LoadWipData(wkstnId);
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

    /// <summary>
    /// Validate workstation ID against C_WKSTN or GRU groups.
    /// Clipper: postWork() in STOKWORK.PRG lines 100-144
    /// </summary>
    private string ValidateWorkstation(string wkstnId)
    {
        // Check GRU groups first
        if (GruGroups.TryGetValue(wkstnId, out var gru))
            return gru.name;

        // Check C_WKSTN table
        if (dtWorkstations != null)
        {
            foreach (DataRow row in dtWorkstations.Rows)
            {
                string id = row["wkstn_id"]?.ToString()?.Trim() ?? "";
                if (id.Equals(wkstnId, StringComparison.OrdinalIgnoreCase))
                    return row["wkstn_nme"]?.ToString()?.Trim() ?? id;
            }
        }

        MessageBox.Show($"Workstation '{wkstnId}' not found.", "Not Found",
            MessageBoxButtons.OK, MessageBoxIcon.Warning);
        return null!;
    }

    /// <summary>
    /// Build SQL WHERE clause based on workstation selection.
    /// Individual workstation: filter by CPWKSTN_ID
    /// GRU groups: filter by CpProc_id IN (process list)
    /// All include: UPPER(B_stat) NOT IN ('C','D','M','T')
    /// </summary>
    private string BuildWhereClause(string wkstnId)
    {
        string statusFilter = "UPPER(B_stat) NOT IN ('C','D','M','T')";

        if (wkstnId == "GRU1")
        {
            // Special: left(CPPROC_ID,3) matches prefix list AND right(cpproc_id,1) in "01234"
            var prefixConditions = string.Join(" OR ",
                Gru1Prefixes.Select(p => $"LEFT(CpProc_id,3) = '{p}'"));
            var suffixCondition = string.Join(",",
                Gru1Suffixes.Select(c => $"'{c}'"));
            return $"({prefixConditions}) AND RIGHT(CpProc_id,1) IN ({suffixCondition}) AND {statusFilter}";
        }

        if (GruGroups.TryGetValue(wkstnId, out var gru) && gru.procs.Length > 0)
        {
            var procList = string.Join(",", gru.procs.Select(p => $"'{p}'"));
            return $"CpProc_id IN ({procList}) AND {statusFilter}";
        }

        // Individual workstation — pad to 4 chars to match DBF field width
        string paddedId = wkstnId.PadRight(4);
        return $"CPWKSTN_ID = '{paddedId}' AND {statusFilter}";
    }

    /// <summary>
    /// Load WIP data from D_LINE.DBF filtered by workstation/group.
    /// Columns match Clipper STOKWORK.PRG CreateBro() lines 204-251.
    /// Sort by islack (matching IWSLACK/iwSlack_N CDX key expression).
    /// </summary>
    private void LoadWipData(string wkstnId)
    {
        string connStr = $"Data Source={dataDir};ServerType=LOCAL;TableType=CDX;LockMode=COMPATIBLE;CharType=OEM;TrimTrailingSpaces=TRUE;";
        using var conn = new AdsConnection(connStr);
        conn.Open();

        // Load process names for Proc Name column (Clipper: Set Relation to cpproc_id into c_proc)
        LoadProcNames(conn);

        string where = BuildWhereClause(wkstnId);
        string sql = $"SELECT * FROM \"D_LINE.DBF\" WHERE {where} ORDER BY islack";

        using var cmd = new AdsCommand(sql, conn);
        using var reader = cmd.ExecuteReader();

        dtResult = BuildDataTable();
        int rowNum = 0;

        var batchRows = new List<(DataRow row, string batchId, DateTime? promDate)>();

        while (reader.Read())
        {
            rowNum++;
            var row = dtResult.NewRow();
            row["#"] = rowNum;

            string batchId = reader["B_id"]?.ToString()?.Trim() ?? "";
            row["Pu"] = reader["b_purp"]?.ToString()?.Trim() ?? "";
            row["Esn"] = reader["esn_id"]?.ToString()?.Trim() ?? "";
            row["Batch"] = batchId;

            // Wfr = cp_bqtyw (Transform "999")
            row["Wfr"] = GetInt(reader, "cp_bqtyw");

            // Pcs = cp_bqtyp (Transform "9,999,999")
            row["Pcs"] = GetInt(reader, "cp_bqtyp");

            // Proc + Proc Name (via c_proc relation)
            string procId = reader["cpproc_id"]?.ToString() ?? "";
            row["Proc"] = procId.Trim();
            if (procNameCache != null && procNameCache.TryGetValue(procId, out var procName))
                row["Proc_Name"] = procName;
            else
                row["Proc_Name"] = "";

            row["St"] = reader["B_stat"]?.ToString()?.Trim() ?? "";
            row["Pr"] = reader["B_prior"]?.ToString()?.Trim() ?? "";

            // Days = Date() - cp_darr
            var arrDate = GetDate(reader, "cp_darr");
            if (arrDate.HasValue)
                row["Days"] = (DateTime.Today - arrDate.Value).Days;

            // Slack will be calculated after lead time lookup
            var promDate = GetDate(reader, "B_dprom");

            row["Comments"] = reader["B_remark"]?.ToString()?.Trim() ?? "";

            // Hidden: Started date for green highlighting
            var startedDate = GetDate(reader, "Cp_dSta");
            row["Started"] = startedDate.HasValue ? startedDate.Value : DBNull.Value;

            dtResult.Rows.Add(row);
            batchRows.Add((row, batchId, promDate));
        }

        // === Calculate Slack with LTime_NoRoute (Clipper: DELIVSCH.PRG) ===
        hasLeadTimeData = SlackCalculator.CalculateSlack(conn, batchRows);

        if (dtResult.Rows.Count == 0)
        {
            MessageBox.Show("No active records found for this workstation!", "Empty",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            grid.DataSource = null;
            countLabel.Text = "Records: 0";
            statusLabel.Text = "No records found";
            return;
        }

        grid.DataSource = dtResult;
        ConfigureGridColumns();

        countLabel.Text = $"Records: {dtResult.Rows.Count}";
        string ltimeInfo = hasLeadTimeData ? "" : " [Slack: no lead time data]";
        statusLabel.Text = $"WIP for workstation {currentWorkstationId} — {lblWorkstationName.Text} — {dtResult.Rows.Count} active batches{ltimeInfo}";
    }

    private static DataTable BuildDataTable()
    {
        var dt = new DataTable();
        dt.Columns.Add("#", typeof(int));
        dt.Columns.Add("Pu", typeof(string));
        dt.Columns.Add("Esn", typeof(string));
        dt.Columns.Add("Batch", typeof(string));
        dt.Columns.Add("Wfr", typeof(int));
        dt.Columns.Add("Pcs", typeof(int));
        dt.Columns.Add("Proc", typeof(string));
        dt.Columns.Add("Proc_Name", typeof(string));
        dt.Columns.Add("St", typeof(string));
        dt.Columns.Add("Pr", typeof(string));
        dt.Columns.Add("Days", typeof(int));
        dt.Columns.Add("Slack", typeof(int));
        dt.Columns.Add("Comments", typeof(string));
        dt.Columns.Add("Started", typeof(DateTime)); // hidden, for green formatting
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

        // Display order matching Clipper STOKWORK.PRG CreateBro() lines 204-251
        var displayOrder = new[]
        {
            ("#",          35),
            ("Pu",         30),
            ("Esn",        55),
            ("Batch",      75),
            ("Wfr",        45),
            ("Pcs",        85),
            ("Proc",       55),
            ("Proc_Name", 140),
            ("St",         30),
            ("Pr",         30),
            ("Days",       55),
            ("Slack",      55),
            ("Comments",  200),
        };

        int idx = 0;
        foreach (var (name, width) in displayOrder)
        {
            if (grid.Columns.Contains(name))
            {
                grid.Columns[name]!.DisplayIndex = idx;
                grid.Columns[name]!.Width = width;
                idx++;
            }
        }

        // Hide internal columns
        if (grid.Columns.Contains("Started"))
            grid.Columns["Started"]!.Visible = false;

        // Number formatting
        if (grid.Columns.Contains("Wfr"))
            grid.Columns["Wfr"]!.DefaultCellStyle.Format = "N0";
        if (grid.Columns.Contains("Pcs"))
            grid.Columns["Pcs"]!.DefaultCellStyle.Format = "N0";

        // Right-align numeric columns
        foreach (var col in new[] { "#", "Wfr", "Pcs", "Days", "Slack" })
        {
            if (grid.Columns.Contains(col))
                grid.Columns[col]!.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
        }
    }

    // === Cell formatting: green background for started batches (Cp_dSta not empty) ===
    // Clipper: colorSpec "W+/B,W+/R,B/Bg", colorBlock {||Iif(EMPTY(Cp_dSta),{1,2},{3,2})}
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
}
