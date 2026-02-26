using System.Data;
using Advantage.Data.Provider;

namespace PdcGui;

/// <summary>
/// Batch Path Query screen — migrated from Clipper BNPATH.PRG
/// Shows all movement steps for a single batch from m_linemv.DBF.
/// Read-only view. Data source: D_LINE.DBF + m_linemv.DBF + C_PROC.DBF
///   + c_bpurp.DBF + c_btype.DBF + m_loc.DBF via ADS.
/// </summary>
public class BatchPathForm : Form
{
    private readonly string dataDir;

    // Controls — top panel
    private TextBox txtBatchNumber = null!;
    private Button btnQuery = null!;

    // Controls — info panel
    private Panel infoPanel = null!;
    private Label lblPurposeType = null!;
    private Label lblBatchLoc = null!;

    // Controls — grid and status
    private DataGridView grid = null!;
    private StatusStrip statusStrip = null!;
    private ToolStripStatusLabel statusLabel = null!;
    private ToolStripStatusLabel countLabel = null!;

    // Data
    private DataTable? dtResult;
    private Dictionary<string, string>? procNameCache;

    public BatchPathForm(string dataDirectory)
    {
        dataDir = dataDirectory;
        InitializeUI();
    }

    private string GetConnectionString() =>
        $"Data Source={dataDir};ServerType=LOCAL;TableType=CDX;LockMode=COMPATIBLE;CharType=OEM;TrimTrailingSpaces=TRUE;";

    // ========== UI Setup ==========

    private void InitializeUI()
    {
        Text = "Batch Path Query — Batch Line Movement";
        Size = new Size(1100, 650);
        StartPosition = FormStartPosition.CenterParent;
        Font = new Font("Segoe UI", 9.5f);
        KeyPreview = true;
        KeyDown += OnFormKeyDown;

        // === Top Panel ===
        var topPanel = new Panel
        {
            Dock = DockStyle.Top,
            Height = 50,
            Padding = new Padding(8, 8, 8, 4),
            BackColor = Color.FromArgb(240, 240, 250),
        };

        var lblBn = new Label
        {
            Text = "B/N:",
            AutoSize = true,
            Location = new Point(10, 15),
            Font = new Font("Segoe UI", 10f, FontStyle.Bold),
        };

        txtBatchNumber = new TextBox
        {
            Location = new Point(55, 12),
            Width = 90,
            MaxLength = 6,
            Font = new Font("Consolas", 12f),
        };
        txtBatchNumber.KeyDown += (s, e) =>
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                ExecuteQuery();
            }
        };
        txtBatchNumber.KeyPress += (s, e) =>
        {
            // Digits only (matching Clipper PICTURE "@K 999999")
            if (!char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar))
                e.Handled = true;
        };

        btnQuery = new Button
        {
            Text = "Query",
            Location = new Point(160, 10),
            Size = new Size(90, 30),
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(0, 120, 215),
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9f, FontStyle.Bold),
        };
        btnQuery.Click += (s, e) => ExecuteQuery();

        topPanel.Controls.AddRange(new Control[] { lblBn, txtBatchNumber, btnQuery });

        // === Info Panel (Purpose/Type + Location) ===
        infoPanel = new Panel
        {
            Dock = DockStyle.Top,
            Height = 50,
            BackColor = Color.FromArgb(0, 51, 102),
            Padding = new Padding(10, 4, 10, 4),
            Visible = false,
        };

        lblPurposeType = new Label
        {
            Text = "",
            AutoSize = true,
            Location = new Point(10, 5),
            Font = new Font("Segoe UI", 10f),
            ForeColor = Color.White,
        };

        lblBatchLoc = new Label
        {
            Text = "",
            AutoSize = true,
            Location = new Point(10, 27),
            Font = new Font("Segoe UI", 10f),
            ForeColor = Color.White,
        };

        infoPanel.Controls.AddRange(new Control[] { lblPurposeType, lblBatchLoc });

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

        // === Status Bar ===
        statusStrip = new StatusStrip();
        statusLabel = new ToolStripStatusLabel("Enter batch number and press Enter or click 'Query'.")
        {
            Spring = true,
            TextAlign = ContentAlignment.MiddleLeft,
        };
        countLabel = new ToolStripStatusLabel("Steps: 0")
        {
            BorderSides = ToolStripStatusLabelBorderSides.Left,
        };
        statusStrip.Items.AddRange(new ToolStripItem[] { statusLabel, countLabel });

        // === Layout (order matters: Bottom first, then Top panels, then Fill) ===
        Controls.Add(grid);
        Controls.Add(infoPanel);
        Controls.Add(topPanel);
        Controls.Add(statusStrip);

        txtBatchNumber.Focus();
    }

    // ========== Main Query ==========

    private void ExecuteQuery()
    {
        // Validate and zero-pad input (Clipper: StrZero(Val(o:varGet()),6))
        string rawText = txtBatchNumber.Text.Trim();
        if (string.IsNullOrWhiteSpace(rawText) || !int.TryParse(rawText, out int batchNum) || batchNum == 0)
        {
            MessageBox.Show("Please enter a valid batch number.", "Input Required",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            txtBatchNumber.Focus();
            return;
        }
        string batchId = batchNum.ToString("D6");
        txtBatchNumber.Text = batchId;

        Cursor = Cursors.WaitCursor;
        statusLabel.Text = $"Querying batch {batchId}...";
        Application.DoEvents();

        try
        {
            using var conn = new AdsConnection(GetConnectionString());
            conn.Open();

            // Step 1: Validate batch exists in D_LINE
            string bPurp = "", bType = "", esnxxId = "", plineId = "";
            int pcCode = 0;

            using (var cmdCheck = new AdsCommand(
                $"SELECT TOP 1 B_id, b_purp, b_type, esnxx_id, pline_id, CP_PCCODE " +
                $"FROM \"D_LINE.DBF\" WHERE B_id = '{batchId}'", conn))
            using (var rdrCheck = cmdCheck.ExecuteReader())
            {
                if (!rdrCheck.Read())
                {
                    MessageBox.Show($"Batch {batchId} not found!", "Not Found",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    statusLabel.Text = $"Batch {batchId} not found.";
                    Cursor = Cursors.Default;
                    txtBatchNumber.Focus();
                    txtBatchNumber.SelectAll();
                    return;
                }

                // Grab D_LINE fields for lookups
                bPurp = rdrCheck["b_purp"]?.ToString()?.Trim() ?? "";
                bType = rdrCheck["b_type"]?.ToString()?.Trim() ?? "";
                esnxxId = rdrCheck["esnxx_id"]?.ToString()?.Trim() ?? "";
                plineId = rdrCheck["pline_id"]?.ToString()?.Trim() ?? "";
                try { pcCode = Convert.ToInt32(rdrCheck["CP_PCCODE"]); } catch { }
            } // cmdCheck + rdrCheck fully disposed here

            // Step 2: Lookup purpose name (c_bpurp)
            string purposeName = LookupPurpose(conn, bPurp);

            // Step 3: Lookup type name (c_btype)
            string typeName = LookupType(conn, bType, esnxxId, plineId);

            // Step 4: Lookup last location (m_loc)
            string locationInfo = LookupLocation(conn, batchId);

            // Step 5: Update info panel
            lblPurposeType.Text = $"Purpose: {purposeName}  Type: {typeName}";
            lblBatchLoc.Text = $"B/N: {batchId}  Loc: {locationInfo}";
            infoPanel.Visible = true;

            // Step 6: Check CP_PCCODE >= 4550 for pack_mem
            if (pcCode >= 4550)
            {
                ShowPackMemo(conn, batchId);
            }

            // Step 7: Load movement path
            LoadMovementData(conn, batchId);

            int count = dtResult?.Rows.Count ?? 0;
            countLabel.Text = $"Steps: {count}";
            statusLabel.Text = count > 0
                ? $"Batch {batchId} — {count} movement steps. ESC to return."
                : $"No movement data found for batch {batchId}.";

            if (count == 0)
            {
                MessageBox.Show($"No information found for this batch!", "Empty",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error querying batch:\n{ex.Message}", "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            statusLabel.Text = "Error.";
        }
        finally
        {
            Cursor = Cursors.Default;
        }
    }

    // ========== Lookup Helpers ==========

    private string LookupPurpose(AdsConnection conn, string bPurp)
    {
        if (string.IsNullOrWhiteSpace(bPurp)) return "";
        try
        {
            string sql = $"SELECT bpurp_nme FROM \"c_bpurp.DBF\" WHERE b_purp = '{bPurp}'";
            using var cmd = new AdsCommand(sql, conn);
            using var rdr = cmd.ExecuteReader();
            if (rdr.Read())
                return rdr["bpurp_nme"]?.ToString()?.Trim() ?? "";
        }
        catch { /* c_bpurp.DBF may not exist */ }
        return "";
    }

    private string LookupType(AdsConnection conn, string bType, string esnxxId, string plineId)
    {
        if (string.IsNullOrWhiteSpace(bType)) return "";
        try
        {
            // c_btype.CDX has alias-qualified index key (c_btype->field) which causes ADS Error 3010
            // in SQL mode. Use DbfDataReader (no ADS/CDX dependency) to read the table directly.
            string tablePath = Path.Combine(dataDir, "c_btype.DBF");
            if (!File.Exists(tablePath)) return "";
            var options = new DbfDataReader.DbfDataReaderOptions { SkipDeletedRecords = true };
            using var rdr = new DbfDataReader.DbfDataReader(tablePath, options);
            while (rdr.Read())
            {
                string bt = rdr["b_type"]?.ToString()?.Trim() ?? "";
                string es = rdr["esnxx_id"]?.ToString()?.Trim() ?? "";
                string pl = rdr["pline_id"]?.ToString()?.Trim() ?? "";
                if (bt == bType && es == esnxxId && pl == plineId)
                    return rdr["btype_nme"]?.ToString()?.Trim() ?? "";
            }
        }
        catch { /* c_btype.DBF may not exist */ }
        return "";
    }

    private string LookupLocation(AdsConnection conn, string batchId)
    {
        // Clipper GetLoc(): seek b_id in m_loc, walk forward to end, skip back 1 → last record
        // SQL: read all rows for b_id, take the last one programmatically (safest for ADS)
        try
        {
            string sql = $"SELECT loc, dadd_rec, tlu_rec FROM \"m_loc.DBF\" WHERE b_id = '{batchId}'";
            using var cmd = new AdsCommand(sql, conn);
            using var rdr = cmd.ExecuteReader();

            string lastLoc = "", lastDate = "", lastTime = "";
            while (rdr.Read())
            {
                lastLoc = rdr["loc"]?.ToString()?.Trim() ?? "";
                try
                {
                    var dateVal = rdr["dadd_rec"];
                    if (dateVal is DateTime dt && dt != DateTime.MinValue)
                        lastDate = dt.ToString("dd/MM/yyyy");
                }
                catch { }
                lastTime = rdr["tlu_rec"]?.ToString()?.Trim() ?? "";
            }

            if (!string.IsNullOrEmpty(lastLoc))
                return $"{lastLoc} #{lastDate} #{lastTime}";
        }
        catch { /* m_loc.DBF may not exist */ }
        return "N/A";
    }

    // ========== Pack Memo (CP_PCCODE >= 4550) ==========

    private void ShowPackMemo(AdsConnection conn, string batchId)
    {
        try
        {
            string sql = $"SELECT pack_mem FROM \"D_LINE.DBF\" WHERE B_id = '{batchId}'";
            using var cmd = new AdsCommand(sql, conn);
            using var rdr = cmd.ExecuteReader();
            if (rdr.Read())
            {
                string memoText = rdr["pack_mem"]?.ToString()?.Trim() ?? "";
                if (!string.IsNullOrEmpty(memoText))
                {
                    using var dlg = new Form
                    {
                        Text = "Pack Memo",
                        Size = new Size(500, 350),
                        StartPosition = FormStartPosition.CenterParent,
                        FormBorderStyle = FormBorderStyle.FixedDialog,
                        MaximizeBox = false,
                        MinimizeBox = false,
                    };
                    var txt = new TextBox
                    {
                        Multiline = true,
                        ReadOnly = true,
                        Dock = DockStyle.Fill,
                        Text = memoText,
                        Font = new Font("Consolas", 10f),
                        ScrollBars = ScrollBars.Vertical,
                    };
                    dlg.Controls.Add(txt);
                    dlg.ShowDialog(this);
                }
            }
        }
        catch { /* pack_mem field or memo file may not exist */ }
    }

    // ========== Movement Data Grid ==========

    private void LoadMovementData(AdsConnection conn, string batchId)
    {
        // Load process name cache
        LoadProcNames(conn);

        // Clipper: ilnmvbs index = b_id+STR(cp_stage,4) → ORDER BY cp_stage
        string sql = $"SELECT b_prior, cp_stage, cpproc_id, cp_bqtyw, cp_bqtys, cp_bqtyp, " +
                     $"arr, cp_darr, cp_tarr, sta, cp_dsta, cp_tsta, fin, cp_dfin, cp_tfin " +
                     $"FROM \"m_linemv.DBF\" WHERE b_id = '{batchId}' ORDER BY cp_stage";

        using var cmd = new AdsCommand(sql, conn);
        using var reader = cmd.ExecuteReader();

        dtResult = BuildDataTable();

        while (reader.Read())
        {
            var row = dtResult.NewRow();

            row["Pr"] = reader["b_prior"]?.ToString()?.Trim() ?? "";
            row["Stage"] = GetInt(reader, "cp_stage");

            // Process & name (Clipper: cpproc_id + " " + ProcName())
            string procId = reader["cpproc_id"]?.ToString() ?? "";
            string? procName = "";
            procNameCache?.TryGetValue(procId.Trim(), out procName);
            row["Process"] = $"{procId.Trim()} {procName ?? ""}";

            row["Wfr"] = GetInt(reader, "cp_bqtyw");
            row["Str"] = GetInt(reader, "cp_bqtys");
            row["Pcs"] = GetInt(reader, "cp_bqtyp");

            // Date/Time columns: "Yes/No" + date + time[:5]
            bool arr = GetBool(reader, "arr");
            bool sta = GetBool(reader, "sta");
            bool fin = GetBool(reader, "fin");

            row["Arr"] = FormatDateTimeCol(arr, reader, "cp_darr", "cp_tarr");
            row["Sta"] = FormatDateTimeCol(sta, reader, "cp_dsta", "cp_tsta");
            row["Fin"] = FormatDateTimeCol(fin, reader, "cp_dfin", "cp_tfin");

            // Hidden flags for color logic
            row["_Arr"] = arr;
            row["_Fin"] = fin;

            dtResult.Rows.Add(row);
        }

        grid.DataSource = dtResult;
        ConfigureGridColumns();
    }

    private void LoadProcNames(AdsConnection conn)
    {
        if (procNameCache != null) return;
        procNameCache = new Dictionary<string, string>();
        try
        {
            string sql = "SELECT proc_id, proc_nme FROM \"C_PROC.DBF\"";
            using var cmd = new AdsCommand(sql, conn);
            using var rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                string id = rdr["proc_id"]?.ToString()?.Trim() ?? "";
                string name = rdr["proc_nme"]?.ToString()?.Trim() ?? "";
                if (!string.IsNullOrEmpty(id))
                    procNameCache.TryAdd(id, name);
            }
        }
        catch { /* C_PROC.DBF may not exist */ }
    }

    private static DataTable BuildDataTable()
    {
        var dt = new DataTable();
        dt.Columns.Add("Pr", typeof(string));
        dt.Columns.Add("Stage", typeof(int));
        dt.Columns.Add("Process", typeof(string));
        dt.Columns.Add("Wfr", typeof(int));
        dt.Columns.Add("Str", typeof(int));
        dt.Columns.Add("Pcs", typeof(int));
        dt.Columns.Add("Arr", typeof(string));
        dt.Columns.Add("Sta", typeof(string));
        dt.Columns.Add("Fin", typeof(string));
        // Hidden columns for color logic
        dt.Columns.Add("_Arr", typeof(bool));
        dt.Columns.Add("_Fin", typeof(bool));
        return dt;
    }

    private void ConfigureGridColumns()
    {
        if (grid.Columns.Count == 0) return;

        var cols = new (string name, string header, int width)[]
        {
            ("Pr",      "Pr",              35),
            ("Stage",   "Stage",           60),
            ("Process", "Process & name", 220),
            ("Wfr",     "Wfr",             50),
            ("Str",     "Str",             70),
            ("Pcs",     "Pcs",             90),
            ("Arr",     "Arr Date    Time",160),
            ("Sta",     "Sta Date    Time",160),
            ("Fin",     "Fin Date    Time",160),
        };

        // Hide internal columns
        if (grid.Columns.Contains("_Arr")) grid.Columns["_Arr"]!.Visible = false;
        if (grid.Columns.Contains("_Fin")) grid.Columns["_Fin"]!.Visible = false;

        int idx = 0;
        foreach (var (name, header, width) in cols)
        {
            if (grid.Columns.Contains(name))
            {
                grid.Columns[name]!.DisplayIndex = idx;
                grid.Columns[name]!.Width = width;
                grid.Columns[name]!.HeaderText = header;
                idx++;
            }
        }

        // Right-align numeric columns
        foreach (var col in new[] { "Stage", "Wfr", "Str", "Pcs" })
        {
            if (grid.Columns.Contains(col))
                grid.Columns[col]!.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
        }

        // Number formatting with thousands separators
        if (grid.Columns.Contains("Str"))
            grid.Columns["Str"]!.DefaultCellStyle.Format = "N0";
        if (grid.Columns.Contains("Pcs"))
            grid.Columns["Pcs"]!.DefaultCellStyle.Format = "N0";
    }

    // ========== Color Logic ==========

    // Clipper: if(m_linemv->arr AND !m_linemv->fin, {2,1}, {1,2})
    // Green background for in-progress stages (arrived but not finished)
    private void OnCellFormatting(object? sender, DataGridViewCellFormattingEventArgs e)
    {
        if (dtResult == null || e.RowIndex < 0 || e.RowIndex >= dtResult.Rows.Count) return;

        var row = dtResult.Rows[e.RowIndex];

        bool arr = false, fin = false;
        try { arr = Convert.ToBoolean(row["_Arr"]); } catch { }
        try { fin = Convert.ToBoolean(row["_Fin"]); } catch { }

        if (arr && !fin)
        {
            e.CellStyle!.ForeColor = Color.Black;
            e.CellStyle!.BackColor = Color.FromArgb(144, 238, 144); // LightGreen — matching B/Bg
        }
    }

    // ========== Keyboard ==========

    // Two-level ESC: grid→input, input→close (matching Clipper BNPATH behavior)
    private void OnFormKeyDown(object? sender, KeyEventArgs e)
    {
        if (e.KeyCode == Keys.Escape)
        {
            if (dtResult != null && dtResult.Rows.Count > 0 && !txtBatchNumber.Focused)
            {
                // ESC from grid → clear and return to batch input
                grid.DataSource = null;
                dtResult = null;
                infoPanel.Visible = false;
                countLabel.Text = "Steps: 0";
                statusLabel.Text = "Enter batch number and press Enter or click 'Query'.";
                txtBatchNumber.Focus();
                txtBatchNumber.SelectAll();
            }
            else
            {
                Close();
            }
            e.Handled = true;
        }
    }

    // ========== Helpers ==========

    private static string FormatDateTimeCol(bool flag, IDataReader reader, string dateField, string timeField)
    {
        string yesNo = flag ? "Yes" : "No ";
        string dateStr = "";
        string timeStr = "";
        try
        {
            var dateVal = reader[dateField];
            if (dateVal is DateTime dt && dt != DateTime.MinValue)
                dateStr = dt.ToString("dd/MM/yyyy");
        }
        catch { }
        try
        {
            string raw = reader[timeField]?.ToString() ?? "";
            timeStr = raw.Length >= 5 ? raw.Substring(0, 5) : raw;
        }
        catch { }
        return $"{yesNo} {dateStr} {timeStr}";
    }

    private static int GetInt(IDataReader reader, string field)
    {
        try
        {
            var val = reader[field];
            if (val == null || val == DBNull.Value) return 0;
            return Convert.ToInt32(val);
        }
        catch { return 0; }
    }

    private static bool GetBool(IDataReader reader, string field)
    {
        try
        {
            var val = reader[field];
            if (val is bool b) return b;
            if (val is string s) return s.Trim().ToUpper() is "T" or "Y" or "TRUE" or "YES";
            return Convert.ToBoolean(val);
        }
        catch { return false; }
    }
}
