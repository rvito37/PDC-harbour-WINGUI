using Advantage.Data.Provider;

namespace PdcGui;

/// <summary>
/// Shared Slack / LTime_NoRoute calculation used by StockInProc and StockInWork screens.
/// Matching Clipper DELIVSCH.PRG lines 1017-1079 and AVXFUNCS.PRG Slack() function.
/// Slack = B_dprom - (Today + LTime_NoRoute)
/// </summary>
public static class SlackCalculator
{
    /// <summary>
    /// Load all c_leadt records into a dictionary for fast lookup.
    /// Key: ptype_id + proc_id + pline_id (or ptype_id + proc_id for U/_/K types)
    /// Value: LEADT_DAYS
    /// Source: Clipper BMS\DELIVSCH.PRG, c_leadt index "itcppr"
    /// </summary>
    public static Dictionary<string, double>? LoadLeadTimes(AdsConnection conn)
    {
        try
        {
            string sql = "SELECT ptype_id, proc_id, pline_id, LEADT_DAYS FROM \"c_leadt.DBF\"";
            using var cmd = new AdsCommand(sql, conn);
            using var reader = cmd.ExecuteReader();

            var dict = new Dictionary<string, double>();
            while (reader.Read())
            {
                string ptype = reader["ptype_id"]?.ToString() ?? "";
                string proc  = reader["proc_id"]?.ToString() ?? "";
                string pline = reader["pline_id"]?.ToString() ?? "";
                double days  = 0;
                try { days = Convert.ToDouble(reader["LEADT_DAYS"]); } catch { }

                // Store both full key and short key (for U/_/K types)
                string fullKey = ptype + proc + pline;
                string shortKey = ptype + proc;
                dict.TryAdd(fullKey, days);
                dict.TryAdd(shortKey, days);
            }
            return dict;
        }
        catch
        {
            return null; // c_leadt.DBF not available
        }
    }

    /// <summary>
    /// Load m_linemv steps for a set of batch IDs, grouped by batch.
    /// Returns dict: batchId -> list of (ptype_id, cpproc_id, pline_id, fin) ordered by cp_stage.
    /// Source: Clipper BMS\DELIVSCH.PRG lines 1046-1065
    /// </summary>
    public static Dictionary<string, List<(string ptype, string proc, string pline, bool fin)>>?
        LoadBatchSteps(AdsConnection conn, IEnumerable<string> batchIds)
    {
        try
        {
            // Build IN clause for all batches
            var ids = string.Join(",", batchIds.Select(b => $"'{b}'"));
            if (string.IsNullOrEmpty(ids)) return null;

            string sql = $"SELECT B_ID, PTYPE_ID, CPPROC_ID, PLINE_ID, FIN, cp_stage " +
                         $"FROM \"m_linemv.DBF\" WHERE B_ID IN ({ids}) ORDER BY B_ID, cp_stage";
            using var cmd = new AdsCommand(sql, conn);
            using var reader = cmd.ExecuteReader();

            var dict = new Dictionary<string, List<(string, string, string, bool)>>();
            while (reader.Read())
            {
                string bid   = reader["B_ID"]?.ToString()?.Trim() ?? "";
                string ptype = reader["PTYPE_ID"]?.ToString() ?? "";
                string proc  = reader["CPPROC_ID"]?.ToString() ?? "";
                string pline = reader["PLINE_ID"]?.ToString() ?? "";
                bool fin     = false;
                try
                {
                    var finVal = reader["FIN"];
                    fin = finVal is bool b ? b : Convert.ToBoolean(finVal);
                }
                catch { }

                if (!dict.ContainsKey(bid))
                    dict[bid] = new List<(string, string, string, bool)>();
                dict[bid].Add((ptype, proc, pline, fin));
            }
            return dict;
        }
        catch
        {
            return null; // m_linemv.DBF not available
        }
    }

    /// <summary>
    /// Calculate LTime_NoRoute for a single batch.
    /// Matching Clipper DELIVSCH.PRG lines 1017-1079:
    ///   1. Find batch steps ordered by stage
    ///   2. Skip finished steps (FIN = true)
    ///   3. For remaining steps, sum lead times from c_leadt
    ///   4. Round up if fractional (NotCommas function)
    /// </summary>
    public static int CalcLTimeNoRoute(
        List<(string ptype, string proc, string pline, bool fin)>? steps,
        Dictionary<string, double>? leadTimes)
    {
        if (steps == null || leadTimes == null) return 0;

        double totalDays = 0;
        bool pastFinished = false;

        foreach (var (ptype, proc, pline, fin) in steps)
        {
            // Skip past all finished steps (Clipper: WHILE FIN; dbskip; END)
            if (!pastFinished)
            {
                if (fin) continue;
                pastFinished = true;
            }

            // Look up lead time
            // Clipper: IF m_linemv->PTYPE_ID $ "U_K" -> seek by PTYPE_ID + CPPROC_ID only
            // The $ operator means "is contained in", so PTYPE_ID is a single char checked
            // against "U_K" -- matches 'U', '_', or 'K'
            double days = 0;
            string ptypeChar = ptype.Length > 0 ? ptype.Substring(0, 1) : "";
            if (ptypeChar == "U" || ptypeChar == "_" || ptypeChar == "K")
            {
                // Short key: ptype + proc (Clipper: c_leadt->(DBSEEK(PTYPE_ID + CPPROC_ID)))
                string key = ptype + proc;
                leadTimes.TryGetValue(key, out days);
            }
            else
            {
                // Full key: ptype + proc + pline
                string key = ptype + proc + pline;
                leadTimes.TryGetValue(key, out days);
            }

            totalDays += days;
        }

        // NotCommas: round up if fractional (Clipper AVXFUNCS.PRG line 3954)
        if (totalDays % 1 > 0)
            return (int)totalDays + 1;
        return (int)totalDays;
    }

    /// <summary>
    /// Calculate Slack for a list of batch rows using LTime_NoRoute.
    /// Returns true if lead time data was available.
    /// </summary>
    public static bool CalculateSlack(
        AdsConnection conn,
        List<(System.Data.DataRow row, string batchId, DateTime? promDate)> batchRows)
    {
        var leadTimes = LoadLeadTimes(conn);
        var allBatchIds = batchRows.Where(b => b.promDate.HasValue).Select(b => b.batchId);
        var batchSteps = LoadBatchSteps(conn, allBatchIds);
        bool hasData = (leadTimes != null && batchSteps != null);

        foreach (var (row, batchId, promDate) in batchRows)
        {
            if (!promDate.HasValue) continue;

            if (hasData)
            {
                // Clipper: Slack = b_dprom - (date() + LTime_NoRoute)
                batchSteps!.TryGetValue(batchId, out var steps);
                int ltime = CalcLTimeNoRoute(steps, leadTimes);
                row["Slack"] = (promDate.Value - DateTime.Today).Days - ltime;
            }
            else
            {
                // Fallback: no lead time tables -> Slack = b_dprom - date()
                row["Slack"] = (promDate.Value - DateTime.Today).Days;
            }
        }

        return hasData;
    }
}
