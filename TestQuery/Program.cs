using Advantage.Data.Provider;

namespace TestQuery;

class Program
{
    static void Main(string[] args)
    {
        string connStr = @"Data Source=C:\Users\AVXUser\BMS\DATA;ServerType=LOCAL;TableType=CDX;LockMode=COMPATIBLE;CharType=OEM;TrimTrailingSpaces=TRUE;";

        try
        {
            using var conn = new AdsConnection(connStr);
            conn.Open();
            Console.WriteLine("Connected OK.");

            // Load lead times
            Console.WriteLine("Loading c_leadt...");
            var leadTimes = new Dictionary<string, double>();
            using (var cmdLt = new AdsCommand("SELECT ptype_id, proc_id, pline_id, LEADT_DAYS FROM \"c_leadt.DBF\"", conn))
            using (var rLt = cmdLt.ExecuteReader())
            {
                while (rLt.Read())
                {
                    string ptype = rLt["ptype_id"]?.ToString() ?? "";
                    string proc = rLt["proc_id"]?.ToString() ?? "";
                    string pline = rLt["pline_id"]?.ToString() ?? "";
                    double days = 0;
                    try { days = Convert.ToDouble(rLt["LEADT_DAYS"]); } catch { }
                    leadTimes.TryAdd(ptype + proc + pline, days);
                    leadTimes.TryAdd(ptype + proc, days);
                }
            }
            Console.WriteLine($"Loaded {leadTimes.Count} lead time entries.");

            // Get batches for process 200.0
            string procId = "200.0";
            Console.WriteLine($"\nQuerying D_LINE for process {procId}...");
            var batchIds = new List<string>();
            var batchPromDates = new Dictionary<string, DateTime>();

            using (var cmdD = new AdsCommand(
                $"SELECT B_id, B_dprom FROM \"D_LINE.DBF\" WHERE CpProc_id = '{procId}' " +
                "AND UPPER(B_stat) NOT IN ('C','D','M','T') ORDER BY islack", conn))
            using (var rD = cmdD.ExecuteReader())
            {
                while (rD.Read())
                {
                    string bid = rD["B_id"]?.ToString()?.Trim() ?? "";
                    batchIds.Add(bid);
                    try
                    {
                        var dprom = rD["B_dprom"];
                        if (dprom is DateTime dt && dt != DateTime.MinValue)
                            batchPromDates[bid] = dt;
                    }
                    catch { }
                }
            }
            Console.WriteLine($"Found {batchIds.Count} batches.");

            // Load m_linemv steps for these batches
            Console.WriteLine("Loading m_linemv steps...");
            var sw = System.Diagnostics.Stopwatch.StartNew();
            var ids = string.Join(",", batchIds.Select(b => $"'{b}'"));
            var batchSteps = new Dictionary<string, List<(string ptype, string proc, string pline, bool fin)>>();

            using (var cmdM = new AdsCommand(
                $"SELECT B_ID, PTYPE_ID, CPPROC_ID, PLINE_ID, FIN, cp_stage " +
                $"FROM \"m_linemv.DBF\" WHERE B_ID IN ({ids}) ORDER BY B_ID, cp_stage", conn))
            using (var rM = cmdM.ExecuteReader())
            {
                while (rM.Read())
                {
                    string bid = rM["B_ID"]?.ToString()?.Trim() ?? "";
                    string ptype = rM["PTYPE_ID"]?.ToString() ?? "";
                    string proc = rM["CPPROC_ID"]?.ToString() ?? "";
                    string pline = rM["PLINE_ID"]?.ToString() ?? "";
                    bool fin = false;
                    try { var fv = rM["FIN"]; fin = fv is bool b ? b : Convert.ToBoolean(fv); } catch { }

                    if (!batchSteps.ContainsKey(bid))
                        batchSteps[bid] = new List<(string, string, string, bool)>();
                    batchSteps[bid].Add((ptype, proc, pline, fin));
                }
            }
            sw.Stop();
            Console.WriteLine($"Loaded steps for {batchSteps.Count} batches in {sw.ElapsedMilliseconds}ms");

            // Calculate and display
            Console.WriteLine($"\n{"#",-4} {"Batch",-8} {"B_dprom",-12} {"Simple",-8} {"LTime",-6} {"Real",-8}");
            Console.WriteLine(new string('-', 55));

            int row = 0;
            foreach (string bid in batchIds)
            {
                row++;
                DateTime today = DateTime.Today;
                int simpleSlack = 0;
                if (batchPromDates.TryGetValue(bid, out var dprom))
                    simpleSlack = (dprom - today).Days;

                // Calculate LTime_NoRoute
                batchSteps.TryGetValue(bid, out var steps);
                int ltime = CalcLTime(steps, leadTimes);
                int realSlack = simpleSlack - ltime;

                Console.WriteLine($"{row,-4} {bid,-8} {dprom:yyyy-MM-dd}  {simpleSlack,-8} {ltime,-6} {realSlack,-8}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"ERROR: {ex.Message}");
            Console.WriteLine(ex.StackTrace);
        }
    }

    static int CalcLTime(
        List<(string ptype, string proc, string pline, bool fin)>? steps,
        Dictionary<string, double> leadTimes)
    {
        if (steps == null) return 0;

        double totalDays = 0;
        bool pastFinished = false;

        foreach (var (ptype, proc, pline, fin) in steps)
        {
            if (!pastFinished)
            {
                if (fin) continue;
                pastFinished = true;
            }

            double days = 0;
            string ptypeChar = ptype.Length > 0 ? ptype.Substring(0, 1) : "";
            if (ptypeChar == "U" || ptypeChar == "_" || ptypeChar == "K")
            {
                leadTimes.TryGetValue(ptype + proc, out days);
            }
            else
            {
                leadTimes.TryGetValue(ptype + proc + pline, out days);
            }
            totalDays += days;
        }

        if (totalDays % 1 > 0)
            return (int)totalDays + 1;
        return (int)totalDays;
    }
}
