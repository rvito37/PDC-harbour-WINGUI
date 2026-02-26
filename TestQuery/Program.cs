using Advantage.Data.Provider;

namespace TestQuery;

class Program
{
    static void Main(string[] args)
    {
        string connStr = @"Data Source=C:\Users\AVXUser\BMS\DATA;ServerType=LOCAL;TableType=CDX;LockMode=COMPATIBLE;CharType=OEM;TrimTrailingSpaces=TRUE;";
        string dbfConn = @"Data Source=C:\Users\AVXUser\BMS\DATA;ServerType=LOCAL;TableType=DBF;LockMode=COMPATIBLE;CharType=OEM;TrimTrailingSpaces=TRUE;";
        string batchId = args.Length > 0 ? args[0] : "977535";

        try
        {
            using var conn = new AdsConnection(connStr);
            conn.Open();
            Console.WriteLine("Connected OK (CDX mode).");

            // 1. Check D_LINE
            Console.WriteLine($"\n=== D_LINE for batch {batchId} ===");
            string bPurp = "", bType = "", esnxx = "", pline = "";
            using (var cmd = new AdsCommand(
                $"SELECT B_id, b_purp, b_type, esnxx_id, pline_id, CP_PCCODE FROM \"D_LINE.DBF\" WHERE B_id = '{batchId}'", conn))
            using (var rdr = cmd.ExecuteReader())
            {
                if (rdr.Read())
                {
                    bPurp = rdr["b_purp"]?.ToString() ?? "";
                    bType = rdr["b_type"]?.ToString() ?? "";
                    esnxx = rdr["esnxx_id"]?.ToString() ?? "";
                    pline = rdr["pline_id"]?.ToString() ?? "";
                    Console.WriteLine($"  b_purp=[{bPurp}] b_type=[{bType}] esnxx=[{esnxx}] pline=[{pline}]");
                }
            }

            // 2. Test C_BTYPE with DBF mode (no CDX)
            Console.WriteLine($"\n=== C_BTYPE via DBF mode (no CDX) ===");
            using (var conn2 = new AdsConnection(dbfConn))
            {
                conn2.Open();
                using var cmd2 = new AdsCommand(
                    $"SELECT btype_nme FROM \"c_btype.DBF\" WHERE b_type = '{bType.Trim()}' AND esnxx_id = '{esnxx}' AND pline_id = '{pline}'", conn2);
                using var rdr2 = cmd2.ExecuteReader();
                if (rdr2.Read())
                    Console.WriteLine($"  btype_nme = [{rdr2["btype_nme"]}]");
                else
                    Console.WriteLine("  NOT FOUND");
            }

            // 3. Purpose lookup (should work with CDX)
            Console.WriteLine($"\n=== C_BPURP lookup b_purp=[{bPurp.Trim()}] ===");
            using (var cmd3 = new AdsCommand(
                $"SELECT bpurp_nme FROM \"C_BPURP.DBF\" WHERE b_purp = '{bPurp.Trim()}'", conn))
            using (var rdr3 = cmd3.ExecuteReader())
            {
                if (rdr3.Read())
                    Console.WriteLine($"  bpurp_nme = [{rdr3["bpurp_nme"]}]");
                else
                    Console.WriteLine("  NOT FOUND");
            }

            Console.WriteLine("\nDone.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"ERROR: {ex.Message}");
            Console.WriteLine(ex.StackTrace);
        }
    }
}
