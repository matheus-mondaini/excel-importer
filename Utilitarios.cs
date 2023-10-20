using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data.SqlTypes;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DesafioImportaExcel
{
    public static class Utilitarios
    {
        //public static DateTime ConverteParaDataValida(string dateText)
        //{
        //    DateTime parsedDate;
        //    if (DateTime.TryParseExact(dateText, new string[] { "MM/dd/yyyy", "M/dd/yyyy", "MM/d/yyyy", "M/d/yyyy", "dd/MM/yyyy", "d/MM/yyyy", "dd/M/yyyy", "d/M/yyyy" }, CultureInfo.InvariantCulture, DateTimeStyles.None, out parsedDate))
        //    {
        //        if (parsedDate >= SqlDateTime.MinValue.Value && parsedDate <= SqlDateTime.MaxValue.Value)
        //        {
        //            return parsedDate;
        //        }
        //        else
        //        {
        //            return SqlDateTime.MinValue.Value;
        //        }
        //    }
        //    return SqlDateTime.MinValue.Value;
        //}

        public static CultureInfo GetSessionCulture(this SqlConnection connection)
        {
            using (SqlCommand cmd = new SqlCommand("SELECT @@language", connection)) // "SELECT [language] FROM sys.dm_exec_sessions WHERE session_id = @@SPID"
            {
                var language = cmd.ExecuteScalar().ToString();
                if (!string.IsNullOrEmpty(language))
                {
                    return new CultureInfo(language);
                }
                else
                {
                    return CultureInfo.InvariantCulture;
                }
            }
        }
    }
}
