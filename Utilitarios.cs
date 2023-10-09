using System;
using System.Collections.Generic;
using System.Data.SqlTypes;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DesafioImportaExcel
{
    public static class Utilitarios
    {
        public static DateTime ConverteParaDataValida(string dateText)
        {
            DateTime parsedDate;
            // Tente analisar a data com vários formatos, incluindo o formato brasileiro (dd/MM/yyyy)
            if (DateTime.TryParseExact(dateText, new string[] { "MM/dd/yyyy", "M/dd/yyyy", "MM/d/yyyy", "M/d/yyyy", "dd/MM/yyyy", "d/MM/yyyy", "dd/M/yyyy", "d/M/yyyy" }, CultureInfo.InvariantCulture, DateTimeStyles.None, out parsedDate))
            {
                // Verifique se a data está dentro do intervalo suportado pelo SQL Server
                if (parsedDate >= SqlDateTime.MinValue.Value && parsedDate <= SqlDateTime.MaxValue.Value)
                {
                    // Se a data estiver dentro do intervalo válido, retorne-a no formato MM/dd/yyyy
                    return parsedDate;
                }
                else
                {
                    // A data está fora do intervalo, retorne a data mínima suportada pelo SQL Server
                    return SqlDateTime.MinValue.Value;
                }
            }
            // Se a conversão falhar, retorne a data mínima suportada pelo SQL Server
            return SqlDateTime.MinValue.Value;
        }
    }
}
