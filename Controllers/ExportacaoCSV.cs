using CsvHelper;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Formats.Asn1;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using CsvHelper.Configuration;
using System.Globalization;
using System.Text.RegularExpressions;

namespace DesafioImportaExcel.Controllers
{
    public class ExportacaoCSV
    {
        public static void Exportar(string filePath, DateTime dataInicio, DateTime dataFim)
        {
            string connectionString = GerenciadorConexaoBancoDados.ReadConnectionStringFromFile("connectionString.txt");

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                CultureInfo culture = Utilitarios.GetSessionCulture(connection);

                var csvConfig = new CsvConfiguration(culture);
                csvConfig.Delimiter = "|";

                using (SqlCommand command = new SqlCommand(
                    "SELECT " +
                        "'CLIENTE' AS Tipo, cli.Nome AS Nome, cli.CPF AS CPF, cli.Cidade AS Cidade, 'DEBITO' AS Tipo, deb.Fatura AS Fatura, deb.Emissao AS Emissao, deb.Vencimento AS Vencimento, " +
                        "deb.Valor AS Valor, deb.ValorPago AS ValorPago, deb.Pagamento AS Pagamento " +
                    "FROM " +
                        "Debitos deb " +
                        "LEFT JOIN CLIENTE cli ON cli.ID = deb.Cliente" +
                    "WHERE Emissao BETWEEN @StartDate AND @EndDate", connection))
                {
                    command.Parameters.AddWithValue("@StartDate", dataInicio);
                    command.Parameters.AddWithValue("@EndDate", dataFim);

                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        using (var writer = new StreamWriter(filePath))
                        using (var csv = new CsvWriter(writer, csvConfig))
                        {
                            csv.WriteRecords(reader);
                        }
                    }
                }
            }
        }
    }
}
