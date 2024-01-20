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
using System.Data;

namespace DesafioImportaExcel.Controllers
{
    public class ExportacaoCSV
    {
        public static void Exportar(string filePath, DateTime dataInicio, DateTime dataFim)
        {
            try
            {
                using (var command = GerenciadorConexaoBancoDados.Conectar().CreateCommand())
                {
                    command.CommandText =
                        "SELECT " +
                            "'CLIENTE' AS Tipo, cli.Nome AS Nome, cli.CPF AS CPF, cli.Cidade AS Cidade, 'DEBITO' AS Tipo, deb.Fatura AS Fatura, deb.Emissao AS Emissao, deb.Vencimento AS Vencimento, " +
                            "deb.Valor AS Valor, deb.ValorPago AS ValorPago, deb.Pagamento AS Pagamento " +
                        "FROM " +
                            "Debitos deb " +
                            "LEFT JOIN CLIENTE cli ON cli.ID = deb.Cliente " +
                        "WHERE deb.Emissao BETWEEN @StartDate AND @EndDate";

                    command.Parameters.Add("@StartDate", SqlDbType.DateTime).Value = dataInicio;
                    command.Parameters.Add("@EndDate", SqlDbType.DateTime).Value = dataFim;

                    using (SqlDataAdapter dataAdapter = new SqlDataAdapter(command))
                    {
                        DataTable dataTable = new DataTable();
                        dataAdapter.Fill(dataTable);

                        StringBuilder csvData = new StringBuilder();

                        foreach (DataColumn column in dataTable.Columns)
                        {
                            csvData.Append(column.ColumnName + "|");
                        }
                        csvData.AppendLine();

                        foreach (DataRow row in dataTable.Rows)
                        {
                            foreach (DataColumn column in dataTable.Columns)
                            {
                                csvData.Append(row[column].ToString() + "|");
                            }
                            csvData.AppendLine();
                        }

                        File.WriteAllText(filePath, csvData.ToString());

                        MessageBox.Show("Data exported to CSV successfully.");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred: " + ex.Message);
            }
        }
    }
}