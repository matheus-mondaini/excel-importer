using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using OfficeOpenXml;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Globalization;

namespace DesafioImportaExcel
{
    public class ImportacaoPlanilhaExcel
    {
        public static List<Debitos> ReadDataFromExcel(FileInfo excelFilePath)
        {
            var debitosList = new List<Debitos>();

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            try
            {

                using (ExcelPackage package = new ExcelPackage(excelFilePath))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[1]; // Sabendo que os dados estão na primeira planilha

                    int rowCount = worksheet.Dimension.Rows;

                    for (int row = 2; row <= rowCount; row++)
                    {
                        //Debitos debito = new Debitos
                        //{
                        //    Fatura = worksheet.Cells[row, 1].ToString(),
                        //    Cliente = int.Parse(worksheet.Cells[row, 2].ToString()),
                        //    Emissao = DateTime.Parse(worksheet.Cells[row, 3].ToString()),
                        //    Vencimento = DateTime.Parse(worksheet.Cells[row, 4].ToString()),
                        //    Valor = decimal.Parse(worksheet.Cells[row, 5].ToString()),
                        //    Juros = decimal.Parse(worksheet.Cells[row, 6].ToString()),
                        //    Descontos = decimal.Parse(worksheet.Cells[row, 7].ToString()),
                        //    Pagamento = DateTime.Parse(worksheet.Cells[row, 8].ToString()),
                        //    ValorPago = decimal.Parse(worksheet.Cells[row, 9].ToString())
                        //};
                        Debitos debito = new Debitos();
                        debito.Fatura = worksheet.Cells[row, 1].Text;

                        int cliente;
                        if (int.TryParse(worksheet.Cells[row, 2].Text, out cliente))
                        {
                            debito.Cliente = cliente;
                        }
                        else
                        {
                            MessageBox.Show($"Erro na linha {row}: Valor inválido para Cliente - {worksheet.Cells[row, 2].Text}");
                            continue;
                        }

                        // Converter data de Emissao (formato MM/dd/yyyy ou M/dd/yyyy)
                        DateTime emissao;
                        string[] dateFormats = { "MM/dd/yyyy", "M/dd/yyyy", "MM/d/yyyy", "M/d/yyyy" };

                        if (DateTime.TryParseExact(worksheet.Cells[row, 3].Text, dateFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out emissao))
                        {
                            debito.Emissao = emissao;
                        }
                        else
                        {
                            MessageBox.Show($"Erro na linha {row}: Data inválida para Emissao - {worksheet.Cells[row, 3].Text}");
                            continue;
                        }

                        // Converter data de Vencimento (formato MM/dd/yyyy ou M/dd/yyyy)
                        DateTime vencimento;
                        if (DateTime.TryParseExact(worksheet.Cells[row, 4].Text, dateFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out vencimento))
                        {
                            debito.Vencimento = vencimento;
                        }
                        else
                        {
                            MessageBox.Show($"Erro na linha {row}: Data inválida para Vencimento - {worksheet.Cells[row, 4].Text}");
                            continue;
                        }

                        decimal valor;
                        if (decimal.TryParse(worksheet.Cells[row, 5].Text, out valor))
                        {
                            debito.Valor = valor;
                        }
                        else
                        {
                            MessageBox.Show($"Erro na linha {row}: Valor inválido para Valor - {worksheet.Cells[row, 5].Text}");
                            continue;
                        }

                        decimal juros;
                        if (!string.IsNullOrEmpty(worksheet.Cells[row, 6].Text) && decimal.TryParse(worksheet.Cells[row, 6].Text, out juros))
                        {
                            debito.Juros = juros;
                        }

                        decimal descontos;
                        if (!string.IsNullOrEmpty(worksheet.Cells[row, 7].Text) && decimal.TryParse(worksheet.Cells[row, 7].Text, out descontos))
                        {
                            debito.Descontos = descontos;
                        }

                        // Converter data de Pagamento (formato MM/dd/yyyy ou M/dd/yyyy)
                        DateTime pagamento;
                        if (DateTime.TryParseExact(worksheet.Cells[row, 8].Text, dateFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out pagamento))
                        {
                            debito.Pagamento = pagamento;
                        }

                        decimal valorPago;
                        if (!string.IsNullOrEmpty(worksheet.Cells[row, 9].Text) && decimal.TryParse(worksheet.Cells[row, 9].Text, out valorPago))
                        {
                            debito.ValorPago = valorPago;
                        }

                        debitosList.Add(debito);
                    }
                }



                //Caso fosse utilizado datateble ao invés de uma instância da classe:

                    //DataTable dataTable = new DataTable();
                    
                    //try
                    //{

                    //    using (var package = new ExcelPackage(new System.IO.FileInfo(excelFilePath)))
                    //    {
                    //        var worksheet = package.Workbook.Worksheets[0]; // Supondo que os dados estão na primeira planilha

                    //        int rowCount = worksheet.Dimension.Rows;

                    //        for (int row = 2; row <= rowCount; row++)
                    //        {
                    //            DataRow dataRow = dataTable.NewRow();
                    //            dataRow["Fatura"] = worksheet.Cells[row, 1].Text;
                    //            dataRow["Cliente"] = int.Parse(worksheet.Cells[row, 2].Text);
                    //            dataRow["Emissao"] = DateTime.Parse(worksheet.Cells[row, 3].Text);
                    //            dataRow["Vencimento"] = DateTime.Parse(worksheet.Cells[row, 4].Text);
                    //            dataRow["Valor"] = decimal.Parse(worksheet.Cells[row, 5].Text);
                    //            dataRow["Juros"] = decimal.Parse(worksheet.Cells[row, 6].Text);
                    //            dataRow["Descontos"] = decimal.Parse(worksheet.Cells[row, 7].Text);
                    //            dataRow["Pagamento"] = DateTime.Parse(worksheet.Cells[row, 8].Text);
                    //            dataRow["ValorPago"] = decimal.Parse(worksheet.Cells[row, 9].Text);
                    //            dataTable.Rows.Add(dataRow);
                    //        }
                    //    }
                    //}
                    
                    //return DataTable;


                //Caso não soubéssemos a planilha previamanete e precisássemos fazer de maneira mais genéria (também pode-se adaptar para uma instancia da classe ao invés de datatable):

                    //DataTable dataTable = new DataTable();

                    //try
                    //{
                        //using (var package = new ExcelPackage(new System.IO.FileInfo(excelFilePath)))
                        //{
                        //    var worksheet = package.Workbook.Worksheets[0];
                        //    foreach (var cell in worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column])
                        //    {
                        //        dataTable.Columns.Add(cell.Text);
                        //    }

                        //    for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                        //    {
                        //        DataRow dataRow = dataTable.NewRow();
                        //        for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                        //        {
                        //            dataRow[col - 1] = worksheet.Cells[row, col].Text;
                        //        }
                        //        dataTable.Rows.Add(dataRow);
                        //    }
                        //}
                    //}

                    //return DataTable;


            }
            catch (Exception ex)
            {
                MessageBox.Show("Ocorreu um erro: " + ex.Message);
            }
            return debitosList;
        }

        public static void InsertDataIntoDatabase(List<DataRow> rows)
        {
            string connectionString = "";

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                foreach (DataRow row in rows)
                {
                    string insertQuery = @"
                    INSERT INTO Debitos (NumeroFatura, Cliente, Emissao, Vencimento, Valor, Juros, Descontos, Pagamento, ValorPago)
                    VALUES (@NumeroFatura, @Cliente, @Emissao, @Vencimento, @Valor, @Juros, @Descontos, @Pagamento, @ValorPago)";

                    using (SqlCommand command = new SqlCommand(insertQuery, connection))
                    {
                        command.Parameters.AddWithValue("@Fatura", row["Fatura"]);
                        command.Parameters.AddWithValue("@Cliente", row["Cliente"]);
                        command.Parameters.AddWithValue("@Emissao", row["Emissao"]);
                        command.Parameters.AddWithValue("@Vencimento", row["Vencimento"]);
                        command.Parameters.AddWithValue("@Valor", row["Valor"]);
                        command.Parameters.AddWithValue("@Juros", row["Juros"]);
                        command.Parameters.AddWithValue("@Descontos", row["Descontos"]);
                        command.Parameters.AddWithValue("@Pagamento", row["Pagamento"]);
                        command.Parameters.AddWithValue("@ValorPago", row["ValorPago"]);

                        command.ExecuteNonQuery();
                    }
                }
            }
        }
 
    /*
    private string _connectionString;

    public ImportacaoPlanilhaExcel(string connectionString)
    {
        _connectionString = connectionString;
    }

    public DataTable ReadDataFromExcel(string excelFilePath)
    {

    public bool InsertDataIntoDatabase(DataTable dataTable)
    {
        try
        {
            using (SqlConnection connection = new SqlConnection(_connectionString))
            {
                connection.Open();

                foreach (DataRow row in dataTable.Rows)
                {
                    string numeroFatura = row["Fatura"].ToString();
                    int cliente = int.Parse(row["Cliente"].ToString());

                    // Verificar se já existe uma linha com a mesma Fatura e Cliente
                    if (IsDuplicateFaturaCliente(connection, numeroFatura, cliente))
                    {
                        DialogResult result = MessageBox.Show("Já existe uma linha com a mesma Fatura e Cliente. Deseja substituir?", "Aviso", MessageBoxButtons.YesNoCancel);

                        if (result == DialogResult.Cancel)
                        {
                            return false; // Cancelar a ação
                        }
                        else if (result == DialogResult.No)
                        {
                            continue; // Ignorar e continuar com a próxima linha
                        }
                        // Se result for DialogResult.Yes, continua para a inserção
                    }

                    // Inserir os dados no banco de dados usando Dapper
                    InsertData(connection, row);
                }
            }

            return true;
        }
        catch (Exception)
        {
            return false;
        }
    }

    private bool IsDuplicateFaturaCliente(SqlConnection connection, string numeroFatura, int cliente)
    {
        string query = "SELECT COUNT(*) FROM SuaTabela WHERE NumeroFatura = @NumeroFatura AND Cliente = @Cliente";
        using (SqlCommand command = new SqlCommand(query, connection))
        {
            command.Parameters.AddWithValue("@NumeroFatura", numeroFatura);
            command.Parameters.AddWithValue("@Cliente", cliente);
            int count = (int)command.ExecuteScalar();

            return count > 0;
        }
    }

    private void InsertData(SqlConnection connection, DataRow row)
    {
        string insertQuery = "INSERT INTO SuaTabela (Fatura, Cliente, Emissao, Vencimento, Valor, Juros, Descontos, Pagamento, ValorPago) " +
                                "VALUES (@Fatura, @Cliente, @Emissao, @Vencimento, @Valor, @Juros, @Descontos, @Pagamento, @ValorPago)";

        using (SqlCommand command = new SqlCommand(insertQuery, connection))
        {
            command.Parameters.AddWithValue("@Fatura", row["Fatura"]);
            command.Parameters.AddWithValue("@Cliente", int.Parse(row["Cliente"].ToString()));
            command.Parameters.AddWithValue("@Emissao", DateTime.Parse(row["Emissao"].ToString()));
            command.Parameters.AddWithValue("@Vencimento", DateTime.Parse(row["Vencimento"].ToString()));
            command.Parameters.AddWithValue("@Valor", decimal.Parse(row["Valor"].ToString()));
            command.Parameters.AddWithValue("@Juros", decimal.Parse(row["Juros"].ToString()));
            command.Parameters.AddWithValue("@Descontos", decimal.Parse(row["Descontos"].ToString()));
            command.Parameters.AddWithValue("@Pagamento", DateTime.Parse(row["Pagamento"].ToString()));
            command.Parameters.AddWithValue("@ValorPago", decimal.Parse(row["ValorPago"].ToString()));

            command.ExecuteNonQuery();
        }
    }
    */
    }
}
