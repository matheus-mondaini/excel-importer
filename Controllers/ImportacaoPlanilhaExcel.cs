using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using OfficeOpenXml;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Globalization;
using DesafioImportaExcel.Models;
using Microsoft.VisualBasic.ApplicationServices;

namespace DesafioImportaExcel.Controllers
{
    public class ImportacaoPlanilhaExcel
    {
        public static List<string> GetWorksheetNames(FileInfo excelFilePath)
        {
            var worksheetNames = new List<string>();

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            try
            {
                using (ExcelPackage package = new ExcelPackage(excelFilePath))
                {
                    foreach (var worksheet in package.Workbook.Worksheets)
                    {
                        worksheetNames.Add(worksheet.Name);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ocorreu um erro: " + ex.Message);
            }

            return worksheetNames;
        }

        public static List<dynamic>? ReadDataFromExcel(FileInfo excelFilePath, int worksheetIndex)
        {
            var listaGenerica = new List<dynamic>();
            if (worksheetIndex == 0)
            {
                List<Cliente> cliente = ReadClientesFromExcel(excelFilePath, worksheetIndex);
                listaGenerica.AddRange(cliente);
            }
            else if (worksheetIndex == 1)
            {
                List<Debitos> debitos = ReadDebitosFromExcel(excelFilePath, worksheetIndex);
                listaGenerica.AddRange(debitos);
            }
            else
            {
                MessageBox.Show("Esta não é um escolha disponível");
                listaGenerica = null;
            }
            return listaGenerica;
        }

        public static List<Debitos> ReadDebitosFromExcel(FileInfo excelFilePath, int worksheetIndex)
        {
            var debitosList = new List<Debitos>();

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            try
            {

                using (ExcelPackage package = new ExcelPackage(excelFilePath))
                {

                    ExcelWorksheet worksheet = package.Workbook.Worksheets[worksheetIndex]; // Sabendo que os dados estão na primeira planilha

                    int rowCount = worksheet.Dimension.Rows;

                    for (int row = 2; row <= rowCount; row++)
                    {
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

                        //string[] dateFormats = { "MM/dd/yyyy", "M/dd/yyyy", "MM/d/yyyy", "M/d/yyyy" };

                        DateTime emissao;
                        //if (DateTime.TryParseExact(worksheet.Cells[row, 3].Text, dateFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out emissao))
                        if (DateTime.TryParse(worksheet.Cells[row, 3].Text, out emissao))
                        {
                            debito.Emissao = emissao;
                        }
                        else
                        {
                            MessageBox.Show($"Erro na linha {row}: Data inválida para Emissao - {worksheet.Cells[row, 3].Text}");
                            continue;
                        }

                        DateTime vencimento;
                        //if (DateTime.TryParseExact(worksheet.Cells[row, 4].Text, dateFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out vencimento))
                        if (DateTime.TryParse(worksheet.Cells[row, 4].Text, out vencimento))
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

                        DateTime pagamento;
                        //if (DateTime.TryParseExact(worksheet.Cells[row, 8].Text, dateFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out pagamento))
                        if (DateTime.TryParse(worksheet.Cells[row, 8].Text, out pagamento))
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
        public static List<Cliente> ReadClientesFromExcel(FileInfo excelFilePath, int worksheetIndex)
        {
            var clientesList = new List<Cliente>();

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            try
            {
                using (ExcelPackage package = new ExcelPackage(excelFilePath))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[worksheetIndex]; // Worksheet "Cliente"

                    int rowCount = worksheet.Dimension.Rows;

                    for (int row = 2; row <= rowCount; row++)
                    {
                        Cliente cliente = new Cliente();
                        cliente.ID = int.Parse(worksheet.Cells[row, 1].Text);
                        cliente.Nome = worksheet.Cells[row, 2].Text;
                        cliente.Cidade = worksheet.Cells[row, 3].Text;
                        cliente.UF = worksheet.Cells[row, 4].Text;
                        cliente.CEP = worksheet.Cells[row, 5].Text;
                        cliente.CPF = worksheet.Cells[row, 6].Text;

                        clientesList.Add(cliente);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ocorreu um erro na leitura da planilha Cliente: " + ex.Message);
            }
            return clientesList;
        }


        public static void InsertDataIntoDatabase<T>(List<T> data, int worksheetIndex)

        {
            string connectionString = "Server = Gemini\\SQL2019; Database = Desafio_Planilha; User Id = sa; Password = cdssql;";

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                foreach (var item in data)
                {
                    string insertQuery = "";
                    if (worksheetIndex == 0 && item is Cliente cliente)
                    {
                        TabelaCliente tabelaCliente = new TabelaCliente(connectionString);
                        tabelaCliente.CriarTabelaSeNaoExistir();
                        {
                            // Habilitar a inserção explícita de valores na coluna de identidade
                            insertQuery = "SET IDENTITY_INSERT Cliente ON; ";

                            insertQuery += @"
                            INSERT INTO Cliente (ID, Nome, Cidade, UF, CEP, CPF)
                            VALUES (@ID, @Nome, @Cidade, @UF, @CEP, @CPF);";

                            // Desabilitar a inserção explícita de valores na coluna de identidade
                            insertQuery += "SET IDENTITY_INSERT Cliente OFF;";
                        }

                        using (SqlCommand command = new SqlCommand(insertQuery, connection))
                        {
                            command.Parameters.AddWithValue("@ID", cliente.ID);
                            command.Parameters.AddWithValue("@Nome", cliente.Nome);
                            command.Parameters.AddWithValue("@Cidade", cliente.Cidade);
                            command.Parameters.AddWithValue("@UF", cliente.UF);
                            command.Parameters.AddWithValue("@CEP", cliente.CEP);
                            command.Parameters.AddWithValue("@CPF", cliente.CPF);

                            command.ExecuteNonQuery();
                        }
                    }
                    else if (worksheetIndex == 1 && item is Debitos debitos)
                    {
                        TabelaDebitos tabelaDebitos = new TabelaDebitos(connectionString);
                        tabelaDebitos.CriarTabelaSeNaoExistir();
                        insertQuery = @"
                        INSERT INTO Debitos (Fatura, Cliente, Emissao, Vencimento, Valor, Juros, Descontos, Pagamento, ValorPago)
                        VALUES (@Fatura, @Cliente, @Emissao, @Vencimento, @Valor, @Juros, @Descontos, @Pagamento, @ValorPago)";

                        debitos.Emissao = Utilitarios.ConverteParaDataValida(debitos.Emissao.ToString());
                        debitos.Vencimento = Utilitarios.ConverteParaDataValida(debitos.Vencimento.ToString());
                        debitos.Pagamento = Utilitarios.ConverteParaDataValida(debitos.Pagamento.ToString());

                        using (SqlCommand command = new SqlCommand(insertQuery, connection))
                        {
                            command.Parameters.AddWithValue("@Fatura", debitos.Fatura);
                            command.Parameters.AddWithValue("@Cliente", debitos.Cliente);
                            command.Parameters.AddWithValue("@Emissao", debitos.Emissao);
                            command.Parameters.AddWithValue("@Vencimento", debitos.Vencimento);
                            command.Parameters.AddWithValue("@Valor", debitos.Valor);
                            command.Parameters.AddWithValue("@Juros", debitos.Juros);
                            command.Parameters.AddWithValue("@Descontos", debitos.Descontos);
                            command.Parameters.AddWithValue("@Pagamento", debitos.Pagamento);
                            command.Parameters.AddWithValue("@ValorPago", debitos.ValorPago);

                            command.ExecuteNonQuery();
                        }
                    }
                    else
                    {
                        MessageBox.Show("Tipo de objeto não suportado ou índice de planilha inválido.");
                    }
                }
            }
        }

        //public static void InsertDataIntoDatabase(List<DataRow> rows, int worksheetIndex)
        //{
        //    string connectionString = "";

        //    using (SqlConnection connection = new SqlConnection(connectionString))
        //    {
        //        connection.Open();

        //        foreach (DataRow row in rows)
        //        {
        //            string insertQuery = "";
        //            if (worksheetIndex == 0)
        //            {
        //                TabelaCliente cliente = new TabelaCliente(connectionString);
        //                cliente.CriarTabelaSeNaoExistir();
        //                {
        //                    insertQuery = @"
        //                    INSERT INTO Clientes (ID, Nome, Cidade, UF, CEP, CPF)
        //                    VALUES (@ID, @Nome, @Cidade, @UF, @CEP, @CPF)";
        //                }

        //                using (SqlCommand command = new SqlCommand(insertQuery, connection))
        //                {
        //                    command.Parameters.AddWithValue("@ID", row["ID"]);
        //                    command.Parameters.AddWithValue("@Nome", row["Nome"]);
        //                    command.Parameters.AddWithValue("@Cidade", row["Cidade"]);
        //                    command.Parameters.AddWithValue("@UF", row["UF"]);
        //                    command.Parameters.AddWithValue("@CEP", row["CEP"]);
        //                    command.Parameters.AddWithValue("@CPF", row["CPF"]);
        //                }
        //            }
        //            else if (worksheetIndex == 1)
        //            {
        //                TabelaDebitos debitos = new TabelaDebitos(connectionString);
        //                debitos.CriarTabelaSeNaoExistir();
        //                insertQuery = @"
        //                INSERT INTO Debitos (Fatura, Cliente, Emissao, Vencimento, Valor, Juros, Descontos, Pagamento, ValorPago)
        //                VALUES (@Fatura, @Cliente, @Emissao, @Vencimento, @Valor, @Juros, @Descontos, @Pagamento, @ValorPago)";

        //                using (SqlCommand command = new SqlCommand(insertQuery, connection))
        //                {
        //                    command.Parameters.AddWithValue("@Fatura", row["Fatura"]);
        //                    command.Parameters.AddWithValue("@Cliente", row["Cliente"]);
        //                    command.Parameters.AddWithValue("@Emissao", row["Emissao"]);
        //                    command.Parameters.AddWithValue("@Vencimento", row["Vencimento"]);
        //                    command.Parameters.AddWithValue("@Valor", row["Valor"]);
        //                    command.Parameters.AddWithValue("@Juros", row["Juros"]);
        //                    command.Parameters.AddWithValue("@Descontos", row["Descontos"]);
        //                    command.Parameters.AddWithValue("@Pagamento", row["Pagamento"]);
        //                    command.Parameters.AddWithValue("@ValorPago", row["ValorPago"]);

        //                    command.ExecuteNonQuery();
        //                }
        //            }
        //            else
        //            {
        //                MessageBox.Show("Tipo de objeto não suportado ou índice de planilha inválido.");
        //            }
        //        }
        //    }
        //}

        //public static void InsertDataIntoDatabase(List<object> dataList, int worksheetIndex)
        //{
        //    string connectionString = "";

        //    using (SqlConnection connection = new SqlConnection(connectionString))
        //    {
        //        connection.Open();

        //        foreach (var data in dataList)
        //        {
        //            string insertQuery = "";

        //            if (worksheetIndex == 0 && data is Cliente)
        //            {
        //                insertQuery = @"
        //        INSERT INTO Clientes (ID, Nome, Cidade, UF, CEP, CPF)
        //        VALUES (@ID, @Nome, @Cidade, @UF, @CEP, @CPF)";
        //            }
        //            else if (worksheetIndex == 1 && data is Debitos)
        //            {
        //                insertQuery = @"
        //        INSERT INTO Debitos (Fatura, Cliente, Emissao, Vencimento, Valor, Juros, Descontos, Pagamento, ValorPago)
        //        VALUES (@Fatura, @Cliente, @Emissao, @Vencimento, @Valor, @Juros, @Descontos, @Pagamento, @ValorPago)";
        //            }

        //            using (SqlCommand command = new SqlCommand(insertQuery, connection))
        //            {
        //                if (data is Cliente cliente)
        //                {
        //                    command.Parameters.AddWithValue("@ID", cliente.ID);
        //                    command.Parameters.AddWithValue("@Nome", cliente.Nome);
        //                    command.Parameters.AddWithValue("@Cidade", cliente.Cidade);
        //                    command.Parameters.AddWithValue("@UF", cliente.UF);
        //                    command.Parameters.AddWithValue("@CEP", cliente.CEP);
        //                    command.Parameters.AddWithValue("@CPF", cliente.CPF);
        //                }
        //                else if (data is Debitos debito)
        //                {
        //                    command.Parameters.AddWithValue("@Fatura", debito.Fatura);
        //                    command.Parameters.AddWithValue("@Cliente", debito.Cliente);
        //                    command.Parameters.AddWithValue("@Emissao", debito.Emissao);
        //                    command.Parameters.AddWithValue("@Vencimento", debito.Vencimento);
        //                    command.Parameters.AddWithValue("@Valor", debito.Valor);
        //                    command.Parameters.AddWithValue("@Juros", debito.Juros);
        //                    command.Parameters.AddWithValue("@Descontos", debito.Descontos);
        //                    command.Parameters.AddWithValue("@Pagamento", debito.Pagamento);
        //                    command.Parameters.AddWithValue("@ValorPago", debito.ValorPago);
        //                }

        //                command.ExecuteNonQuery();
        //            }
        //        }
        //    }
        //}


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
                        string fatura = row["Fatura"].ToString();
                        int cliente = int.Parse(row["Cliente"].ToString());

                        // Verificar se já existe uma linha com a mesma Fatura e Cliente
                        if (IsDuplicateFaturaCliente(connection, fatura, cliente))
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

        private bool IsDuplicateFaturaCliente(SqlConnection connection, string fatura, int cliente)
        {
            string query = "SELECT COUNT(*) FROM SuaTabela WHERE Fatura = @Fatura AND Cliente = @Cliente";
            using (SqlCommand command = new SqlCommand(query, connection))
            {
                command.Parameters.AddWithValue("@Fatura", fatura);
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
