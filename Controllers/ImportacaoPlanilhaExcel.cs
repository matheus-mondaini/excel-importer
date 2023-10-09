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

                    ExcelWorksheet worksheet = package.Workbook.Worksheets[worksheetIndex];

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

                        DateTime emissao;
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
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[worksheetIndex];

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

        private static string ReadConnectionStringFromFile(string filePath)
        {
            try
            {
                string connectionString = File.ReadAllText(filePath);
                return connectionString;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erro ao ler a Connection String do arquivo: {ex.Message}");
                return string.Empty;
            }
        }

        public static void InsertDataIntoDatabase<T>(List<T> data, int worksheetIndex)

        {
            string connectionString = ReadConnectionStringFromFile("connectionString.txt"); //Removido antes do push por privacidade

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
                            insertQuery = "SET IDENTITY_INSERT Cliente ON; ";

                            insertQuery += @"
                            INSERT INTO Cliente (ID, Nome, Cidade, UF, CEP, CPF)
                            VALUES (@ID, @Nome, @Cidade, @UF, @CEP, @CPF);";

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
    }
}
