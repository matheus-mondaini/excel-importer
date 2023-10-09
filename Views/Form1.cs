using OfficeOpenXml;
using System.Data;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Diagnostics.Metrics;
using System.Drawing;
using System;
using Microsoft.VisualBasic.ApplicationServices;
using DesafioImportaExcel.Controllers;
using DesafioImportaExcel.Models;

namespace DesafioImportaExcel
{
    public partial class Form1 : Form
    {
        private bool planilhaLida = false; // Indica se a planilha foi lida com sucesso
        int? worksheetIndex = null;

        public Form1()
        {
            InitializeComponent();
        }

        //GerenciadorConexaoBancoDados ConectaSQL = new GerenciadorConexaoBancoDados("x\\x", "x", "x", "x");

        private void btnLerPlanilha_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Arquivos Excel|*.xlsx|Todos os Arquivos|*.*";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                FileInfo excelFilePath = new FileInfo(openFileDialog.FileName);

                try
                {
                    List<string> planilhas = ImportacaoPlanilhaExcel.GetWorksheetNames(excelFilePath);

                    if (planilhas.Count == 0)
                    {
                        MessageBox.Show("O arquivo Excel não contém planilhas.");
                    }
                    else
                    {
                        int? planilhaSelecionadaIndex = EscolherPlanilha(excelFilePath, planilhas);

                        if (planilhaSelecionadaIndex != null)
                        {
                            string planilhaSelecionadaNome = planilhas[planilhaSelecionadaIndex.Value];
                            List<dynamic> dados = ImportacaoPlanilhaExcel.ReadDataFromExcel(excelFilePath, (int)planilhaSelecionadaIndex);
                            dataGridView1.DataSource = dados;
                            
                            worksheetIndex = planilhaSelecionadaIndex;
                            planilhaLida = true;
                            btnInserirNoBanco.Enabled = true;
                        }
                        else
                        {
                            MessageBox.Show("Nenhum WorkSheet foi selecionado.");
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ocorreu um erro ao ler a planilha: " + ex.Message);
                }
            }
        }

        private int? EscolherPlanilha(FileInfo excelFilePath, List<string> planilhas)
        {
            using (var package = new ExcelPackage(excelFilePath))
            {
                var worksheets = package.Workbook.Worksheets;

                // Abre um forms para selecionar a planilha desejada
                //var escolhaPlanilhaForm = new EscolhaPlanilhaForm(worksheets.Select(ws => ws.Name).ToList());
                var escolhaPlanilhaForm = new EscolhaPlanilhaForm(planilhas);
                if (escolhaPlanilhaForm.ShowDialog() == DialogResult.OK)
                {
                    //return escolhaPlanilhaForm.PlanilhaSelecionadaIndex;
                    return planilhas.IndexOf(escolhaPlanilhaForm.PlanilhaSelecionada);
                }
            }
            return null;
        }

        private void btnInserirNoBanco_Click(object sender, EventArgs e)
        {
            if (planilhaLida)
            {
                if (worksheetIndex != null)
                {
                    try
                    {
                        // Obtenha os dados da DataGridView em uma lista de objetos
                        List<object> data = new List<object>();
                        foreach (DataGridViewRow dgvRow in dataGridView1.Rows)
                        {
                            if (!dgvRow.IsNewRow)
                            {
                                // Adicione o objeto correto com base no tipo de planilha (Cliente ou Debitos)
                                if (worksheetIndex == 0)
                                {
                                    data.Add((Cliente)dgvRow.DataBoundItem);
                                }
                                else if (worksheetIndex == 1)
                                {
                                    data.Add((Debitos)dgvRow.DataBoundItem);
                                }
                            }
                        }

                        // Chame o método InsertDataIntoDatabase passando a lista de objetos e o worksheetIndex
                        ImportacaoPlanilhaExcel.InsertDataIntoDatabase(data, (int)worksheetIndex);
                        MessageBox.Show("Dados inseridos com sucesso no banco de dados!");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Ocorreu um erro ao inserir os dados: " + ex.Message);
                    }
                }


                /*
                    string connectionString = @"Server = x\x; Database = xr; User Id = x; Password = x;";

                    try
                    {
                        TabelaDebitos debitos = new TabelaDebitos(connectionString);
                        ImportacaoPlanilhaExcel importaDados = new ImportacaoPlanilhaExcel(connectionString);

                        debitos.CriarTabelaSeNaoExistir();
                        bool success = importaDados.InsertDataIntoDatabase(dataTable);

                        if (success)
                        {
                            MessageBox.Show("Dados inseridos com sucesso no banco de dados!");
                        }
                        else
                        {
                            MessageBox.Show("Falha ao inserir dados no banco de dados.");
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Ocorreu um erro: " + ex.Message);
                    }
                }
                */
            }

        }

    }
}