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
        private bool planilhaLida = false;
        int? worksheetIndex = null;

        public Form1()
        {
            InitializeComponent();
        }

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
                            List<dynamic>? dados = ImportacaoPlanilhaExcel.ReadDataFromExcel(excelFilePath, (int)planilhaSelecionadaIndex);
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

                var escolhaPlanilhaForm = new EscolhaPlanilhaForm(planilhas);
                if (escolhaPlanilhaForm.ShowDialog() == DialogResult.OK)
                {
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
                        List<object> data = new List<object>();
                        foreach (DataGridViewRow dgvRow in dataGridView1.Rows)
                        {
                            if (!dgvRow.IsNewRow)
                            {
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

                        ImportacaoPlanilhaExcel.InsertDataIntoDatabase(data, (int)worksheetIndex);
                        MessageBox.Show("Dados inseridos com sucesso no banco de dados!");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Ocorreu um erro ao inserir os dados: " + ex.Message);
                    }
                }
            }

        }

        private void btnExportar_Click(object sender, EventArgs e)
        {

        }
    }
}