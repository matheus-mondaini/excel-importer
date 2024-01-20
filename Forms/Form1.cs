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
using System.IO.Packaging;
using System.Data.Common;

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
                        MessageBox.Show("O arquivo Excel n�o cont�m planilhas.");
                    }
                    else
                    {
                        int? planilhaSelecionadaIndex = null;
                        var escolhaPlanilhaForm = new EscolhaPlanilhaForm(planilhas);
                        if (escolhaPlanilhaForm.ShowDialog() == DialogResult.OK)
                        {
                            planilhaSelecionadaIndex = planilhas.IndexOf(escolhaPlanilhaForm.PlanilhaSelecionada ?? "");
                        }

                        if (planilhaSelecionadaIndex == null)
                        {
                            MessageBox.Show("Nenhum WorkSheet foi selecionado.");
                        }
                        else
                        {
                            List<object>? dados = ImportacaoPlanilhaExcel.LerDados(excelFilePath, (int)planilhaSelecionadaIndex);
                            dataGridView1.DataSource = dados;

                            worksheetIndex = planilhaSelecionadaIndex;
                            planilhaLida = true;
                            btnInserirNoBanco.Enabled = true;
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ocorreu um erro ao ler a planilha: " + ex.Message);
                }
            }
        }

        private void btnInserirNoBanco_Click(object sender, EventArgs e)
        {
            if (planilhaLida)
            {
                if (worksheetIndex != null)
                {
                    try
                    {
                        List<object> dados = new List<object>();
                        foreach (DataGridViewRow dgvRow in dataGridView1.Rows)
                        {
                            if (!dgvRow.IsNewRow)
                            {
                                if (worksheetIndex == 0)
                                {
                                    dados.Add((Cliente)dgvRow.DataBoundItem);
                                }
                                else if (worksheetIndex == 1)
                                {
                                    dados.Add((Debitos)dgvRow.DataBoundItem);
                                }
                            }
                        }

                        ImportacaoPlanilhaExcel.InserirNoBanco(dados, (int)worksheetIndex);
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
            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "Arquivos CSV (*.csv)|*.csv";
                saveFileDialog.Title = "Escolha o local e o nome do arquivo de exporta��o";
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    DateTimePicker startDatePicker = new DateTimePicker();
                    DateTimePicker endDatePicker = new DateTimePicker();
                    startDatePicker.Format = DateTimePickerFormat.Short;
                    endDatePicker.Format = DateTimePickerFormat.Short;
                    startDatePicker.Location = new Point(10, 10);
                    endDatePicker.Location = new Point(10, 40);
                    Label startDateLabel = new Label();
                    startDateLabel.Text = "Data de In�cio:";
                    startDateLabel.Location = new Point(10, 10);
                    Label endDateLabel = new Label();
                    endDateLabel.Text = "Data de T�rmino:";
                    endDateLabel.Location = new Point(10, 40);
                    Form dateRangeForm = new Form();
                    dateRangeForm.Text = "Escolha o Per�odo de Exporta��o";
                    dateRangeForm.ClientSize = new Size(250, 130);
                    Button exportButton = new Button();
                    exportButton.Text = "Exportar";
                    exportButton.Location = new Point(10, 70);
                    exportButton.Click += (s, ev) =>
                    {
                        ExportacaoCSV.Exportar(saveFileDialog.FileName, startDatePicker.Value, endDatePicker.Value);
                        dateRangeForm.Close();
                    };
                    dateRangeForm.Controls.Add(startDatePicker);
                    dateRangeForm.Controls.Add(endDatePicker);
                    dateRangeForm.Controls.Add(startDateLabel);
                    dateRangeForm.Controls.Add(endDateLabel);
                    dateRangeForm.Controls.Add(exportButton);
                    dateRangeForm.ShowDialog();
                }
            }

        }
    }
}