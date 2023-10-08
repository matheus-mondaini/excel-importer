using OfficeOpenXml;
using System.Data;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Diagnostics.Metrics;
using System.Drawing;
using System;
using Microsoft.VisualBasic.ApplicationServices;

namespace DesafioImportaExcel
{
    public partial class Form1 : Form
    {
        private List<Debitos> debitosList; // Para armazenar os dados da planilha
        private bool planilhaLida = false; // Indica se a planilha foi lida com sucesso


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
                    List<Debitos> debitosList = ImportacaoPlanilhaExcel.ReadDataFromExcel(excelFilePath);
                    dataGridView1.DataSource = debitosList;
                        //DataTable dataTable = ImportacaoPlanilhaExcel.ReadDataFromExcel(excelFilePath);
                        //dataGridView1.DataSource = dataTable;
                    planilhaLida = true;
                    btnInserirNoBanco.Enabled = true; // Ativar o botão de inserção após a leitura
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
                try
                {
                    // Obtenha os dados da DataGridView em uma lista de DataRow
                    List<DataRow> rows = new List<DataRow>();
                    foreach (DataGridViewRow dgvRow in dataGridView1.Rows)
                    {
                        if (!dgvRow.IsNewRow)
                        {
                            rows.Add(((DataRowView)dgvRow.DataBoundItem).Row);
                        }
                    }

                    ImportacaoPlanilhaExcel.InsertDataIntoDatabase(rows);
                    MessageBox.Show("Dados inseridos com sucesso no banco de dados!");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ocorreu um erro ao inserir os dados: " + ex.Message);
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