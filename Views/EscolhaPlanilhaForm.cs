using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DesafioImportaExcel
{
    public partial class EscolhaPlanilhaForm : Form
    {
        public string PlanilhaSelecionada { get; private set; }

        private List<string> planilhasDisponiveis;

        public EscolhaPlanilhaForm(List<string> planilhas)
        {
            InitializeComponent();
            planilhasDisponiveis = planilhas;
            // Preencha o ComboBox com as opções de planilha
            comboBoxEscolherWorksheet.DataSource = planilhasDisponiveis;
        }

        private void btnConfirmarSelecao_Click(object sender, EventArgs e)
        {
            // Obtenha a planilha selecionada a partir do ComboBox
            PlanilhaSelecionada = comboBoxEscolherWorksheet.SelectedItem.ToString();
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}
