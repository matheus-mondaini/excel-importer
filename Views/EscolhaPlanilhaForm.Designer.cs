namespace DesafioImportaExcel
{
    partial class EscolhaPlanilhaForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            comboBoxEscolherWorksheet = new ComboBox();
            btnConfirmarSelecao = new Button();
            SuspendLayout();
            // 
            // comboBoxEscolherWorksheet
            // 
            comboBoxEscolherWorksheet.BackColor = SystemColors.ButtonHighlight;
            comboBoxEscolherWorksheet.Font = new Font("Segoe UI", 10F, FontStyle.Bold, GraphicsUnit.Point);
            comboBoxEscolherWorksheet.FormattingEnabled = true;
            comboBoxEscolherWorksheet.Location = new Point(21, 12);
            comboBoxEscolherWorksheet.Name = "comboBoxEscolherWorksheet";
            comboBoxEscolherWorksheet.Size = new Size(273, 25);
            comboBoxEscolherWorksheet.TabIndex = 0;
            comboBoxEscolherWorksheet.Text = "Escolha sua Worksheet";
            // 
            // btnConfirmarSelecao
            // 
            btnConfirmarSelecao.AutoSize = true;
            btnConfirmarSelecao.BackColor = SystemColors.ButtonHighlight;
            btnConfirmarSelecao.Font = new Font("Segoe UI", 10F, FontStyle.Bold, GraphicsUnit.Point);
            btnConfirmarSelecao.Location = new Point(83, 342);
            btnConfirmarSelecao.Name = "btnConfirmarSelecao";
            btnConfirmarSelecao.Size = new Size(145, 29);
            btnConfirmarSelecao.TabIndex = 1;
            btnConfirmarSelecao.Text = "Confirmar Seleção";
            btnConfirmarSelecao.UseVisualStyleBackColor = false;
            btnConfirmarSelecao.Click += btnConfirmarSelecao_Click;
            // 
            // EscolhaPlanilhaForm
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = SystemColors.ControlText;
            ClientSize = new Size(314, 383);
            Controls.Add(btnConfirmarSelecao);
            Controls.Add(comboBoxEscolherWorksheet);
            Name = "EscolhaPlanilhaForm";
            StartPosition = FormStartPosition.CenterParent;
            Text = "Escolher Tabela";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private ComboBox comboBoxEscolherWorksheet;
        private Button btnConfirmarSelecao;
    }
}