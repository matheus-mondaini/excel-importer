namespace DesafioImportaExcel
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            btnLerPlanilha = new Button();
            dataGridView1 = new DataGridView();
            btnInserirNoBanco = new Button();
            btnExportar = new Button();
            ((System.ComponentModel.ISupportInitialize)dataGridView1).BeginInit();
            SuspendLayout();
            // 
            // btnLerPlanilha
            // 
            btnLerPlanilha.BackColor = SystemColors.ButtonHighlight;
            btnLerPlanilha.Font = new Font("Segoe UI", 10F, FontStyle.Bold, GraphicsUnit.Point);
            btnLerPlanilha.ForeColor = SystemColors.ControlText;
            btnLerPlanilha.Location = new Point(14, 31);
            btnLerPlanilha.Margin = new Padding(3, 4, 3, 4);
            btnLerPlanilha.Name = "btnLerPlanilha";
            btnLerPlanilha.Size = new Size(125, 59);
            btnLerPlanilha.TabIndex = 0;
            btnLerPlanilha.Text = "Ler Planilha";
            btnLerPlanilha.UseVisualStyleBackColor = false;
            btnLerPlanilha.Click += btnLerPlanilha_Click;
            // 
            // dataGridView1
            // 
            dataGridView1.BackgroundColor = SystemColors.ButtonHighlight;
            dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridView1.Location = new Point(14, 120);
            dataGridView1.Margin = new Padding(3, 4, 3, 4);
            dataGridView1.Name = "dataGridView1";
            dataGridView1.RowHeadersWidth = 51;
            dataGridView1.RowTemplate.Height = 25;
            dataGridView1.Size = new Size(735, 375);
            dataGridView1.TabIndex = 1;
            // 
            // btnInserirNoBanco
            // 
            btnInserirNoBanco.BackColor = SystemColors.ButtonHighlight;
            btnInserirNoBanco.Font = new Font("Segoe UI", 10F, FontStyle.Bold, GraphicsUnit.Point);
            btnInserirNoBanco.Location = new Point(273, 31);
            btnInserirNoBanco.Margin = new Padding(3, 4, 3, 4);
            btnInserirNoBanco.Name = "btnInserirNoBanco";
            btnInserirNoBanco.Size = new Size(167, 59);
            btnInserirNoBanco.TabIndex = 2;
            btnInserirNoBanco.Text = "Inserir No Banco";
            btnInserirNoBanco.UseVisualStyleBackColor = false;
            btnInserirNoBanco.Click += btnInserirNoBanco_Click;
            // 
            // btnExportar
            // 
            btnExportar.BackColor = SystemColors.ButtonHighlight;
            btnExportar.Font = new Font("Segoe UI", 10F, FontStyle.Bold, GraphicsUnit.Point);
            btnExportar.Location = new Point(555, 31);
            btnExportar.Name = "btnExportar";
            btnExportar.Size = new Size(152, 59);
            btnExportar.TabIndex = 3;
            btnExportar.Text = "Exportar Dados";
            btnExportar.UseVisualStyleBackColor = false;
            btnExportar.Click += btnExportar_Click;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = SystemColors.ControlText;
            ClientSize = new Size(762, 511);
            Controls.Add(btnExportar);
            Controls.Add(btnInserirNoBanco);
            Controls.Add(dataGridView1);
            Controls.Add(btnLerPlanilha);
            ForeColor = SystemColors.ControlText;
            Margin = new Padding(3, 4, 3, 4);
            Name = "Form1";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "Desafio 1";
            ((System.ComponentModel.ISupportInitialize)dataGridView1).EndInit();
            ResumeLayout(false);
        }

        #endregion

        private Button btnLerPlanilha;
        private DataGridView dataGridView1;
        private Button btnInserirNoBanco;
        private Button btnExportar;
    }
}