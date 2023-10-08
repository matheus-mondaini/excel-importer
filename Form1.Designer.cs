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
            this.btnLerPlanilha = new System.Windows.Forms.Button();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.btnInserirNoBanco = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // btnLerPlanilha
            // 
            this.btnLerPlanilha.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.btnLerPlanilha.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.btnLerPlanilha.ForeColor = System.Drawing.SystemColors.ControlText;
            this.btnLerPlanilha.Location = new System.Drawing.Point(123, 22);
            this.btnLerPlanilha.Name = "btnLerPlanilha";
            this.btnLerPlanilha.Size = new System.Drawing.Size(109, 44);
            this.btnLerPlanilha.TabIndex = 0;
            this.btnLerPlanilha.Text = "Ler Planilha";
            this.btnLerPlanilha.UseVisualStyleBackColor = false;
            this.btnLerPlanilha.Click += new System.EventHandler(this.btnLerPlanilha_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.BackgroundColor = System.Drawing.SystemColors.ButtonHighlight;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(12, 90);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 25;
            this.dataGridView1.Size = new System.Drawing.Size(643, 281);
            this.dataGridView1.TabIndex = 1;
            // 
            // btnInserirNoBanco
            // 
            this.btnInserirNoBanco.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.btnInserirNoBanco.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.btnInserirNoBanco.Location = new System.Drawing.Point(378, 23);
            this.btnInserirNoBanco.Name = "btnInserirNoBanco";
            this.btnInserirNoBanco.Size = new System.Drawing.Size(146, 44);
            this.btnInserirNoBanco.TabIndex = 2;
            this.btnInserirNoBanco.Text = "Inserir No Banco";
            this.btnInserirNoBanco.UseVisualStyleBackColor = false;
            this.btnInserirNoBanco.Click += new System.EventHandler(this.btnInserirNoBanco_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlText;
            this.ClientSize = new System.Drawing.Size(667, 383);
            this.Controls.Add(this.btnInserirNoBanco);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.btnLerPlanilha);
            this.ForeColor = System.Drawing.SystemColors.ControlText;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Desafio 1";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private Button btnLerPlanilha;
        private DataGridView dataGridView1;
        private Button btnInserirNoBanco;
    }
}