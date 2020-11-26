namespace ExportarDiagApp
{
    partial class Form1
    {
        /// <summary>
        /// Variável de designer necessária.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Limpar os recursos que estão sendo usados.
        /// </summary>
        /// <param name="disposing">verdade se for necessário descartar os recursos gerenciados; caso contrário, falso.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region código gerado pelo Windows Form Designer

        /// <summary>
        /// Método necessário para suporte do Designer - não modifique
        /// o conteúdo deste método com o editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            this.txtDirEntrada = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtDirSaida = new System.Windows.Forms.TextBox();
            this.btnDirEntrada = new System.Windows.Forms.Button();
            this.btnDirSaida = new System.Windows.Forms.Button();
            this.btnGerarArquivo = new System.Windows.Forms.Button();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.SuspendLayout();
            // 
            // txtDirEntrada
            // 
            this.txtDirEntrada.Location = new System.Drawing.Point(11, 25);
            this.txtDirEntrada.Name = "txtDirEntrada";
            this.txtDirEntrada.Size = new System.Drawing.Size(353, 20);
            this.txtDirEntrada.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(101, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Diretório de Entrada";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 55);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(93, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "Diretório de Saída";
            // 
            // txtDirSaida
            // 
            this.txtDirSaida.Location = new System.Drawing.Point(11, 71);
            this.txtDirSaida.Name = "txtDirSaida";
            this.txtDirSaida.Size = new System.Drawing.Size(353, 20);
            this.txtDirSaida.TabIndex = 2;
            // 
            // btnDirEntrada
            // 
            this.btnDirEntrada.Location = new System.Drawing.Point(370, 23);
            this.btnDirEntrada.Name = "btnDirEntrada";
            this.btnDirEntrada.Size = new System.Drawing.Size(26, 23);
            this.btnDirEntrada.TabIndex = 4;
            this.btnDirEntrada.Text = "...";
            this.btnDirEntrada.UseVisualStyleBackColor = true;
            this.btnDirEntrada.Click += new System.EventHandler(this.btnDirEntrada_Click);
            // 
            // btnDirSaida
            // 
            this.btnDirSaida.Location = new System.Drawing.Point(370, 69);
            this.btnDirSaida.Name = "btnDirSaida";
            this.btnDirSaida.Size = new System.Drawing.Size(26, 23);
            this.btnDirSaida.TabIndex = 5;
            this.btnDirSaida.Text = "...";
            this.btnDirSaida.UseVisualStyleBackColor = true;
            this.btnDirSaida.Click += new System.EventHandler(this.btnDirSaida_Click);
            // 
            // btnGerarArquivo
            // 
            this.btnGerarArquivo.Location = new System.Drawing.Point(11, 98);
            this.btnGerarArquivo.Name = "btnGerarArquivo";
            this.btnGerarArquivo.Size = new System.Drawing.Size(385, 23);
            this.btnGerarArquivo.TabIndex = 6;
            this.btnGerarArquivo.Text = "Gerar Arquivo de Exportação";
            this.btnGerarArquivo.UseVisualStyleBackColor = true;
            this.btnGerarArquivo.Click += new System.EventHandler(this.btnGerarArquivo_Click);
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(11, 128);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(385, 23);
            this.progressBar1.TabIndex = 7;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(409, 160);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.btnGerarArquivo);
            this.Controls.Add(this.btnDirSaida);
            this.Controls.Add(this.btnDirEntrada);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtDirSaida);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtDirEntrada);
            this.Name = "Form1";
            this.Text = "Exportar Diagnóstico - SIEM";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtDirEntrada;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtDirSaida;
        private System.Windows.Forms.Button btnDirEntrada;
        private System.Windows.Forms.Button btnDirSaida;
        private System.Windows.Forms.Button btnGerarArquivo;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;

    }
}

