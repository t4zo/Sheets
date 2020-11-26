namespace ExportarDiagApp
{
    partial class ExportarTurma
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
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.btnGerarArquivo = new System.Windows.Forms.Button();
            this.btnDirSaida = new System.Windows.Forms.Button();
            this.btnDirEntrada = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.txtDirSaida = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.txtDirEntrada = new System.Windows.Forms.TextBox();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.SuspendLayout();
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(12, 127);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(385, 23);
            this.progressBar1.TabIndex = 15;
            // 
            // btnGerarArquivo
            // 
            this.btnGerarArquivo.Location = new System.Drawing.Point(12, 97);
            this.btnGerarArquivo.Name = "btnGerarArquivo";
            this.btnGerarArquivo.Size = new System.Drawing.Size(385, 23);
            this.btnGerarArquivo.TabIndex = 14;
            this.btnGerarArquivo.Text = "Gerar Arquivo de Exportação";
            this.btnGerarArquivo.UseVisualStyleBackColor = true;
            this.btnGerarArquivo.Click += new System.EventHandler(this.btnGerarArquivo_Click);
            // 
            // btnDirSaida
            // 
            this.btnDirSaida.Location = new System.Drawing.Point(371, 68);
            this.btnDirSaida.Name = "btnDirSaida";
            this.btnDirSaida.Size = new System.Drawing.Size(26, 23);
            this.btnDirSaida.TabIndex = 13;
            this.btnDirSaida.Text = "...";
            this.btnDirSaida.UseVisualStyleBackColor = true;
            this.btnDirSaida.Click += new System.EventHandler(this.btnDirSaida_Click);
            // 
            // btnDirEntrada
            // 
            this.btnDirEntrada.Location = new System.Drawing.Point(371, 22);
            this.btnDirEntrada.Name = "btnDirEntrada";
            this.btnDirEntrada.Size = new System.Drawing.Size(26, 23);
            this.btnDirEntrada.TabIndex = 12;
            this.btnDirEntrada.Text = "...";
            this.btnDirEntrada.UseVisualStyleBackColor = true;
            this.btnDirEntrada.Click += new System.EventHandler(this.btnDirEntrada_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(13, 54);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(93, 13);
            this.label2.TabIndex = 11;
            this.label2.Text = "Diretório de Saída";
            // 
            // txtDirSaida
            // 
            this.txtDirSaida.Location = new System.Drawing.Point(12, 70);
            this.txtDirSaida.Name = "txtDirSaida";
            this.txtDirSaida.Size = new System.Drawing.Size(353, 20);
            this.txtDirSaida.TabIndex = 10;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 8);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(101, 13);
            this.label1.TabIndex = 9;
            this.label1.Text = "Diretório de Entrada";
            // 
            // txtDirEntrada
            // 
            this.txtDirEntrada.Location = new System.Drawing.Point(12, 24);
            this.txtDirEntrada.Name = "txtDirEntrada";
            this.txtDirEntrada.Size = new System.Drawing.Size(353, 20);
            this.txtDirEntrada.TabIndex = 8;
            // 
            // ExportarTurma
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(407, 162);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.btnGerarArquivo);
            this.Controls.Add(this.btnDirSaida);
            this.Controls.Add(this.btnDirEntrada);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtDirSaida);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtDirEntrada);
            this.Name = "ExportarTurma";
            this.Text = "ExportarTurma";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Button btnGerarArquivo;
        private System.Windows.Forms.Button btnDirSaida;
        private System.Windows.Forms.Button btnDirEntrada;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtDirSaida;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtDirEntrada;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
    }
}