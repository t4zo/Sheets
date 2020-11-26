using System;
using System.Linq;
using System.IO;
using System.Windows.Forms;
using ExportarDiagApp.Models;
using ExportarDiagApp.Properties;
using Excel = Microsoft.Office.Interop.Excel;


namespace ExportarDiagApp
{
    public partial class ExportarTurma : Form
    {
        public ExportarTurma()
        {
            InitializeComponent();
        }

        //Seleciona o Diretório de Entrada
        private void btnDirEntrada_Click(object sender, EventArgs e)
        {
            folderBrowserDialog1.ShowDialog();
            txtDirEntrada.Text = folderBrowserDialog1.SelectedPath + @"\";
        }

        //Seleciona o Diretório de Saída
        private void btnDirSaida_Click(object sender, EventArgs e)
        {
            folderBrowserDialog1.ShowDialog();
            txtDirSaida.Text = folderBrowserDialog1.SelectedPath + @"\";
        }

        private void btnGerarArquivo_Click(object sender, EventArgs e)
        {
            progressBar1.Value = 0;

            if (string.IsNullOrEmpty(txtDirEntrada.Text) || string.IsNullOrEmpty(txtDirSaida.Text))
            {
                MessageBox.Show(Resources.Diretorios_Nao_Selecionados);
            }
            else
            {
                var arquivoSaida = txtDirSaida.Text + "Saida-PerfilTurma.xlsx";

                var arquivos = Directory.GetFiles(txtDirEntrada.Text, "*.xls*");
                if (!arquivos.Any())
                {
                    MessageBox.Show(Resources.Diretorio_Vazio);
                }
                else
                {
                    var myExcelApp = new Excel.Application
                    {
                        Visible = false,
                        DisplayAlerts = false
                    };

                    arquivos = arquivos.Where(x => !x.Contains("$")).ToArray();

                    progressBar1.Maximum = arquivos.Count();

                    foreach (var arquivo in arquivos)
                    {
                        try
                        {
                            // Aqui é onde deve ser alterado para funcionar para outras planilhas
                            var formulario = new PlanilhasPerfilTurma();

                            var saidaExiste = File.Exists(arquivoSaida);

                            var myBook = myExcelApp.Workbooks.Open(arquivo);
                            var saidaBook = saidaExiste ? myExcelApp.Workbooks.Open(arquivoSaida) : myExcelApp.Workbooks.Add();

                            foreach (var planilha in formulario.Planilhas)
                            {
                                #region Abre a planilha

                                var mySheet = (Excel.Worksheet)myBook.Sheets[planilha.Ordem];
                                mySheet.Unprotect(planilha.Senha);

                                var linhaSaida = 1;
                                Excel.Worksheet saidaSheet;
                                if (saidaExiste)
                                {
                                    saidaSheet = (Excel.Worksheet)saidaBook.Sheets[planilha.OrdemSaida];
                                    linhaSaida = saidaSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                                }
                                else if (planilha.OrdemSaida == 1)
                                {
                                    saidaSheet = (Excel.Worksheet)saidaBook.Sheets[planilha.OrdemSaida];
                                    saidaSheet.Name = planilha.Nome;
                                    AdicionaTitulo(ref saidaSheet, planilha);
                                }
                                else
                                {
                                    saidaSheet =
                                        (Excel.Worksheet)saidaBook.Sheets.Add(After: saidaBook.Sheets[planilha.OrdemSaida - 1]);
                                    saidaSheet.Name = planilha.Nome;
                                    AdicionaTitulo(ref saidaSheet, planilha);
                                }

                                var lastRow = mySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                                #endregion

                                var celulaFixa = mySheet.Cells[planilha.LinhaFixa, planilha.ColunaFixa].Value == null ? string.Empty : mySheet.Cells[planilha.LinhaFixa, planilha.ColunaFixa].Value.ToString().Trim().ToUpper();

                                #region Cadastra as linhas da planilha
                                for (var row = planilha.LinhaInicial; row <= lastRow; row++)
                                {
                                    var colVerif = mySheet.Cells[row, planilha.CelulaVerif];

                                    if (colVerif.Value != null)
                                    {
                                        int verify = 0;
                                        int.TryParse(colVerif.Value.ToString(), out verify);

                                        if (verify > 0)
                                        {
                                            planilha.Linhas.Add(new Linha(row, planilha.Celulas));
                                        }
                                    }
                                }
                                #endregion

                                #region Percorre todos os valores da planilha
                                foreach (var linha in planilha.Linhas)
                                {
                                    linhaSaida++;
                                    foreach (var celula in linha.Celulas)
                                    {
                                        #region Ler os valores
                                        var valorCelula = mySheet.Cells[linha.Numero, celula.Col].Value == null ? string.Empty : mySheet.Cells[linha.Numero, celula.Col].Value.ToString();

                                        switch (celula.Tipo)
                                        {
                                            case TipoCelula.Numero:
                                                int auxInt;
                                                int.TryParse(valorCelula, out auxInt);
                                                celula.Valor = auxInt;
                                                break;
                                            case TipoCelula.Texto:
                                                celula.Valor = valorCelula;
                                                break;
                                            default:
                                                throw new ArgumentOutOfRangeException();
                                        }
                                        #endregion

                                        #region Escrever valores

                                        saidaSheet.Cells[linhaSaida, 1].Value = celulaFixa;
                                        saidaSheet.Cells[linhaSaida, celula.ColSaida].Value = celula.Valor;

                                        #endregion
                                    }
                                }
                                #endregion
                            }
                            myBook.Close();
                            if (saidaExiste)
                            {
                                saidaBook.Save();
                            }
                            else
                            {
                                saidaBook.SaveAs(arquivoSaida);
                            }
                            saidaBook.Close();

                            progressBar1.Increment(1);
                            //File.Delete(arquivo);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }
                    }
                    MessageBox.Show(Resources.tarefa_concluida);
                    myExcelApp.Quit();
                }

            }
        }

        private void AdicionaTitulo(ref Excel.Worksheet sheet, Planilha planilha)
        {
            sheet.Cells[1, 1].Value = "NomeEscola";
            foreach (var celula in planilha.Celulas)
            {
                sheet.Cells[1, celula.ColSaida].Value = celula.Titulo;
            }
        }
    }
}
