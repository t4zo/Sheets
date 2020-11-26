using System;
using System.Linq;
using System.IO;
using System.Windows.Forms;
//using ExportarDiagApp.Models;
using ExportarDiagApp.Properties;
using Excel = Microsoft.Office.Interop.Excel;
using ExportarDiagApp.Models.Perfil;

namespace ExportarDiagApp
{
    public partial class ExportarPerfil : Form
    {
        public ExportarPerfil()
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
                var arquivoSaida = txtDirSaida.Text + "Saida-PerfilEscola.xlsx";

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

                    var perfil = new PlanilhaPerfilEscola();
                    var planilha = perfil.Planilha;


                    var saidaExiste = File.Exists(arquivoSaida);
                    var saidaBook = saidaExiste ? myExcelApp.Workbooks.Open(arquivoSaida) : myExcelApp.Workbooks.Add();
                    Excel.Worksheet saidaSheet;
                    var linhaSaida = 1;
                    if (saidaExiste)
                    {
                        saidaSheet = (Excel.Worksheet)saidaBook.Sheets[1];
                        linhaSaida = saidaSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    }
                    else
                    {
                        saidaSheet = (Excel.Worksheet)saidaBook.Sheets[1];
                        saidaSheet.Name = planilha.Nome;
                        AdicionaTitulo(ref saidaSheet, planilha);
                    }

                    foreach (var arquivo in arquivos)
                    {
                        linhaSaida++;

                        try
                        {
                            var myBook = myExcelApp.Workbooks.Open(arquivo);
                            
                            var mySheet = (Excel.Worksheet)myBook.Sheets[planilha.Ordem];

                            foreach (var celula in planilha.Celulas)
                            {
                                var valorCelula = mySheet.Cells[celula.Linha, celula.Coluna].Value == null ? string.Empty :  mySheet.Cells[celula.Linha, celula.Coluna].Value.ToString().Trim().ToUpper();

                                switch (celula.Tipo)
                                {
                                    case Models.TipoCelula.Texto:
                                        celula.Valor = valorCelula;
                                        break;
                                    case Models.TipoCelula.Numero:
                                        int auxInt;
                                        int.TryParse(valorCelula, out auxInt);
                                        celula.Valor = auxInt;
                                        break;
                                    default:
                                        throw new ArgumentOutOfRangeException();
                                }

                                saidaSheet.Cells[linhaSaida, celula.ColSaida].Value = celula.Valor;
                            }

                            progressBar1.Increment(1);
                            myBook.Close();
                            //File.Delete(arquivo);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }
                    }
                    if (saidaExiste)
                    {
                        saidaBook.Save();
                    }
                    else
                    {
                        saidaBook.SaveAs(arquivoSaida);
                    }
                    saidaBook.Close();
                    MessageBox.Show(Resources.tarefa_concluida);
                    myExcelApp.Quit();
                }

            }
        }

        private void AdicionaTitulo(ref Excel.Worksheet sheet, Planilha planilha)
        {
            foreach (var celula in planilha.Celulas)
            {
                sheet.Cells[1, celula.ColSaida].Value = celula.Titulo;
            }
        }
    }
}
