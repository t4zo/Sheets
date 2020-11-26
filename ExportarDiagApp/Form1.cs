using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExportarDiagApp
{
    public partial class Form1 : Form
    {
        public Form1()
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

        #region Declarações
        private static Excel.Application MyApp = null;
        private static Excel.Workbook MyBook = null;

        private static Excel.Workbook SaidaBook = null;

        private static Excel.Worksheet MySheetNota10 = null;

        private static Excel.Worksheet SaidaNota10 = null;

        private static List<Nota10> nota10List = null;


        private const int startRow = 8;
        private const string Senha = "+4@N0$";
        private const string Senha2 = "leobrisio23";
        #endregion

        //Aqui é onde a mágica acontece!!!
        private void btnGerarArquivo_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(txtDirEntrada.Text) || String.IsNullOrEmpty(txtDirSaida.Text))
            {//Se os diretórios não estiverem selecionados, exibe a mensagem de erro
                MessageBox.Show("Selecione os diretórios de Entrada e Saída");
            }
            else//Se tudo estiver certo, inicia a procura dos arquivos
            {
                try
                {
                    //Cria a lista de arquivos encontrados no diretório de entrada
                    string[] arquivos = Directory.GetFiles(txtDirEntrada.Text, "*.xls*");
                    if (arquivos.Count() == 0)//Se não encontrar arquivo exibe a mensagem de erro
                    {
                        MessageBox.Show("Nenhum arquivo encontrado em " + txtDirEntrada);
                    }
                    else//Caso encontre arquivos
                    {
                        progressBar1.Maximum = arquivos.Count();
                        foreach (var arquivo in arquivos)//Passa por todos os arquivos
                        {
                            try
                            {
                                MyApp = new Excel.Application();
                                MyApp.Visible = false;
                                MyApp.DisplayAlerts = false;

                                if (!arquivo.Contains("~$"))
                                {
                                    MyBook = MyApp.Workbooks.Open(arquivo);

                                    Nota10.Abrir();

                                    Nota10.Ler();

                                    MyBook.Close();//Fecha o arquivo após a leitura

                                    string ArquivoSaida = txtDirSaida.Text + "Saida-Nota10.xlsx";
                                    int rowNota10 = 0;

                                    AbrirSaida(
                                        ArquivoSaida,
                                        ref  rowNota10
                                    );

                                    Nota10.Escrever(rowNota10);

                                    if (File.Exists(ArquivoSaida))
                                    {
                                        RenomearPlanilhasSaida();
                                        SaidaBook.Save();
                                    }
                                    else
                                    {
                                        SaidaBook.SaveAs(ArquivoSaida);
                                    }
                                    SaidaBook.Close();
                                }
                                progressBar1.Increment(1);
                                File.Delete(arquivo);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message);
                            }
                            finally
                            {
                                //SaidaBook.Close();
                                //MyBook.Close();
                            }
                        }
                    }
                    MyApp.Quit();//Sai da aplicação no excel
                    MessageBox.Show("Concluído com sucesso!");
                }

                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message + "\n" + ex.StackTrace);
                }
                finally
                {
                    MyApp.Quit();
                }
            }
        }

        #region Auxiliares
        private static void RenomearPlanilhasSaida()
        {
            SaidaNota10.Name = "Nota10";
        }

        private static void AbrirSaida(
            string ArquivoSaida,
            ref int rowNota10
            )
        {
            if (File.Exists(ArquivoSaida))
            {
                SaidaBook = MyApp.Workbooks.Open(ArquivoSaida);

                SaidaNota10 = (Excel.Worksheet)SaidaBook.Worksheets[1];

                rowNota10 = SaidaNota10.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            }
            else
            {
                SaidaBook = MyApp.Workbooks.Add();

                SaidaNota10 = (Excel.Worksheet)SaidaBook.Worksheets[1];

                Nota10.Titulo();

                rowNota10 = 1;
            }
        }

        private class Nota10
        {
            #region propriedades
            public int Unidade { get; set; }
            public int TurmaID { get; set; }
            public int Inicio { get; set; }
            public int Entrada { get; set; }
            public int Saida { get; set; }
            public int Recebido { get; set; }
            public int Expedido { get; set; }
            public int Falecido { get; set; }
            public int Fim { get; set; }
            public int Especiais { get; set; }
            public int Evadido { get; set; }

            public int AAMPort { get; set; }
            public int AAMMat { get; set; }
            public int AAMHist { get; set; }
            public int AAMGeo { get; set; }
            public int AAMCien { get; set; }
            public int AAMArt { get; set; }
            public int AAMEdFis { get; set; }
            public int AAMIng { get; set; }
            public int AAMEmp { get; set; }
            public int AAMRel { get; set; }
            #endregion

            public static void Abrir()
            {
                MySheetNota10 = (Excel.Worksheet)MyBook.Sheets[1];
                try
                {
                    MySheetNota10.Unprotect(Senha);
                }
                catch (Exception)
                {
                    MySheetNota10.Unprotect(Senha2);
                }

                nota10List = new List<Nota10>();
            }
            public static void Ler()
            {
                int lastRow = MySheetNota10.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;

                for (int row = startRow; row <= lastRow; row++)
                {
                    if (MySheetNota10.Cells[row, 5].Value != null)
                    {
                        #region Declaracao Colunas da Planilha
                        object colSerieID = MySheetNota10.Cells[row, 5].Value;

                        object colUnidade = MySheetNota10.Cells[row, 2].Value;
                        object colTurmaID = MySheetNota10.Cells[row, 4].Value;
                        object colInicio = MySheetNota10.Cells[row, 9].Value;
                        object colEntrada = MySheetNota10.Cells[row, 10].Value;
                        object colSaida = MySheetNota10.Cells[row, 11].Value;
                        object colRecebido = MySheetNota10.Cells[row, 12].Value;
                        object colExpedido = MySheetNota10.Cells[row, 13].Value;
                        object colFalecido = MySheetNota10.Cells[row, 14].Value;
                        object colFim = MySheetNota10.Cells[row, 15].Value;
                        object colEspeciais = MySheetNota10.Cells[row, 16].Value;
                        object colEvadido = MySheetNota10.Cells[row, 17].Value;

                        object colAAMPort = MySheetNota10.Cells[row, 18].Value;
                        object colAAMMat = MySheetNota10.Cells[row, 19].Value;
                        object colAAMHist = MySheetNota10.Cells[row, 20].Value;
                        object colAAMGeo = MySheetNota10.Cells[row, 21].Value;
                        object colAAMCien = MySheetNota10.Cells[row, 22].Value;
                        object colAAMArt = MySheetNota10.Cells[row, 23].Value;
                        object colAAMEdFis = MySheetNota10.Cells[row, 24].Value;
                        object colAAMIng = MySheetNota10.Cells[row, 25].Value;
                        object colAAMEmp = MySheetNota10.Cells[row, 26].Value;
                        object colAAMRel = MySheetNota10.Cells[row, 27].Value;
                        #endregion

                        #region Seta os Valores das Colunas da Planilha
                        //Adiciona a linha na lista dos diagnósticos
                        if (colSerieID != null && !String.IsNullOrEmpty(colSerieID.ToString()))
                        {
                            nota10List.Add(new Nota10
                            {
                                Unidade = colUnidade != null && !String.IsNullOrEmpty(colUnidade.ToString()) ? int.Parse(colUnidade.ToString().Trim()) : 0,
                                TurmaID = colTurmaID != null && !String.IsNullOrEmpty(colTurmaID.ToString()) ? int.Parse(colTurmaID.ToString().Trim()) : 0,
                                Inicio = colInicio != null && !String.IsNullOrEmpty(colInicio.ToString()) ? int.Parse(colInicio.ToString().Trim()) : 0,
                                Entrada = colEntrada != null && !String.IsNullOrEmpty(colEntrada.ToString()) ? int.Parse(colEntrada.ToString().Trim()) : 0,
                                Saida = colSaida != null && !String.IsNullOrEmpty(colSaida.ToString()) ? int.Parse(colSaida.ToString().Trim()) : 0,
                                Recebido = colRecebido != null && !String.IsNullOrEmpty(colRecebido.ToString()) ? int.Parse(colRecebido.ToString().Trim()) : 0,
                                Expedido = colExpedido != null && !String.IsNullOrEmpty(colExpedido.ToString()) ? int.Parse(colExpedido.ToString().Trim()) : 0,
                                Falecido = colFalecido != null && !String.IsNullOrEmpty(colFalecido.ToString()) ? int.Parse(colFalecido.ToString().Trim()) : 0,
                                Fim = colFim != null && !String.IsNullOrEmpty(colFim.ToString()) ? int.Parse(colFim.ToString().Trim()) : 0,
                                Especiais = colEspeciais != null && !String.IsNullOrEmpty(colEspeciais.ToString()) ? int.Parse(colEspeciais.ToString().Trim()) : 0,
                                Evadido = colEvadido != null && !String.IsNullOrEmpty(colEvadido.ToString()) ? int.Parse(colEvadido.ToString().Trim()) : 0,

                                AAMPort = colAAMPort != null && !String.IsNullOrEmpty(colAAMPort.ToString()) ? int.Parse(colAAMPort.ToString().Trim()) : 0,
                                AAMMat = colAAMMat != null && !String.IsNullOrEmpty(colAAMMat.ToString()) ? int.Parse(colAAMMat.ToString().Trim()) : 0,
                                AAMHist = colAAMHist != null && !String.IsNullOrEmpty(colAAMHist.ToString()) ? int.Parse(colAAMHist.ToString().Trim()) : 0,
                                AAMGeo = colAAMGeo != null && !String.IsNullOrEmpty(colAAMGeo.ToString()) ? int.Parse(colAAMGeo.ToString().Trim()) : 0,
                                AAMCien = colAAMCien != null && !String.IsNullOrEmpty(colAAMCien.ToString()) ? int.Parse(colAAMCien.ToString().Trim()) : 0,
                                AAMArt = colAAMArt != null && !String.IsNullOrEmpty(colAAMArt.ToString()) ? int.Parse(colAAMArt.ToString().Trim()) : 0,
                                AAMEdFis = colAAMEdFis != null && !String.IsNullOrEmpty(colAAMEdFis.ToString()) ? int.Parse(colAAMEdFis.ToString().Trim()) : 0,
                                AAMIng = colAAMIng != null && !String.IsNullOrEmpty(colAAMIng.ToString()) ? int.Parse(colAAMIng.ToString().Trim()) : 0,
                                AAMEmp = colAAMEmp != null && !String.IsNullOrEmpty(colAAMEmp.ToString()) ? int.Parse(colAAMEmp.ToString().Trim()) : 0,
                                AAMRel = colAAMRel != null && !String.IsNullOrEmpty(colAAMRel.ToString()) ? int.Parse(colAAMRel.ToString().Trim()) : 0

                            });
                        }
                        #endregion
                    }
                }
            }
            public static void Titulo()
            {
                SaidaNota10.Cells[1, 1].Value = "Unidade";
                SaidaNota10.Cells[1, 2].Value = "TurmaID";
                SaidaNota10.Cells[1, 3].Value = "Inicio";
                SaidaNota10.Cells[1, 4].Value = "Entrada";
                SaidaNota10.Cells[1, 5].Value = "Saida";
                SaidaNota10.Cells[1, 6].Value = "Recebido";
                SaidaNota10.Cells[1, 7].Value = "Expedido";
                SaidaNota10.Cells[1, 8].Value = "Falecido";
                SaidaNota10.Cells[1, 9].Value = "Fim";
                SaidaNota10.Cells[1, 10].Value = "Especiais";
                SaidaNota10.Cells[1, 11].Value = "Evadido";
                SaidaNota10.Cells[1, 12].Value = "AAMPort";
                SaidaNota10.Cells[1, 13].Value = "AAMMat";
                SaidaNota10.Cells[1, 14].Value = "AAMHist";
                SaidaNota10.Cells[1, 15].Value = "AAMGeo";
                SaidaNota10.Cells[1, 16].Value = "AAMCien";
                SaidaNota10.Cells[1, 17].Value = "AAMArt";
                SaidaNota10.Cells[1, 18].Value = "AAMEdFis";
                SaidaNota10.Cells[1, 19].Value = "AAMIng";
                SaidaNota10.Cells[1, 20].Value = "AAMEmp";
                SaidaNota10.Cells[1, 21].Value = "AAMRel";
            }
            public static void Escrever(int row)
            {
                foreach (var item in nota10List)
                {
                    ++row;

                    SaidaNota10.Cells[row, 1].Value = item.Unidade;
                    SaidaNota10.Cells[row, 2].Value = item.TurmaID;
                    SaidaNota10.Cells[row, 3].Value = item.Inicio;
                    SaidaNota10.Cells[row, 4].Value = item.Entrada;
                    SaidaNota10.Cells[row, 5].Value = item.Saida;
                    SaidaNota10.Cells[row, 6].Value = item.Recebido;
                    SaidaNota10.Cells[row, 7].Value = item.Expedido;
                    SaidaNota10.Cells[row, 8].Value = item.Falecido;
                    SaidaNota10.Cells[row, 9].Value = item.Fim;
                    SaidaNota10.Cells[row, 10].Value = item.Especiais;
                    SaidaNota10.Cells[row, 11].Value = item.Evadido;

                    SaidaNota10.Cells[row, 12].Value = item.AAMPort;
                    SaidaNota10.Cells[row, 13].Value = item.AAMMat;
                    SaidaNota10.Cells[row, 14].Value = item.AAMHist;
                    SaidaNota10.Cells[row, 15].Value = item.AAMGeo;
                    SaidaNota10.Cells[row, 16].Value = item.AAMCien;
                    SaidaNota10.Cells[row, 17].Value = item.AAMArt;
                    SaidaNota10.Cells[row, 18].Value = item.AAMEdFis;
                    SaidaNota10.Cells[row, 19].Value = item.AAMIng;
                    SaidaNota10.Cells[row, 20].Value = item.AAMEmp;
                    SaidaNota10.Cells[row, 21].Value = item.AAMRel;
                }
            }
        }
        #endregion
    }
}
