using System.Collections.Generic;

namespace ExportarDiagApp.Models
{
    public class PlanilhasDiags
    {
        public ICollection<Planilha> Planilhas { get; set; }

        public PlanilhasDiags()
        {
            Planilhas = new List<Planilha>();

            #region Gerenciamento
            Planilhas.Add(new Planilha
            {
                Nome = "Gerenciamento",
                Senha = "",
                Ordem = 2,
                OrdemSaida = 1,
                CelulaVerif = 6,
                LinhaInicial = 8,
                Linhas = new List<Linha>(),
                Celulas = new List<Celula>
                {
                    new Celula { Titulo = "Unidade", Col = 2, Tipo = TipoCelula.Numero, ColSaida = 1 },
                    new Celula { Titulo = "TurmaID", Col = 4, Tipo = TipoCelula.Numero, ColSaida = 2 },
                    new Celula { Titulo = "Inicio", Col = 9, Tipo = TipoCelula.Numero, ColSaida = 3 },
                    new Celula { Titulo = "Entradas", Col = 10, Tipo = TipoCelula.Numero, ColSaida = 4 },
                    new Celula { Titulo = "Saidas", Col = 11, Tipo = TipoCelula.Numero, ColSaida = 5 },
                    new Celula { Titulo = "Recebidas", Col = 12, Tipo = TipoCelula.Numero, ColSaida = 6 },
                    new Celula { Titulo = "Expedidas", Col = 13, Tipo = TipoCelula.Numero, ColSaida = 7 },
                    new Celula { Titulo = "Falecidos", Col = 14, Tipo = TipoCelula.Numero, ColSaida = 8 },
                    new Celula { Titulo = "Fim", Col = 15, Tipo = TipoCelula.Numero, ColSaida = 9 },
                    new Celula { Titulo = "Evadidos", Col = 16, Tipo = TipoCelula.Numero, ColSaida = 10 },
                    new Celula { Titulo = "Especiais", Col = 17, Tipo = TipoCelula.Numero, ColSaida = 11 }
                }
            });
            #endregion

            #region Leitura
            Planilhas.Add(new Planilha
            {
                Nome = "Leitura",
                Senha = "",
                Ordem = 7,
                OrdemSaida = 2,
                CelulaVerif = 24,
                LinhaInicial = 7,
                Linhas = new List<Linha>(),
                Celulas = new List<Celula>
                {
                    new Celula { Titulo = "Unidade", Col = 2, Tipo = TipoCelula.Numero, ColSaida = 1 },
                    new Celula { Titulo = "TurmaID", Col = 4, Tipo = TipoCelula.Numero, ColSaida = 2 },
                    new Celula { Titulo = "ND", Col = 12, Tipo = TipoCelula.Numero, ColSaida = 3 },
                    new Celula { Titulo = "Q1", Col = 13, Tipo = TipoCelula.Numero, ColSaida = 4 },
                    new Celula { Titulo = "Q2", Col = 14, Tipo = TipoCelula.Numero, ColSaida = 5 },
                    new Celula { Titulo = "Q3", Col = 15, Tipo = TipoCelula.Numero, ColSaida = 6 },
                    new Celula { Titulo = "Q4", Col = 16, Tipo = TipoCelula.Numero, ColSaida = 7 },
                    new Celula { Titulo = "Q5", Col = 17, Tipo = TipoCelula.Numero, ColSaida = 8 },
                    new Celula { Titulo = "Q6", Col = 18, Tipo = TipoCelula.Numero, ColSaida = 9 },
                    new Celula { Titulo = "Q7", Col = 19, Tipo = TipoCelula.Numero, ColSaida = 10 },
                    new Celula { Titulo = "Q8", Col = 20, Tipo = TipoCelula.Numero, ColSaida = 11 },
                    new Celula { Titulo = "Q9", Col = 21, Tipo = TipoCelula.Numero, ColSaida = 12 },
                    new Celula { Titulo = "Q10", Col = 22, Tipo = TipoCelula.Numero, ColSaida = 13 }/*,
                    new Celula { Titulo = "H11", Col = 23, Tipo = TipoCelula.Numero, ColSaida = 14 },
                    new Celula { Titulo = "H12", Col = 24, Tipo = TipoCelula.Numero, ColSaida = 15 },
                    new Celula { Titulo = "H13", Col = 25, Tipo = TipoCelula.Numero, ColSaida = 16 },
                    new Celula { Titulo = "H14", Col = 26, Tipo = TipoCelula.Numero, ColSaida = 17 },
                    new Celula { Titulo = "H15", Col = 27, Tipo = TipoCelula.Numero, ColSaida = 18 },
                    new Celula { Titulo = "H16", Col = 28, Tipo = TipoCelula.Numero, ColSaida = 19 },
                    new Celula { Titulo = "H17", Col = 29, Tipo = TipoCelula.Numero, ColSaida = 20 },
                    new Celula { Titulo = "H18", Col = 30, Tipo = TipoCelula.Numero, ColSaida = 21 },
                    new Celula { Titulo = "H19", Col = 31, Tipo = TipoCelula.Numero, ColSaida = 22 },
                    new Celula { Titulo = "H20", Col = 32, Tipo = TipoCelula.Numero, ColSaida = 23 },
                    new Celula { Titulo = "H21", Col = 33, Tipo = TipoCelula.Numero, ColSaida = 24 },
                    new Celula { Titulo = "H22", Col = 34, Tipo = TipoCelula.Numero, ColSaida = 25 },
                    new Celula { Titulo = "AE", Col = 35, Tipo = TipoCelula.Numero, ColSaida = 26 },
                    new Celula { Titulo = "EC", Col = 36, Tipo = TipoCelula.Numero, ColSaida = 27 },
                    new Celula { Titulo = "C", Col = 37, Tipo = TipoCelula.Numero, ColSaida = 28 }*/
                }
            });
            #endregion

            #region Escrita
            Planilhas.Add(new Planilha
            {
                Nome = "Escrita",
                Senha = "",
                Ordem = 8,
                OrdemSaida = 3,
                CelulaVerif = 28,
                LinhaInicial = 7,
                Linhas = new List<Linha>(),
                Celulas = new List<Celula>
                {
                    new Celula { Titulo = "Unidade", Col = 2, Tipo = TipoCelula.Numero, ColSaida = 1 },
                    new Celula { Titulo = "TurmaID", Col = 4, Tipo = TipoCelula.Numero, ColSaida = 2 },
                    new Celula { Titulo = "ND", Col = 12, Tipo = TipoCelula.Numero, ColSaida = 3 },
                    new Celula { Titulo = "N1", Col = 13, Tipo = TipoCelula.Numero, ColSaida = 4 },
                    new Celula { Titulo = "N2", Col = 14, Tipo = TipoCelula.Numero, ColSaida = 5 },
                    new Celula { Titulo = "N3", Col = 15, Tipo = TipoCelula.Numero, ColSaida = 6 },
                    new Celula { Titulo = "N4", Col = 16, Tipo = TipoCelula.Numero, ColSaida = 7 },
                    new Celula { Titulo = "N5", Col = 17, Tipo = TipoCelula.Numero, ColSaida = 8 },
                    new Celula { Titulo = "N6", Col = 18, Tipo = TipoCelula.Numero, ColSaida = 9 },
                    new Celula { Titulo = "N7", Col = 19, Tipo = TipoCelula.Numero, ColSaida = 10 },
                    new Celula { Titulo = "N8", Col = 20, Tipo = TipoCelula.Numero, ColSaida = 11 },
                    new Celula { Titulo = "N9", Col = 21, Tipo = TipoCelula.Numero, ColSaida = 12 },
                    new Celula { Titulo = "N10", Col = 22, Tipo = TipoCelula.Numero, ColSaida = 13 },
                    new Celula { Titulo = "N11", Col = 23, Tipo = TipoCelula.Numero, ColSaida = 14 },
                    new Celula { Titulo = "AE", Col = 24, Tipo = TipoCelula.Numero, ColSaida = 15 },
                    new Celula { Titulo = "EC", Col = 25, Tipo = TipoCelula.Numero, ColSaida = 16 },
                    new Celula { Titulo = "C", Col = 26, Tipo = TipoCelula.Numero, ColSaida = 17 }
                }
            });
            #endregion

            #region ProdIniciais
            Planilhas.Add(new Planilha
            {
                Nome = "Prod Iniciais",
                Senha = "",
                Ordem = 9,
                OrdemSaida = 4,
                CelulaVerif = 33,
                LinhaInicial = 7,
                Linhas = new List<Linha>(),
                Celulas = new List<Celula>
                {
                    new Celula { Titulo = "Unidade", Col = 2, Tipo = TipoCelula.Numero, ColSaida = 1 },
                    new Celula { Titulo = "TurmaID", Col = 4, Tipo = TipoCelula.Numero, ColSaida = 2 },
                    new Celula { Titulo = "ND", Col = 12, Tipo = TipoCelula.Numero, ColSaida = 3 },
                    new Celula { Titulo = "H1", Col = 13, Tipo = TipoCelula.Numero, ColSaida = 4 },
                    new Celula { Titulo = "H2", Col = 14, Tipo = TipoCelula.Numero, ColSaida = 5 },
                    new Celula { Titulo = "H3", Col = 15, Tipo = TipoCelula.Numero, ColSaida = 6 },
                    new Celula { Titulo = "H4", Col = 16, Tipo = TipoCelula.Numero, ColSaida = 7 },
                    new Celula { Titulo = "H5", Col = 17, Tipo = TipoCelula.Numero, ColSaida = 8 },
                    new Celula { Titulo = "H6", Col = 18, Tipo = TipoCelula.Numero, ColSaida = 9 },
                    new Celula { Titulo = "H7", Col = 19, Tipo = TipoCelula.Numero, ColSaida = 10 },
                    new Celula { Titulo = "H8", Col = 20, Tipo = TipoCelula.Numero, ColSaida = 11 },
                    new Celula { Titulo = "H9", Col = 21, Tipo = TipoCelula.Numero, ColSaida = 12 },
                    new Celula { Titulo = "H10", Col = 22, Tipo = TipoCelula.Numero, ColSaida = 13 },
                    new Celula { Titulo = "H11", Col = 23, Tipo = TipoCelula.Numero, ColSaida = 14 },
                    new Celula { Titulo = "H12", Col = 24, Tipo = TipoCelula.Numero, ColSaida = 15 },
                    new Celula { Titulo = "H13", Col = 25, Tipo = TipoCelula.Numero, ColSaida = 16 },
                    new Celula { Titulo = "H14", Col = 26, Tipo = TipoCelula.Numero, ColSaida = 17 },
                    new Celula { Titulo = "H15", Col = 27, Tipo = TipoCelula.Numero, ColSaida = 18 },
                    new Celula { Titulo = "H16", Col = 28, Tipo = TipoCelula.Numero, ColSaida = 19 },
                    new Celula { Titulo = "AE", Col = 29, Tipo = TipoCelula.Numero, ColSaida = 20 },
                    new Celula { Titulo = "EC", Col = 30, Tipo = TipoCelula.Numero, ColSaida = 21 },
                    new Celula { Titulo = "C", Col = 31, Tipo = TipoCelula.Numero, ColSaida = 22 }
                }
            });
            #endregion

            #region Matematica
            Planilhas.Add(new Planilha
            {
                Nome = "Matematica",
                Senha = "",
                Ordem = 11,
                OrdemSaida = 5,
                CelulaVerif = 11,
                LinhaInicial = 7,
                Linhas = new List<Linha>(),
                Celulas = new List<Celula>
                {
                    new Celula { Titulo = "Unidade", Col = 2, Tipo = TipoCelula.Numero, ColSaida = 1 },
                    new Celula { Titulo = "TurmaID", Col = 4, Tipo = TipoCelula.Numero, ColSaida = 2 },
                    new Celula { Titulo = "ND", Col = 12, Tipo = TipoCelula.Numero, ColSaida = 3 },
                    new Celula { Titulo = "Q1", Col = 13, Tipo = TipoCelula.Numero, ColSaida = 4 },
                    new Celula { Titulo = "Q2", Col = 14, Tipo = TipoCelula.Numero, ColSaida = 5 },
                    new Celula { Titulo = "Q3", Col = 15, Tipo = TipoCelula.Numero, ColSaida = 6 },
                    new Celula { Titulo = "Q4", Col = 16, Tipo = TipoCelula.Numero, ColSaida = 7 },
                    new Celula { Titulo = "QH5", Col = 17, Tipo = TipoCelula.Numero, ColSaida = 8 },
                    new Celula { Titulo = "Q6", Col = 18, Tipo = TipoCelula.Numero, ColSaida = 9 },
                    new Celula { Titulo = "Q7", Col = 19, Tipo = TipoCelula.Numero, ColSaida = 10 },
                    /*new Celula { Titulo = "H8", Col = 20, Tipo = TipoCelula.Numero, ColSaida = 11 },
                    new Celula { Titulo = "H9", Col = 21, Tipo = TipoCelula.Numero, ColSaida = 12 },
                    new Celula { Titulo = "H10", Col = 22, Tipo = TipoCelula.Numero, ColSaida = 13 },
                    new Celula { Titulo = "H11", Col = 23, Tipo = TipoCelula.Numero, ColSaida = 14 },*/
                    new Celula { Titulo = "Total", Col = 20, Tipo = TipoCelula.Numero, ColSaida = 11 }
                }
            });
            #endregion

            #region Finais
            /*Planilhas.Add(new Planilha
            {
                Nome = "Finais",
                Senha = "123456",
                Ordem = 12,
                OrdemSaida = 6,
                CelulaVerif = 23,
                LinhaInicial = 7,
                Linhas = new List<Linha>(),
                Celulas = new List<Celula>
                {
                    new Celula { Titulo = "Unidade", Col = 2, Tipo = TipoCelula.Numero, ColSaida = 1 },
                    new Celula { Titulo = "TurmaID", Col = 4, Tipo = TipoCelula.Numero, ColSaida = 2 },
                    new Celula { Titulo = "ND", Col = 12, Tipo = TipoCelula.Numero, ColSaida = 3 },
                    new Celula { Titulo = "P1", Col = 13, Tipo = TipoCelula.Numero, ColSaida = 4 },
                    new Celula { Titulo = "P2", Col = 14, Tipo = TipoCelula.Numero, ColSaida = 5 },
                    new Celula { Titulo = "P3", Col = 15, Tipo = TipoCelula.Numero, ColSaida = 6 },
                    new Celula { Titulo = "P4", Col = 16, Tipo = TipoCelula.Numero, ColSaida = 7 },
                    new Celula { Titulo = "P5", Col = 17, Tipo = TipoCelula.Numero, ColSaida = 8 },
                    new Celula { Titulo = "P6", Col = 18, Tipo = TipoCelula.Numero, ColSaida = 9 },
                    new Celula { Titulo = "P7", Col = 19, Tipo = TipoCelula.Numero, ColSaida = 10 },
                    new Celula { Titulo = "P8", Col = 20, Tipo = TipoCelula.Numero, ColSaida = 11 },
                    new Celula { Titulo = "P9", Col = 21, Tipo = TipoCelula.Numero, ColSaida = 12 },
                    new Celula { Titulo = "P10", Col = 22, Tipo = TipoCelula.Numero, ColSaida = 13 },
                    new Celula { Titulo = "PTotal", Col = 23, Tipo = TipoCelula.Numero, ColSaida = 14 },
                    new Celula { Titulo = "M1", Col = 24, Tipo = TipoCelula.Numero, ColSaida = 15 },
                    new Celula { Titulo = "M2", Col = 25, Tipo = TipoCelula.Numero, ColSaida = 16 },
                    new Celula { Titulo = "M3", Col = 26, Tipo = TipoCelula.Numero, ColSaida = 17 },
                    new Celula { Titulo = "M4", Col = 27, Tipo = TipoCelula.Numero, ColSaida = 18 },
                    new Celula { Titulo = "M5", Col = 28, Tipo = TipoCelula.Numero, ColSaida = 19 },
                    new Celula { Titulo = "M6", Col = 29, Tipo = TipoCelula.Numero, ColSaida = 20 },
                    new Celula { Titulo = "M7", Col = 30, Tipo = TipoCelula.Numero, ColSaida = 21 },
                    new Celula { Titulo = "M8", Col = 31, Tipo = TipoCelula.Numero, ColSaida = 22 },
                    new Celula { Titulo = "M9", Col = 32, Tipo = TipoCelula.Numero, ColSaida = 23 },
                    new Celula { Titulo = "M10", Col = 33, Tipo = TipoCelula.Numero, ColSaida = 24 },
                    new Celula { Titulo = "MTotal", Col = 34, Tipo = TipoCelula.Numero, ColSaida = 25 }
                }
            });*/
            #endregion
        }
    }
}
