using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportarDiagApp.Models
{
    class PlanilhasICEP
    {
        public ICollection<Planilha> Planilhas { get; set; }

        public PlanilhasICEP()
        {
            Planilhas = new List<Planilha>
            {
                new Planilha
                {
                    Nome = "Diag. Escrita - Fund. I",
                    Senha = "3dUc@c@0",
                    Ordem = 4,
                    OrdemSaida = 1,
                    CelulaVerif = 2,
                    LinhaInicial = 14,
                    Linhas = new List<Linha>(),
                    Celulas = new List<Celula>
                {
                    new Celula { Titulo = "MesID", Col = 2, Tipo = TipoCelula.Numero, ColSaida = 1 },
                    new Celula { Titulo = "EscolaID", Col = 4, Tipo = TipoCelula.Numero, ColSaida = 2 },
                    new Celula { Titulo = "SerieID", Col = 5, Tipo = TipoCelula.Numero, ColSaida = 3 },
                    new Celula { Titulo = "TGTotal", Col = 6, Tipo = TipoCelula.Numero, ColSaida = 4 },
                    new Celula { Titulo = "TGPreSilabico", Col = 7, Tipo = TipoCelula.Numero, ColSaida = 5 },
                    new Celula { Titulo = "TGPreSilPorcentagem", Col = 8, Tipo = TipoCelula.Texto, ColSaida = 6 },
                    new Celula { Titulo = "TGSilabico1", Col = 9, Tipo = TipoCelula.Numero, ColSaida = 7 },
                    new Celula { Titulo = "TGSil1Porcentagem", Col = 10, Tipo = TipoCelula.Texto, ColSaida = 8 },
                    new Celula { Titulo = "TGSilabico2", Col = 11, Tipo = TipoCelula.Numero, ColSaida = 9 },
                    new Celula { Titulo = "TGSil2Porcentagem", Col = 12, Tipo = TipoCelula.Texto, ColSaida = 10 },
                    new Celula { Titulo = "TGSilabico3", Col = 13, Tipo = TipoCelula.Numero, ColSaida = 11 },
                    new Celula { Titulo = "TGSil3Porcentagem", Col = 14, Tipo = TipoCelula.Texto, ColSaida = 12 },
                    new Celula { Titulo = "TGSilabicoAlfabetico", Col = 15, Tipo = TipoCelula.Numero, ColSaida = 13 },
                    new Celula { Titulo = "TGSilAlfPorcentagem", Col = 16, Tipo = TipoCelula.Texto, ColSaida = 14 },
                    new Celula { Titulo = "TGAlfabetico", Col = 17, Tipo = TipoCelula.Numero, ColSaida = 15 },
                    new Celula { Titulo = "TGAlfPorcentagem", Col = 18, Tipo = TipoCelula.Texto, ColSaida = 16 },
                }
                },
                //new Planilha
                //{
                //    Nome = "Diag. Prod. Textual",
                //    Senha = "3dUc@c@0",
                //    Ordem = 4,
                //    OrdemSaida = 1,
                //    CelulaVerif = 5,
                //    LinhaInicial = 10,
                //    Linhas = new List<Linha>(),
                //    Celulas = new List<Celula>
                //{
                //    new Celula { Titulo = "MesID", Col = 2, Tipo = TipoCelula.Numero, ColSaida = 1 },
                //    new Celula { Titulo = "EscolaID", Col = 4, Tipo = TipoCelula.Numero, ColSaida = 2 },
                //    new Celula { Titulo = "SerieID", Col = 5, Tipo = TipoCelula.Numero, ColSaida = 3 },
                //    new Celula { Titulo = "Turma(s)", Col = 6, Tipo = TipoCelula.Texto, ColSaida = 4 },
                //    new Celula { Titulo = "QtdAlunos", Col = 7, Tipo = TipoCelula.Numero, ColSaida = 5 },
                //    new Celula { Titulo = "CGContoH1", Col = 8, Tipo = TipoCelula.Numero, ColSaida = 6 },
                //    new Celula { Titulo = "CGContoH2", Col = 9, Tipo = TipoCelula.Numero, ColSaida = 7 },
                //    new Celula { Titulo = "CGContoH3", Col = 10, Tipo = TipoCelula.Numero, ColSaida = 8 },
                //    new Celula { Titulo = "CGResumoH1", Col = 11, Tipo = TipoCelula.Numero, ColSaida = 9 },
                //    new Celula { Titulo = "CGResumoH2", Col = 12, Tipo = TipoCelula.Numero, ColSaida = 10 },
                //    new Celula { Titulo = "CGResumoH3", Col = 13, Tipo = TipoCelula.Numero, ColSaida = 11 },
                //    new Celula { Titulo = "CGResumoH4", Col = 14, Tipo = TipoCelula.Numero, ColSaida = 12 },
                //    new Celula { Titulo = "CGRecomLiterH1", Col = 15, Tipo = TipoCelula.Numero, ColSaida = 13 },
                //    new Celula { Titulo = "CGRecomLiterH2", Col = 16, Tipo = TipoCelula.Numero, ColSaida = 14 },
                //    new Celula { Titulo = "CGRecomLiterH3", Col = 17, Tipo = TipoCelula.Numero, ColSaida = 15 },
                //    new Celula { Titulo = "PEH1", Col = 18, Tipo = TipoCelula.Numero, ColSaida = 16 },
                //    new Celula { Titulo = "PEH2", Col = 19, Tipo = TipoCelula.Numero, ColSaida = 17 },
                //    new Celula { Titulo = "PEH3", Col = 20, Tipo = TipoCelula.Numero, ColSaida = 18 },
                //    new Celula { Titulo = "PEH4", Col = 21, Tipo = TipoCelula.Numero, ColSaida = 19 },
                //    new Celula { Titulo = "PEH5", Col = 22, Tipo = TipoCelula.Numero, ColSaida = 20 },
                //    new Celula { Titulo = "PEH6", Col = 23, Tipo = TipoCelula.Numero, ColSaida = 21 },
                //    new Celula { Titulo = "PEH7", Col = 24, Tipo = TipoCelula.Numero, ColSaida = 22 }
                //}
                //},
            };
        }
    }
}
