using System.Collections.Generic;

namespace ExportarDiagApp.Models
{
    public class PlanilhasMovimentoRendimento
    {
        public ICollection<Planilha> Planilhas { get; set; }

        public PlanilhasMovimentoRendimento()
        {
            Planilhas = new List<Planilha>
            {
                //new Planilha
                //{
                //    Nome = "MovimentoRedimento",
                //    Senha = "",
                //    Ordem = 1,
                //    OrdemSaida = 1,
                //    CelulaVerif = 2,
                //    LinhaInicial = 25,
                //    Linhas = new List<Linha>(),
                //    Celulas = new List<Celula>
                //    {
                //        new Celula { Titulo = "INEP", Col = 2, Tipo = TipoCelula.Numero, ColSaida = 1 },
                //        new Celula { Titulo = "Escola", Col = 3, Tipo = TipoCelula.Numero, ColSaida = 2 },
                //        new Celula { Titulo = "Aprovados", Col = 9, Tipo = TipoCelula.Texto, ColSaida = 3 },
                //        new Celula { Titulo = "Reprovados", Col = 10, Tipo = TipoCelula.Numero, ColSaida = 4 },
                //        new Celula { Titulo = "Transferidos", Col = 12, Tipo = TipoCelula.Numero, ColSaida = 5 },
                //        new Celula { Titulo = "Evadidos", Col = 13, Tipo = TipoCelula.Numero, ColSaida = 6 },
                //        new Celula { Titulo = "Falecidos", Col = 14, Tipo = TipoCelula.Numero, ColSaida = 7 },
                //        new Celula { Titulo = "CASM", Col = 15, Tipo = TipoCelula.Numero, ColSaida = 8 },
                //    }
                //},
                //new Planilha
                //{
                //    Nome = "MovimentoRedimento2018",
                //    Senha = "",
                //    Ordem = 1,
                //    OrdemSaida = 1,
                //    CelulaVerif = 27,
                //    LinhaInicial = 24,
                //    Linhas = new List<Linha>(),
                //    Celulas = new List<Celula>
                //    {
                //        new Celula { Titulo = "INEP", Col = 1, Tipo = TipoCelula.Texto, ColSaida = 1 },
                //        new Celula { Titulo = "RM", Col = 8, Tipo = TipoCelula.Texto, ColSaida = 2 },
                //        new Celula { Titulo = "1º Ano", Col = 9, Tipo = TipoCelula.Numero, ColSaida = 3 },
                //        new Celula { Titulo = "2º Ano", Col = 10, Tipo = TipoCelula.Numero, ColSaida = 4 },
                //        new Celula { Titulo = "3º Ano", Col = 11, Tipo = TipoCelula.Numero, ColSaida = 5 },
                //        new Celula { Titulo = "4º Ano", Col = 12, Tipo = TipoCelula.Numero, ColSaida = 6 },
                //        new Celula { Titulo = "5º Ano", Col = 13, Tipo = TipoCelula.Numero, ColSaida = 7 },
                //        new Celula { Titulo = "6º Ano", Col = 15, Tipo = TipoCelula.Numero, ColSaida = 8 },
                //        new Celula { Titulo = "7º Ano", Col = 16, Tipo = TipoCelula.Numero, ColSaida = 9 },
                //        new Celula { Titulo = "8º Ano", Col = 17, Tipo = TipoCelula.Numero, ColSaida = 10 },
                //        new Celula { Titulo = "9º Ano", Col = 18, Tipo = TipoCelula.Numero, ColSaida = 11 },
                //    }
                //},
                //new Planilha
                //{
                //    Nome = "Redimento2019",
                //    Senha = "",
                //    Ordem = 1,
                //    OrdemSaida = 1,
                //    CelulaVerif = 27,
                //    LinhaInicial = 24,
                //    Linhas = new List<Linha>(),
                //    Celulas = new List<Celula>
                //    {
                //        new Celula { Titulo = "INEP", Col = 1, Tipo = TipoCelula.Texto, ColSaida = 1 },
                //        new Celula { Titulo = "RM", Col = 8, Tipo = TipoCelula.Texto, ColSaida = 2 },
                //        new Celula { Titulo = "1º Ano", Col = 9, Tipo = TipoCelula.Numero, ColSaida = 3 },
                //        new Celula { Titulo = "2º Ano", Col = 10, Tipo = TipoCelula.Numero, ColSaida = 4 },
                //        new Celula { Titulo = "3º Ano", Col = 11, Tipo = TipoCelula.Numero, ColSaida = 5 },
                //        new Celula { Titulo = "4º Ano", Col = 12, Tipo = TipoCelula.Numero, ColSaida = 6 },
                //        new Celula { Titulo = "5º Ano", Col = 13, Tipo = TipoCelula.Numero, ColSaida = 7 },
                //        new Celula { Titulo = "6º Ano", Col = 15, Tipo = TipoCelula.Numero, ColSaida = 8 },
                //        new Celula { Titulo = "7º Ano", Col = 16, Tipo = TipoCelula.Numero, ColSaida = 9 },
                //        new Celula { Titulo = "8º Ano", Col = 17, Tipo = TipoCelula.Numero, ColSaida = 10 },
                //        new Celula { Titulo = "9º Ano", Col = 18, Tipo = TipoCelula.Numero, ColSaida = 11 },
                //    }
                //},
                new Planilha
                {
                    Nome = "MovimentoRedimento2019",
                    Senha = "",
                    Ordem = 1,
                    OrdemSaida = 1,
                    CelulaVerif = 46,
                    LinhaInicial = 25,
                    Linhas = new List<Linha>(),
                    Celulas = new List<Celula>
                    {
                        new Celula { Titulo = "INEP", Col = 1, Tipo = TipoCelula.Texto, ColSaida = 1 },
                        new Celula { Titulo = "RM", Col = 8, Tipo = TipoCelula.Texto, ColSaida = 2 },
                        new Celula { Titulo = "Creche", Col = 9, Tipo = TipoCelula.Numero, ColSaida = 3 },
                        new Celula { Titulo = "Pre-Escola", Col = 10, Tipo = TipoCelula.Numero, ColSaida = 4 },
                        new Celula { Titulo = "1º Ano", Col = 12, Tipo = TipoCelula.Numero, ColSaida = 5 },
                        new Celula { Titulo = "2º Ano", Col = 13, Tipo = TipoCelula.Numero, ColSaida = 6 },
                        new Celula { Titulo = "3º Ano", Col = 14, Tipo = TipoCelula.Numero, ColSaida = 7 },
                        new Celula { Titulo = "4º Ano", Col = 15, Tipo = TipoCelula.Numero, ColSaida = 8 },
                        new Celula { Titulo = "5º Ano", Col = 16, Tipo = TipoCelula.Numero, ColSaida = 9 },
                        new Celula { Titulo = "6º Ano", Col = 18, Tipo = TipoCelula.Numero, ColSaida = 10 },
                        new Celula { Titulo = "7º Ano", Col = 19, Tipo = TipoCelula.Numero, ColSaida = 11 },
                        new Celula { Titulo = "8º Ano", Col = 20, Tipo = TipoCelula.Numero, ColSaida = 12 },
                        new Celula { Titulo = "9º Ano", Col = 21, Tipo = TipoCelula.Numero, ColSaida = 13 },
                        new Celula { Titulo = "Anos Iniciais", Col = 42, Tipo = TipoCelula.Numero, ColSaida = 14 },
                        new Celula { Titulo = "Anos Finais", Col = 43, Tipo = TipoCelula.Numero, ColSaida = 15 },
                    }
                },
            };
        }
    }
}
