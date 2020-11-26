using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportarDiagApp.Models
{
    public class PlanilhasPerfilTurma
    {
        public ICollection<Planilha> Planilhas { get; set; }

        public PlanilhasPerfilTurma()
        {
            Planilhas = new List<Planilha>();

            Planilhas.Add(new Planilha
            {
                Nome = "Perfil_Turma",
                Senha = "",
                Ordem = 1,
                OrdemSaida = 1,
                LinhaInicial = 6,
                LinhaFixa = 2,
                ColunaFixa = 1,
                CelulaVerif = 5,
                Linhas = new List<Linha>(),
                Celulas = new List<Celula>
                {
                    new Celula { Titulo = "Ano", Col = 1, Tipo = TipoCelula.Texto, ColSaida = 2 },
                    new Celula { Titulo = "Turma", Col = 2, Tipo = TipoCelula.Texto, ColSaida = 3 },
                    new Celula { Titulo = "Turno", Col = 3, Tipo = TipoCelula.Texto, ColSaida = 4 },
                    new Celula { Titulo = "Forma", Col = 4, Tipo = TipoCelula.Texto, ColSaida = 5 },
                    new Celula { Titulo = "QtdAlunos", Col = 5, Tipo = TipoCelula.Numero, ColSaida = 6 },
                    new Celula { Titulo = "QtdMeninos", Col = 6, Tipo = TipoCelula.Numero, ColSaida = 7 },
                    new Celula { Titulo = "QtdMeninas", Col = 7, Tipo = TipoCelula.Numero, ColSaida = 8 },
                    new Celula { Titulo = "Qtd0Ano", Col = 8, Tipo = TipoCelula.Numero, ColSaida = 9 },
                    new Celula { Titulo = "Qtd1Ano", Col = 9, Tipo = TipoCelula.Numero, ColSaida = 10 },
                    new Celula { Titulo = "Qtd2Anos", Col = 10, Tipo = TipoCelula.Numero, ColSaida = 11 },
                    new Celula { Titulo = "Qtd3Anos", Col = 11, Tipo = TipoCelula.Numero, ColSaida = 12 },
                    new Celula { Titulo = "Qtd4Anos", Col = 12, Tipo = TipoCelula.Numero, ColSaida = 13 },
                    new Celula { Titulo = "Qtd5Anos", Col = 13, Tipo = TipoCelula.Numero, ColSaida = 14 },
                    new Celula { Titulo = "Qtd6Anos", Col = 14, Tipo = TipoCelula.Numero, ColSaida = 15 },
                    new Celula { Titulo = "Qtd7Anos", Col = 15, Tipo = TipoCelula.Numero, ColSaida = 16 },
                    new Celula { Titulo = "Qtd8Anos", Col = 16, Tipo = TipoCelula.Numero, ColSaida = 17 },
                    new Celula { Titulo = "Qtd9Anos", Col = 17, Tipo = TipoCelula.Numero, ColSaida = 18 },
                    new Celula { Titulo = "Qtd10Anos", Col = 18, Tipo = TipoCelula.Numero, ColSaida = 19 },
                    new Celula { Titulo = "Qtd11Anos", Col = 19, Tipo = TipoCelula.Numero, ColSaida = 20 },
                    new Celula { Titulo = "Qtd12Anos", Col = 20, Tipo = TipoCelula.Numero, ColSaida = 21 },
                    new Celula { Titulo = "Qtd13Anos", Col = 21, Tipo = TipoCelula.Numero, ColSaida = 22 },
                    new Celula { Titulo = "Qtd14Anos", Col = 22, Tipo = TipoCelula.Numero, ColSaida = 23 },
                    new Celula { Titulo = "Qtd15Anos", Col = 23, Tipo = TipoCelula.Numero, ColSaida = 24 },
                    new Celula { Titulo = "QtdMaior15Anos", Col = 24, Tipo = TipoCelula.Numero, ColSaida = 25 },
                }
            });
        }
    }
}
