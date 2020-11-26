using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportarDiagApp.Models
{
    public class PlanilhaDemonstrativoMatricula
    {
        public ICollection<Planilha> Planilhas { get; set; }

        public PlanilhaDemonstrativoMatricula()
        {
            Planilhas = new List<Planilha>();

            Planilhas.Add(new Planilha
            {
                Nome = "Demonstrativo_Matricula",
                Senha = "",
                Ordem = 1,
                OrdemSaida = 1,
                LinhaInicial = 6,
                //LinhaFixa = 2,
                //ColunaFixa = 2,
                CelulaVerif = 6,
                Linhas = new List<Linha>(),
                Celulas = new List<Celula>
                {
                    new Celula { Titulo = "EscolaID", Col = 1, Tipo = TipoCelula.Numero, ColSaida = 1 },
                    new Celula { Titulo = "Serie", Col = 2, Tipo = TipoCelula.Texto, ColSaida = 2 },
                    new Celula { Titulo = "Turma", Col = 3, Tipo = TipoCelula.Texto, ColSaida = 3 },
                    new Celula { Titulo = "Turno", Col = 4, Tipo = TipoCelula.Texto, ColSaida = 4 },
                    new Celula { Titulo = "Forma", Col = 5, Tipo = TipoCelula.Texto, ColSaida = 5 },
                    new Celula { Titulo = "Total", Col = 6, Tipo = TipoCelula.Numero, ColSaida = 6 },
                    new Celula { Titulo = "AEE", Col = 7, Tipo = TipoCelula.Numero, ColSaida = 7 },
                    new Celula { Titulo = "Transporte", Col = 8, Tipo = TipoCelula.Numero, ColSaida = 8 },
                    new Celula { Titulo = "Vagas", Col = 9, Tipo = TipoCelula.Numero, ColSaida = 9 },
                    new Celula { Titulo = "Obs", Col = 10, Tipo = TipoCelula.Texto, ColSaida = 10 },
                }
            });
        }
    }
}
