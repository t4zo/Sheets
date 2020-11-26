using System.Collections.Generic;

namespace ExportarDiagApp.Models.Perfil
{
    public class PlanilhaPerfilEscola
    {
        public Planilha Planilha { get; set; }

        public PlanilhaPerfilEscola()
        {
            Planilha = new Planilha
            {
                Nome = "Perfil_Escola",
                Ordem = 1,
                Celulas = new List<Celula>
                {
                    new Celula { Titulo = "NomeEscola", Tipo = TipoCelula.Texto, Coluna = 1, Linha = 2, ColSaida = 1 },
                    new Celula { Titulo = "Distrito", Tipo = TipoCelula.Texto, Coluna = 1, Linha = 5, ColSaida = 2 },
                    new Celula { Titulo = "Endereco", Tipo = TipoCelula.Texto, Coluna = 4, Linha = 5, ColSaida = 3 },
                    new Celula { Titulo = "Bairro", Tipo = TipoCelula.Texto, Coluna = 10, Linha = 5, ColSaida = 4 },
                    new Celula { Titulo = "Numero", Tipo = TipoCelula.Texto, Coluna = 1, Linha = 8, ColSaida = 5 },
                    new Celula { Titulo = "Complemento", Tipo = TipoCelula.Texto, Coluna = 2, Linha = 8, ColSaida = 6 },
                    new Celula { Titulo = "Email", Tipo = TipoCelula.Texto, Coluna = 4, Linha = 8, ColSaida = 7 },
                    new Celula { Titulo = "Telefone1", Tipo = TipoCelula.Texto, Coluna = 10, Linha = 8, ColSaida = 8 },
                    new Celula { Titulo = "Telefone2", Tipo = TipoCelula.Texto, Coluna = 12, Linha = 8, ColSaida = 9 },
                    new Celula { Titulo = "Telefone3", Tipo = TipoCelula.Texto, Coluna = 14, Linha = 8, ColSaida = 10 },
                    new Celula { Titulo = "GestorNome", Tipo = TipoCelula.Texto, Coluna = 3, Linha = 13, ColSaida = 11 },
                    new Celula { Titulo = "GestorTelefone", Tipo = TipoCelula.Texto, Coluna = 9, Linha = 13, ColSaida = 12 },
                    new Celula { Titulo = "GestorEmail", Tipo = TipoCelula.Texto, Coluna = 11, Linha = 13, ColSaida = 13 },
                    new Celula { Titulo = "ViceNome", Tipo = TipoCelula.Texto, Coluna = 3, Linha = 14, ColSaida = 14 },
                    new Celula { Titulo = "ViceTelefone", Tipo = TipoCelula.Texto, Coluna = 9, Linha = 14, ColSaida = 15 },
                    new Celula { Titulo = "ViceEmail", Tipo = TipoCelula.Texto, Coluna = 11, Linha = 14, ColSaida = 16 },
                    new Celula { Titulo = "SecretarioNome", Tipo = TipoCelula.Texto, Coluna = 3, Linha = 15, ColSaida = 17 },
                    new Celula { Titulo = "SecretarioTelefone", Tipo = TipoCelula.Texto, Coluna = 9, Linha = 15, ColSaida = 18 },
                    new Celula { Titulo = "SecretarioEmail", Tipo = TipoCelula.Texto, Coluna = 11, Linha = 15, ColSaida = 19 },
                    new Celula { Titulo = "Coord1Nome", Tipo = TipoCelula.Texto, Coluna = 3, Linha = 16, ColSaida = 20 },
                    new Celula { Titulo = "Coord1Telefone", Tipo = TipoCelula.Texto, Coluna = 9, Linha = 16, ColSaida = 21 },
                    new Celula { Titulo = "Coord1Email", Tipo = TipoCelula.Texto, Coluna = 11, Linha = 16, ColSaida = 22 },
                    new Celula { Titulo = "Coord2Nome", Tipo = TipoCelula.Texto, Coluna = 3, Linha = 17, ColSaida = 23 },
                    new Celula { Titulo = "Coord2Telefone", Tipo = TipoCelula.Texto, Coluna = 9, Linha = 17, ColSaida = 24 },
                    new Celula { Titulo = "Coord2Email", Tipo = TipoCelula.Texto, Coluna = 11, Linha = 17, ColSaida = 25 },
                    new Celula { Titulo = "AETJNome", Tipo = TipoCelula.Texto, Coluna = 3, Linha = 18, ColSaida = 26 },
                    new Celula { Titulo = "AETJTelefone", Tipo = TipoCelula.Texto, Coluna = 9, Linha = 18, ColSaida = 27 },
                    new Celula { Titulo = "AETJEmail", Tipo = TipoCelula.Texto, Coluna = 11, Linha = 18, ColSaida = 28 },
                }
            };
        }
    }
}
