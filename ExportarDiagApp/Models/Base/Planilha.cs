using System.Collections.Generic;

namespace ExportarDiagApp.Models
{
    public class Planilha
    {
        public int Ordem { get; set; }
        public int OrdemSaida { get; set; }
        public string Senha { get; set; }
        public string Nome { get; set; }
        public int CelulaVerif { get; set; }
        public int LinhaInicial { get; set; }
        public ICollection<Linha> Linhas { get; set; }
        public ICollection<Celula> Celulas { get; set; }

        public int LinhaFixa { get; set; }
        public int ColunaFixa { get; set; }
    }
}
