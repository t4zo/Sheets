using System.Collections.Generic;

namespace ExportarDiagApp.Models.Perfil
{
    public class Planilha
    {
        public string Nome { get; set; }
        public int Ordem { get; set; }
        public ICollection<Celula> Celulas { get; set; }
    }
}
