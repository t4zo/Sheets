using System.Collections.Generic;

namespace ExportarDiagApp.Models
{
    public class Linha
    {
        public int Numero { get; set; }
        public ICollection<Celula> Celulas { get; set; }

        public Linha(int numero, ICollection<Celula> celulas )
        {
            Numero = numero;
            Celulas = celulas;
        }
    }
}
