namespace ExportarDiagApp.Models
{
    public class Celula
    {
        public int Col { get; set; }
        public int ColSaida { get; set; }
        public string Titulo { get; set; }
        public TipoCelula Tipo { get; set; }
        public object Valor { get; set; }
    }
}
