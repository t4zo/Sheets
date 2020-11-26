namespace ExportarDiagApp.Models.Perfil
{
    public class Celula
    {
        public string Titulo { get; set; }
        public TipoCelula Tipo { get; set; }
        public object Valor { get; set; }
        public int Linha { get; set; }
        public int Coluna { get; set; }
        public int ColSaida { get; set; }
    }
}
