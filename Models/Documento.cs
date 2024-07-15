namespace Docx.Models
{
    public class Documento 
    {
        public string Titulo { get; set; }
        public string SubTitulo { get; set; }
        public List<Seccion> Secciones  { get; set; }
        public string Bibliografia { get; set; }

        public Documento(string titulo, string subTitulo, List<Seccion> secciones, string bibliografia)
        {
            Titulo = titulo;
            SubTitulo = subTitulo;
            Secciones = secciones ?? new List<Seccion>();
            Bibliografia = bibliografia;
        }

        public void AgregarSeccion(Seccion seccion)
        {
            Secciones.Add(seccion);
        }
    }
}
