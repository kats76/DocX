namespace Docx.Models
{
    public class Seccion
    {
        public string TituloSeccion { get; set; }
        public string Parrafo { get; set; }
        public IFormFile Imagen { get; set; }
        public string TituloImagen { get; set; }

        public Seccion(string tituloSeccion, string parrafo, IFormFile imagen, string tituloImagen)
        {
            TituloSeccion = tituloSeccion;
            Parrafo = parrafo;
            Imagen = imagen;
            TituloImagen = tituloImagen;
        } 

    }
}