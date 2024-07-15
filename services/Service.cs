using Microsoft.AspNetCore.Mvc;
using Docx.Models;
using Xceed.Words.NET;
using Xceed.Document.NET;

public class Servicio
{
    public Documento CrearDocumento(string titulo, string subTitulo, string[] seccionTitulo, string[] seccionParrafo, IFormFile[] seccionImagen, string[] seccionTituloImagen, string bibliografia)
    {
        // Asegurarse de que las imágenes y los títulos de imágenes tengan la misma longitud que los títulos y párrafos de las secciones
        if (seccionImagen.Length != seccionTitulo.Length)
        {
            Array.Resize(ref seccionImagen, seccionTitulo.Length);
        }
        if (seccionTituloImagen.Length != seccionTitulo.Length)
        {
            Array.Resize(ref seccionTituloImagen, seccionTitulo.Length);
        }

        List<Seccion> secciones = new List<Seccion>();
        for (int i = 0; i < seccionTitulo.Length; i++)
        {
            Seccion seccion = new Seccion(seccionTitulo[i], seccionParrafo[i], seccionImagen[i], seccionTituloImagen[i]);
            secciones.Add(seccion);
        }

        Documento documento = new Documento(titulo, subTitulo, secciones, bibliografia);
        return documento;
    }

    public FileStreamResult GenerarDocumento(Documento doc)
    {
        var fileName = "NuevoDocumento.docx";
        var path = Path.Combine(Directory.GetCurrentDirectory(), fileName);
        var templatePath = Path.Combine(Directory.GetCurrentDirectory(), "template.docx");

        try
        {
            // Cargar la plantilla existente sin modificarla
            using (var template = DocX.Load(templatePath))
            {
                // Crear un nuevo documento basado en la plantilla
                using (var docX = DocX.Create(path))
                {
                    docX.InsertDocument(template);

                    // Reemplazar título y subtítulo en la portada
                    docX.ReplaceText("titulo", doc.Titulo);
                    docX.ReplaceText("subtitulo", doc.SubTitulo);
                    Console.WriteLine("titulo", doc.Titulo);
                    Console.WriteLine("subtitulo", doc.SubTitulo);

                    // Agregar secciones
                    foreach (var seccion in doc.Secciones)
                    {
                        docX.InsertParagraph(seccion.TituloSeccion).Bold().FontSize(16).Heading(HeadingType.Heading1);
                        docX.InsertParagraph(seccion.Parrafo);

                        // Si hay una imagen, agregarla al documento
                        if (seccion.Imagen != null && seccion.Imagen.Length > 0)
                        {
                            using (var stream = new MemoryStream())
                            {
                                seccion.Imagen.CopyTo(stream);
                                stream.Position = 0;
                                var image = docX.AddImage(stream);
                                var picture = image.CreatePicture(200, 400); // Ajusta el tamaño aquí
                                docX.InsertParagraph().AppendPicture(picture).Alignment = Alignment.center;
                                docX.InsertParagraph(seccion.TituloImagen).Alignment = Alignment.center;
                            }
                        }

                        docX.InsertParagraph(); // Párrafo en blanco después de cada sección
                    }

                    // Agregar bibliografía
                    docX.InsertParagraph("Bibliografía").Bold().FontSize(16).Heading(HeadingType.Heading1);
                    docX.InsertParagraph(doc.Bibliografia);

                    // Guardar el documento
                    docX.Save();
                }
            }

            // Descargar el archivo generado
            var fileStream = System.IO.File.OpenRead(path);
            return new FileStreamResult(fileStream, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            {
                FileDownloadName = fileName
            };
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
            throw;
        }
    }
}