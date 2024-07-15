using Microsoft.AspNetCore.Mvc;
using Docx.Models;
using Xceed.Words.NET;
using Xceed.Document.NET;
using System.Drawing;

public class Servicio
{
    public Documento CrearDocumento(string titulo, string subTitulo, IFormFile imagenPortada, string[] seccionTitulo, string[] seccionParrafo, IFormFile[] seccionImagen, string[] seccionTituloImagen, string bibliografia)
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

        Documento documento = new Documento(titulo, subTitulo, imagenPortada, secciones, bibliografia);
        return documento;
    }

    public FileStreamResult GenerarDocumento(Documento doc)
    {
        var fileName = "NuevoDocumento.docx";
        var path = Path.Combine(Directory.GetCurrentDirectory(), fileName);

        try
        {
            // Crear un nuevo documento
            using (var docX = DocX.Create(path))
            {

                // Agregar título
                var titleParagraph = docX.InsertParagraph();
                titleParagraph.Append(doc.Titulo).Bold().FontSize(32).Color(Color.Black);
                titleParagraph.Alignment = Alignment.center;
                titleParagraph.SpacingBefore(100);

                // Crear y agregar la portada con imagen de fondo
                if (doc.ImagenPortada != null)
                {
                    using (var stream = new MemoryStream())
                    {
                        doc.ImagenPortada.CopyTo(stream);
                        stream.Position = 0;
                        var image = docX.AddImage(stream);
                        var picture = image.CreatePicture(200, 400);

                        // Agregar imagen de fondo como párrafo independiente
                        var backgroundParagraph = docX.InsertParagraph();
                        var pictureParagraph = backgroundParagraph.InsertParagraphBeforeSelf("").AppendPicture(picture);
                        pictureParagraph.Alignment = Alignment.center;
                    }
                }

                // Agregar subtítulo debajo del título
                var subTitleParagraph = docX.InsertParagraph();
                subTitleParagraph.Append(doc.SubTitulo).FontSize(24).Color(Color.Black);
                subTitleParagraph.Alignment = Alignment.center;
                subTitleParagraph.SpacingBefore(20);

                docX.InsertParagraph().InsertPageBreakAfterSelf();

                // Insertar un encabezado para la tabla de contenido
                var tocSwitches = new Dictionary<TableOfContentsSwitches, string>
            {
                { TableOfContentsSwitches.O, "1-3" }, // Incluir niveles de encabezado 1 a 3
                { TableOfContentsSwitches.U, "on" }, // Usar hipervínculos
                { TableOfContentsSwitches.Z, "on" }  // Mantener formato original
            };

                // Insertar la tabla de contenido
                docX.InsertTableOfContents("Índice", tocSwitches);
                docX.InsertParagraph().InsertPageBreakAfterSelf();

                // Agregar secciones
                foreach (var seccion in doc.Secciones)
                {
                    var sectionTitle = docX.InsertParagraph(seccion.TituloSeccion).Bold().FontSize(14).Heading(HeadingType.Heading1).Color(Color.Black);
                    sectionTitle.Font("Century Gothic"); // Establecer fuente para el título de la sección

                    var sectionParagraph = docX.InsertParagraph(seccion.Parrafo);
                    sectionParagraph.Font("Century Gothic").FontSize(10); // Establecer fuente para el párrafo de la sección
                    sectionParagraph.Alignment = Alignment.both;

                    docX.InsertParagraph();

                    // Si hay una imagen, agregarla al documento
                    if (seccion.Imagen != null && seccion.Imagen.Length > 0)
                    {
                        using (var stream = new MemoryStream())
                        {
                            seccion.Imagen.CopyTo(stream);
                            stream.Position = 0;
                            var image = docX.AddImage(stream);
                            var picture = image.CreatePicture(200, 400);
                            docX.InsertParagraph().AppendPicture(picture).Alignment = Alignment.center;

                            var imageTitleParagraph = docX.InsertParagraph(seccion.TituloImagen);
                            imageTitleParagraph.Alignment = Alignment.center;
                            imageTitleParagraph.Font("Century Gothic").FontSize(9); // Establecer fuente para el título de la imagen
                        }
                    }

                    docX.InsertParagraph(); // Párrafo en blanco 
                }

                // Agregar bibliografía
                var bibliographyTitleParagraph = docX.InsertParagraph("Referencias bibliográficas").Bold().FontSize(14).Heading(HeadingType.Heading1).Color(Color.Black);
                bibliographyTitleParagraph.Font("Century Gothic").FontSize(14); // Establecer fuente para el título de la bibliografía

                var bibliographyParagraph = docX.InsertParagraph(doc.Bibliografia);
                bibliographyParagraph.Font("Century Gothic").FontSize(10); // Establecer fuente para el contenido de la bibliografía
                bibliographyParagraph.Alignment = Alignment.center;
                docX.InsertParagraph();
                // Guardar el documento
                docX.Save();
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