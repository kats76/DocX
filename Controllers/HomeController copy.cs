using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Http;
using Xceed.Document.NET;
using Xceed.Words.NET;
using System.IO;
using System.Collections.Generic;

public class HomeControllerc : Controller
{
    public IActionResult Index()
    {
        return View();
    }

    [HttpPost]
    public IActionResult CreateDoc(string name, string title, string subtitle, string[] sectionTitles, string[] paragraphs, IFormFile image, string imageText, string bibliography, string closing)
    {
        var fileName = name + ".docx";
        var path = Path.Combine(Directory.GetCurrentDirectory(), fileName);

        using (var doc = DocX.Create(path))
        {
            // Crear la portada
            doc.InsertParagraph(title).FontSize(20).Bold().Alignment = Alignment.center;
            doc.InsertParagraph(subtitle).FontSize(16).Italic().Alignment = Alignment.center;
            doc.InsertParagraph().InsertPageBreakAfterSelf();

            // Crear el índice utilizando el método recomendado
            var tocSwitches = new Dictionary<TableOfContentsSwitches, string>
            {
                { TableOfContentsSwitches.O, "" }
            };
            doc.InsertTableOfContents("Índice", tocSwitches);
            doc.InsertParagraph().InsertPageBreakAfterSelf();

            // Crear secciones con títulos y párrafos
            for (int i = 0; i < sectionTitles.Length; i++)
            {
                var sectionTitle = sectionTitles[i];
                var paragraph = paragraphs[i];
                doc.InsertParagraph(sectionTitle).Heading(HeadingType.Heading1);
                doc.InsertParagraph(paragraph);
            }

            // Insertar imagen
            if (image != null)
            {
                var imageStream = new MemoryStream();
                image.CopyTo(imageStream);
                imageStream.Position = 0;
                var picture = doc.AddImage(imageStream).CreatePicture();
                doc.InsertParagraph().AppendPicture(picture).Alignment = Alignment.center;
                doc.InsertParagraph(imageText).Alignment = Alignment.center;
            }

            // Crear la sección de bibliografía
            doc.InsertParagraph("Referencias bibliográficas").Heading(HeadingType.Heading1);
            doc.InsertParagraph(bibliography);

            // Crear la hoja de cierre
            doc.InsertParagraph(closing).InsertPageBreakAfterSelf();

            // Guardar el documento
            doc.Save();
        }

        var fileStream = System.IO.File.OpenRead(path);
        return File(fileStream, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", fileName);
    }
}
