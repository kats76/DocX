using Microsoft.AspNetCore.Mvc;
using Xceed.Document.NET;
using Xceed.Words.NET;

public class HomeController : Controller
{
	public IActionResult Index()
{
    return View();
}

    [HttpPost]
    public IActionResult CreateDoc(string title, string subtitle, string paragraph, string footer)
    {
        using (var doc = DocX.Create("Documento.docx"))
        {
            // Crear el título
            var titleFormat = new Formatting();
            titleFormat.Size = 18D;
            titleFormat.Position = 12;
            titleFormat.Bold = true;

            doc.InsertParagraph(title, false, titleFormat);

            // Crear el subtítulo
            var subtitleFormat = new Formatting();
            subtitleFormat.Size = 14D;
            subtitleFormat.Position = 10;
            subtitleFormat.Italic = true;

            doc.InsertParagraph(subtitle, false, subtitleFormat);

            // Crear el párrafo
            doc.InsertParagraph(paragraph);

            // Crear el pie de página
            var footerFormat = new Formatting();
            footerFormat.Size = 10D;
            footerFormat.Position = 10;

            doc.AddFooters();
            doc.Footers.Even.InsertParagraph(footer, false, footerFormat);
            doc.Footers.Odd.InsertParagraph(footer, false, footerFormat);

            // Guardar el documento
            doc.Save();
        }

        return RedirectToAction("Index");
    }
}
