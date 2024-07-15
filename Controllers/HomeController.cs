using Microsoft.AspNetCore.Mvc;
public class HomeController : Controller
{
    private readonly Servicio _servicio;

    public HomeController(Servicio servicio)
    {
        _servicio = servicio;
    }

    public IActionResult Index()
    {
        return View();
    }

    [HttpPost]
    public IActionResult CreateDoc(string title, string subtitle, string[] sectionTitles, string[] paragraphs, IFormFile[] images, string[] imageTexts, string bibliography)
    {
        var documento = _servicio.CrearDocumento(title, subtitle, sectionTitles, paragraphs, images, imageTexts, bibliography);
        return _servicio.GenerarDocumento(documento);
    }

}