using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;

namespace PruebaComceptoCrearDocumentos.Pages
{
    public class DownloadModel : PageModel
    {
        private readonly IWebHostEnvironment _env;

        public DownloadModel(IWebHostEnvironment env)
        {
            _env = env;
        }
        public IActionResult OnGet(string file)
        {
            if (string.IsNullOrEmpty(file)){ return NotFound(); }


            var filePath = Path.Combine(@$"A:\Programas\programas de c#\Creacion de documentos en web\PruebaComceptoCrearDocumentos\PruebaComceptoCrearDocumentos", "PDF", file);

            byte[] fileBytes = System.IO.File.ReadAllBytes(filePath);

            return File(fileBytes, "application/force-download", file);
        }
    }
}
