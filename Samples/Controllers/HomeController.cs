using Microsoft.AspNetCore.Mvc;

namespace DDDN.Office.ODT.Samples
{
    public class HomeController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }
    }
}
