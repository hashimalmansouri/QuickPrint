using Microsoft.AspNetCore.Mvc;
using Microsoft.Reporting.NETCore;
using QuickPrint.Extensions;
using QuickPrint.Models;
using System.Diagnostics;
using System.Reflection;

namespace QuickPrint.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Print()
        {
            using var report = new LocalReport();
            using var rs = Assembly.GetExecutingAssembly().GetManifestResourceStream("QuickPrint.wwwroot.reports.rpt.rdlc");
            report.LoadReportDefinition(rs);
            report.PrintToPrinter();
            return Json(true);
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}