using Import_Freight_BOI.Models;
using Import_Freight_BOI.Models.TSQL;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;

namespace Import_Freight_BOI.Controllers
{
    public class HomeController : Controller
    {
        private readonly IConfiguration _configuration;
        public const string SessionID = "";
        public const string Session_fullname = "";
        //protected SqlConnection conn = new SqlConnection("Data source=;Initail Catalog=;user ID; password;");
        //private readonly ILogger<HomeController> _logger;

        public HomeController(IConfiguration configuration)
        {

            _configuration = configuration;

        }
        // public HomeController(ILogger<HomeController> logger)
        // {
        //     // _logger = logger;
        // }

        public IActionResult Index()
        {
            //return View();
            return RedirectToAction("Login", "Manage_Freight");
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



        public IActionResult Logout()
        {
            var session = HttpContext.Session.GetString("SessionID");
            var session_fullname = HttpContext.Session.GetString("Session_fullname");
            if (session != "")
            {
                HttpContext.Session.Remove("SessionID");
                HttpContext.Session.Remove("session_fullname");
            }

            return RedirectToAction("Index", "Home");
        }

      
    }
}
