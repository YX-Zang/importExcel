using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.IO;
using OfficeOpenXml;
using ReadExcel.Models;
using System.Data;

namespace ImportExcel.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
        public ActionResult ReadExcelUsingEpplus()
        {
            return View();
        }
        [HttpPost]
        public ActionResult ReadExcel(HttpPostedFileBase  upload)
        {
            if (Path.GetExtension(upload.FileName)==".xlsx" || Path.GetExtension(upload.FileName) == ".xls")
            {
                ExcelPackage package = new ExcelPackage(upload.InputStream);
                DataTable Dt = ExcelPackageExtersions.ToDataTable(package);                
            }
                return View();
        }
    }
}