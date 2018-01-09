using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using ExcelDataReader;
using Excel_Parsing_Experimenting.Models;

namespace Excel_Parsing_Experimenting.Controllers
{
    public class UploadController : Controller
    {
        //// GET: Upload
        //public ActionResult Index()
        //{
        //    return View();
        //}

        public ActionResult ImportFitbitData(HttpPostedFile file)
        {
            var viewModel = new FitbitData();
            if (file.ContentLength > 0)
            {
                var reader = ExcelReaderFactory.CreateBinaryReader(file.InputStream);
                
            }
            return View()
        }
    }
}