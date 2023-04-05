using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Readfromexcel.Models;
namespace Readfromexcel.Controllers
{
    public class ProductController : Controller
    {
        // GET: Product
        public ActionResult Index()
        {
            return View();
        }


        [HttpPost]
        public ActionResult Import(HttpPostedFileBase excelfile)
        {

            if (excelfile == null ||  excelfile.ContentLength == 0)
            {
                ViewBag.Error = "Please Select Excel File";
                return View("Index");
            }

            else
            {
                if(excelfile.FileName.EndsWith("xlsx") || excelfile.FileName.EndsWith("xls")) {

                    string path=Server.MapPath("~/Content/" + excelfile.FileName);
                    if(System.IO.File.Exists(path))
                        System.IO.File.Delete(path);

                    excelfile.SaveAs(path);

                    //Read Data From Excel

                    Excel.Application appli= new Excel.Application();
                   Excel.Workbook work = appli.Workbooks.Open(path);
                    Excel.Worksheet ws = work.ActiveSheet;
                    Excel.Range range = ws.UsedRange;
                    List<Product> products = new List<Product>();
                    for(int row =3;row<=range.Rows.Count;row++)
                    {
                        Product p= new Product();
                        p.Textt = ((Excel.Range)range.Cells[row,1]).Text;
                        products.Add(p);


                    }

                    ViewBag.Products = products;
                     
                    
                    return View("Success");


                }

                else
                {
                    ViewBag.Error = "File type is Incorrect";
                    return View("Index");
                }
            }
        }
    }
}