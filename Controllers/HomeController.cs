using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.IO;
using System.Data;
using System.Data.OleDb;
using Roullete.Models;
using Microsoft.Office.Interop;
using Excel = Microsoft.Office.Interop.Excel;

namespace Roulette_Application.Controllers
{
    public class HomeController : Controller
    {
        //
        // GET: /Home/
        public ActionResult ExcelData()
        {
            DataTable dt = new DataTable();
            string filePath = Server.MapPath("~/ExcelFile/");
            string filepaths = Server.MapPath("~/ExcelFile/");
            string excelfilepath = "";
            if (Directory.Exists(filePath))
            {
                string[] files = Directory.GetFiles(filePath);
                foreach (string f in files)
                {
                    excelfilepath = f.ToString();
                }
            }
            return View();
        }
        [HttpGet]
        public ActionResult ExcelDataResponse()
        { 
            return View(); 
        }

        [HttpPost]
        public ActionResult ExcelDataResponse(HttpPostedFileBase file)
        {
            try
            {
                string filepath = Server.MapPath("~/ExcelFile/");
                if (Directory.Exists(filepath))
                {
                    string[] files = Directory.GetFiles(filepath);
                    foreach (string f in files)
                    {
                        try
                        {
                            System.IO.File.Delete(f);
                            //FileInfo FInfo = new FileInfo(f);
                            //IsFileLocked(FInfo);
                        }
                        catch {
                            //The process cannot access the file 'E:\working\projects\Testing Projects\Roulette Application\Roulette Application\ExcelFile\Book1.xlsx' because it is being used by another process
                        }
                    }
                }
                //string filename = System.IO.Path.GetFileName(file.FileName);
                string filename = "Book1.xlsx";
                file.SaveAs(Server.MapPath("~/ExcelFile/" + filename));
                string filepathtosave = "ExcelFile/" + filename;
                ViewBag.Message = "File Uploaded successfully.";
            }
            catch
            {
                ViewBag.Message = "Error while uploading the files.";
            }
            return View();
        }

        protected virtual bool IsFileLocked(FileInfo file)
        {
            FileStream stream = null;

            try
            {
                stream = file.Open(FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            }
            catch (IOException)
            {
                //the file is unavailable because it is:
                //still being written to
                //or being processed by another thread
                //or does not exist (has already been processed)
                return true;
            }
            finally
            {
                if (stream != null)
                    stream.Close();
            }

            //file is not locked
            return false;
        }
    }
}
