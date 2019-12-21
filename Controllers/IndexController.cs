using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Roullete.Models;
using Microsoft.Office.Interop;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data;

namespace Roulette_Application.Controllers
{
    public class IndexController : Controller
    {
        //
        // GET: /Index/
        private static List<int> _wins = new List<int>();
        private static DataTable table = new DataTable();
        public static int Prioritynum;
        //Error:Retrieving the COM class factory for component with CLSID {00024500-0000-0000-C000-000000000046} failed due to the following error: 80070005 Access is denied. (Exception from HRESULT: 0x80070005 (E_ACCESSDENIED))
        public ActionResult Index()
        {
            try
            {
                table.Clear();
                _wins.Clear();
                if (Session["Precsion_Number"] == null)
                {
                    Session["Precsion_Number"] = ExcelToDataTable();
                }
                return View();
            }
            catch (Exception ex)
            { return View(); }
        }

        public ActionResult Add(int id)
        {
            try
            {
                Winnumbers win = new Winnumbers();
                win.Priority_High_Low = Prioritynum;
                List<int> TempWin = new List<int>();
                if (id != null)
                {
                    if (_wins.Count < 9)
                    {
                        _wins.Add(id);
                    }
                    else
                    {
                        _wins.RemoveAt(0);
                        _wins.Add(id);
                    }
                    if (_wins.Count > 8)
                    {
                        foreach (int num in _wins)
                        {
                            TempWin = GetExcelDatas(num);
                            if (TempWin != null)
                            {
                                #region Core Logic for Precision Making system
                                if (Prioritynum == 2)//Get 'Low' Priority Logic
                                {
                                    #region Section 1  0 to 12
                                    if (win.Section1 == 50)
                                    {
                                        foreach (int leastNum in TempWin)
                                        {
                                            if (leastNum == 1 || leastNum == 4 || leastNum == 7 || leastNum == 10)
                                            {
                                                if (win.Section1 == 50)
                                                {
                                                    win.Section1 = leastNum;
                                                }
                                                else if (win.Section1 > leastNum)
                                                {
                                                    win.Section1 = leastNum;
                                                }
                                            }
                                        }
                                    }
                                    if (win.Section2 == 50)
                                    {
                                        foreach (int leastNum in TempWin)
                                        {
                                            if (leastNum == 0 || leastNum == 2 || leastNum == 5 || leastNum == 8 || leastNum == 11)
                                            {
                                                if (win.Section2 == 50)
                                                {
                                                    win.Section2 = leastNum;
                                                }
                                                else if (win.Section2 > leastNum)
                                                {
                                                    win.Section2 = leastNum;
                                                }
                                            }
                                        }
                                    }
                                    if (win.Section3 == 50)
                                    {
                                        foreach (int leastNum in TempWin)
                                        {
                                            if (leastNum == 3 || leastNum == 6 || leastNum == 9 || leastNum == 12)
                                            {
                                                if (win.Section3 == 50)
                                                {
                                                    win.Section3 = leastNum;
                                                }
                                                else if (win.Section3 > leastNum)
                                                {
                                                    win.Section3 = leastNum;
                                                }
                                            }
                                        }
                                    }
                                    #endregion
                                    #region Section 2  13 To 24
                                    if (win.Section4 == 50)
                                    {
                                        foreach (int leastNum in TempWin)
                                        {
                                            if (leastNum == 13 || leastNum == 16 || leastNum == 19 || leastNum == 22)
                                            {
                                                if (win.Section4 == 50)
                                                {
                                                    win.Section4 = leastNum;
                                                }
                                                else if (win.Section4 > leastNum)
                                                {
                                                    win.Section4 = leastNum;
                                                }
                                            }
                                        }
                                    }
                                    if (win.Section5 == 50)
                                    {
                                        foreach (int leastNum in TempWin)
                                        {
                                            if (leastNum == 14 || leastNum == 17 || leastNum == 20 || leastNum == 23)
                                            {
                                                if (win.Section5 == 50)
                                                {
                                                    win.Section5 = leastNum;
                                                }
                                                else if (win.Section5 > leastNum)
                                                {
                                                    win.Section5 = leastNum;
                                                }
                                            }
                                        }
                                    }
                                    if (win.Section6 == 50)
                                    {
                                        foreach (int leastNum in TempWin)
                                        {
                                            if (leastNum == 15 || leastNum == 18 || leastNum == 21 || leastNum == 24)
                                            {
                                                if (win.Section6 == 50)
                                                {
                                                    win.Section6 = leastNum;
                                                }
                                                else if (win.Section6 > leastNum)
                                                {
                                                    win.Section6 = leastNum;
                                                }
                                            }
                                        }
                                    }
                                    #endregion
                                    #region Section 3  25 TO 36
                                    if (win.Section7 == 50)
                                    {
                                        foreach (int leastNum in TempWin)
                                        {
                                            if (leastNum == 25 || leastNum == 28 || leastNum == 31 || leastNum == 34)
                                            {
                                                if (win.Section7 == 50)
                                                {
                                                    win.Section7 = leastNum;
                                                }
                                                else if (win.Section7 > leastNum)
                                                {
                                                    win.Section7 = leastNum;
                                                }
                                            }
                                        }
                                    }
                                    if (win.Section8 == 50)
                                    {
                                        foreach (int leastNum in TempWin)
                                        {
                                            if (leastNum == 26 || leastNum == 29 || leastNum == 32 || leastNum == 35)
                                            {
                                                if (win.Section8 == 50)
                                                {
                                                    win.Section8 = leastNum;
                                                }
                                                else if (win.Section8 > leastNum)
                                                {
                                                    win.Section8 = leastNum;
                                                }
                                            }
                                        }
                                    }
                                    if (win.Section9 == 50)
                                    {
                                        foreach (int leastNum in TempWin)
                                        {
                                            if (leastNum == 27 || leastNum == 30 || leastNum == 33 || leastNum == 36)
                                            {
                                                if (win.Section9 == 50)
                                                {
                                                    win.Section9 = leastNum;
                                                }
                                                else if (win.Section9 > leastNum)
                                                {
                                                    win.Section9 = leastNum;
                                                }
                                            }
                                        }
                                    }
                                    #endregion
                                }
                                else if (Prioritynum == 3)//Get 'High' Priority Logic
                                {
                                    #region Section 1  0 to 12
                                    if (win.Section1 == 50)
                                    {
                                        foreach (int leastNum in TempWin)
                                        {
                                            if (leastNum == 1 || leastNum == 4 || leastNum == 7 || leastNum == 10)
                                            {
                                                if (win.Section1 == 50)
                                                {
                                                    win.Section1 = leastNum;
                                                }
                                                else if (win.Section1 < leastNum)
                                                {
                                                    win.Section1 = leastNum;
                                                }
                                            }
                                        }
                                    }
                                    if (win.Section2 == 50)
                                    {
                                        foreach (int leastNum in TempWin)
                                        {
                                            if (leastNum == 0 || leastNum == 2 || leastNum == 5 || leastNum == 8 || leastNum == 11)
                                            {
                                                if (win.Section2 == 50)
                                                {
                                                    win.Section2 = leastNum;
                                                }
                                                else if (win.Section2 < leastNum)
                                                {
                                                    win.Section2 = leastNum;
                                                }
                                            }
                                        }
                                    }
                                    if (win.Section3 == 50)
                                    {
                                        foreach (int leastNum in TempWin)
                                        {
                                            if (leastNum == 3 || leastNum == 6 || leastNum == 9 || leastNum == 12)
                                            {
                                                if (win.Section3 == 50)
                                                {
                                                    win.Section3 = leastNum;
                                                }
                                                else if (win.Section3 < leastNum)
                                                {
                                                    win.Section3 = leastNum;
                                                }
                                            }
                                        }
                                    }
                                    #endregion
                                    #region Section 2  13 To 24
                                    if (win.Section4 == 50)
                                    {
                                        foreach (int leastNum in TempWin)
                                        {
                                            if (leastNum == 13 || leastNum == 16 || leastNum == 19 || leastNum == 22)
                                            {
                                                if (win.Section4 == 50)
                                                {
                                                    win.Section4 = leastNum;
                                                }
                                                else if (win.Section4 < leastNum)
                                                {
                                                    win.Section4 = leastNum;
                                                }
                                            }
                                        }
                                    }
                                    if (win.Section5 == 50)
                                    {
                                        foreach (int leastNum in TempWin)
                                        {
                                            if (leastNum == 14 || leastNum == 17 || leastNum == 20 || leastNum == 23)
                                            {
                                                if (win.Section5 == 50)
                                                {
                                                    win.Section5 = leastNum;
                                                }
                                                else if (win.Section5 < leastNum)
                                                {
                                                    win.Section5 = leastNum;
                                                }
                                            }
                                        }
                                    }
                                    if (win.Section6 == 50)
                                    {
                                        foreach (int leastNum in TempWin)
                                        {
                                            if (leastNum == 15 || leastNum == 18 || leastNum == 21 || leastNum == 24)
                                            {
                                                if (win.Section6 == 50)
                                                {
                                                    win.Section6 = leastNum;
                                                }
                                                else if (win.Section6 < leastNum)
                                                {
                                                    win.Section6 = leastNum;
                                                }
                                            }
                                        }
                                    }
                                    #endregion
                                    #region Section 3  25 TO 36
                                    if (win.Section7 == 50)
                                    {
                                        foreach (int leastNum in TempWin)
                                        {
                                            if (leastNum == 25 || leastNum == 28 || leastNum == 31 || leastNum == 34)
                                            {
                                                if (win.Section7 == 50)
                                                {
                                                    win.Section7 = leastNum;
                                                }
                                                else if (win.Section7 < leastNum)
                                                {
                                                    win.Section7 = leastNum;
                                                }
                                            }
                                        }
                                    }
                                    if (win.Section8 == 50)
                                    {
                                        foreach (int leastNum in TempWin)
                                        {
                                            if (leastNum == 26 || leastNum == 29 || leastNum == 32 || leastNum == 35)
                                            {
                                                if (win.Section8 == 50)
                                                {
                                                    win.Section8 = leastNum;
                                                }
                                                else if (win.Section8 < leastNum)
                                                {
                                                    win.Section8 = leastNum;
                                                }
                                            }
                                        }
                                    }
                                    if (win.Section9 == 50)
                                    {
                                        foreach (int leastNum in TempWin)
                                        {
                                            if (leastNum == 27 || leastNum == 30 || leastNum == 33 || leastNum == 36)
                                            {
                                                if (win.Section9 == 50)
                                                {
                                                    win.Section9 = leastNum;
                                                }
                                                else if (win.Section9 < leastNum)
                                                {
                                                    win.Section9 = leastNum;
                                                }
                                            }
                                        }
                                    }
                                    #endregion
                                }
                                else//Get 'As It is' Priority Logic
                                {
                                    #region Section 1  0 to 12
                                    if (win.Section1 == 50)
                                    {
                                        foreach (int leastNum in TempWin)
                                        {
                                            if (leastNum == 1 || leastNum == 4 || leastNum == 7 || leastNum == 10)
                                            {
                                                if (win.Section1 == 50)
                                                {
                                                    win.Section1 = leastNum;
                                                }
                                            }
                                        }
                                    }
                                    if (win.Section2 == 50)
                                    {
                                        foreach (int leastNum in TempWin)
                                        {
                                            if (leastNum == 0 || leastNum == 2 || leastNum == 5 || leastNum == 8 || leastNum == 11)
                                            {
                                                if (win.Section2 == 50)
                                                {
                                                    win.Section2 = leastNum;
                                                }
                                            }
                                        }
                                    }
                                    if (win.Section3 == 50)
                                    {
                                        foreach (int leastNum in TempWin)
                                        {
                                            if (leastNum == 3 || leastNum == 6 || leastNum == 9 || leastNum == 12)
                                            {
                                                if (win.Section3 == 50)
                                                {
                                                    win.Section3 = leastNum;
                                                }
                                            }
                                        }
                                    }
                                    #endregion
                                    #region Section 2  13 To 24
                                    if (win.Section4 == 50)
                                    {
                                        foreach (int leastNum in TempWin)
                                        {
                                            if (leastNum == 13 || leastNum == 16 || leastNum == 19 || leastNum == 22)
                                            {
                                                if (win.Section4 == 50)
                                                {
                                                    win.Section4 = leastNum;
                                                }
                                            }
                                        }
                                    }
                                    if (win.Section5 == 50)
                                    {
                                        foreach (int leastNum in TempWin)
                                        {
                                            if (leastNum == 14 || leastNum == 17 || leastNum == 20 || leastNum == 23)
                                            {
                                                if (win.Section5 == 50)
                                                {
                                                    win.Section5 = leastNum;
                                                }
                                            }
                                        }
                                    }
                                    if (win.Section6 == 50)
                                    {
                                        foreach (int leastNum in TempWin)
                                        {
                                            if (leastNum == 15 || leastNum == 18 || leastNum == 21 || leastNum == 24)
                                            {
                                                if (win.Section6 == 50)
                                                {
                                                    win.Section6 = leastNum;
                                                }
                                            }
                                        }
                                    }
                                    #endregion
                                    #region Section 3  25 TO 36
                                    if (win.Section7 == 50)
                                    {
                                        foreach (int leastNum in TempWin)
                                        {
                                            if (leastNum == 25 || leastNum == 28 || leastNum == 31 || leastNum == 34)
                                            {
                                                if (win.Section7 == 50)
                                                {
                                                    win.Section7 = leastNum;
                                                }
                                            }
                                        }
                                    }
                                    if (win.Section8 == 50)
                                    {
                                        foreach (int leastNum in TempWin)
                                        {
                                            if (leastNum == 26 || leastNum == 29 || leastNum == 32 || leastNum == 35)
                                            {
                                                if (win.Section8 == 50)
                                                {
                                                    win.Section8 = leastNum;
                                                }
                                            }
                                        }
                                    }
                                    if (win.Section9 == 50)
                                    {
                                        foreach (int leastNum in TempWin)
                                        {
                                            if (leastNum == 27 || leastNum == 30 || leastNum == 33 || leastNum == 36)
                                            {
                                                if (win.Section9 == 50)
                                                {
                                                    win.Section9 = leastNum;
                                                }
                                            }
                                        }
                                    }
                                    #endregion
                                }
                                #endregion
                                if (win.Section1 != 50 && win.Section2 != 50 && win.Section3 != 50 && win.Section4 != 50 && win.Section5 != 50 && win.Section6 != 50 && win.Section7 != 50
                                    && win.Section8 != 50 && win.Section9 != 50)
                                {
                                    win.PossibleWin.Add(win.Section1);
                                    win.PossibleWin.Add(win.Section2);
                                    win.PossibleWin.Add(win.Section3);
                                    win.PossibleWin.Add(win.Section4);
                                    win.PossibleWin.Add(win.Section5);
                                    win.PossibleWin.Add(win.Section6);
                                    win.PossibleWin.Add(win.Section7);
                                    win.PossibleWin.Add(win.Section8);
                                    win.PossibleWin.Add(win.Section9);
                                    win.GetPrecionsFrom.Add(num);
                                    ViewBag.ErrorMessage = "";
                                    break;
                                }
                                else { ViewBag.ErrorMessage = "incomplete Perceptions Available!"; }
                                win.GetPrecionsFrom.Add(num);
                            }
                            else
                            {
                                ViewBag.ErrorMessage = "incomplete Perceptions Available!";
                                break;
                            }
                        }
                    }
                    win.numbers = _wins;
                    return View("Index", win);
                }
                else
                {
                    ViewBag.ErrorMessage = "incomplete Perceptions Available!";
                    return RedirectToAction("Index", "Index");
                }
            }
            catch (Exception ex)
            {
                ViewBag.ErrorMessage = "incomplete Perceptions Available!";
                return RedirectToAction("Index", "Index");
            }
        }

        public ActionResult Clear()
        {
            try
            {
                ViewBag.Error = "";
                ViewBag.ErrorMessage = "";
                Prioritynum = 0;
                Session.Clear();
                Session["Precsion_Number"] = null;
                _wins.Clear();
                return RedirectToAction("Index", "Index");
            }
            catch (Exception ex)
            { return RedirectToAction("Index", "Index"); }
        }

        public void SetPriorityNum(int id)
        {
            try
            {
                Prioritynum = id;
                ViewBag.prioId = id;
                //return RedirectToAction("Index", "Index");
            }
            catch (Exception ex)
            { //return RedirectToAction("Index", "Index"); }
            }
        }

        private List<int> GetExcelDatas(int HighPriorityNumber)
        {
            try
            {
                if (HighPriorityNumber == 0)
                {
                    HighPriorityNumber = 36;
                }
                else { HighPriorityNumber = HighPriorityNumber - 1; }
                DataTable dt = new DataTable();
                List<int> Winnumbers = new List<int>();
                if (Session["Precsion_Number"] != null)
                {
                    dt = (DataTable)Session["Precsion_Number"];
                }
                else
                {
                    dt = ExcelToDataTable();
                    Session["Precsion_Number"] = dt;
                }
                for (int j = 1; j < dt.Columns.Count; j++)
                {
                    Winnumbers.Add(Convert.ToInt32(dt.Rows[HighPriorityNumber]["" + j + ""]));
                }
                return Winnumbers;
            }
            catch (Exception ex)
            {
                ViewBag.Error = ex.Message;
                return null;
            }
        }

        private DataTable ExcelToDataTable()
        {
            try
            {
                Winnumbers win = new Winnumbers();
                win.Priority_High_Low = Prioritynum;
                DataTable dt = new DataTable();
                //table.Clear();
                table = new DataTable();
                List<int> Winnumbers = new List<int>();
                Excel.Application ex = new Excel.Application();
                string path = Server.MapPath("~/ExcelFile/Book1.xlsx");
                //string path = @"D:/Roullete Files/ExcelFile/Book1.xlsx";
                ViewBag.Path += path;
                ex.Workbooks.Open(path);
                Excel.Worksheet activeWorksheet = ex.ActiveSheet;
                int ExcelCount = 0;
                for (int j = 1; j < 37; j++)
                {
                    Excel.Range currentCell = activeWorksheet.Cells[1, j];
                    var color = currentCell.Font.Color;
                    string colorval = Convert.ToString(color);
                    if (colorval.Contains("255") || j == 1)
                    {
                        ExcelCount += 1;
                    }
                }
                for (int i = 0; i < ExcelCount; i++)
                {
                    table.Columns.Add(i.ToString(), typeof(int));
                }
                for (int i = 1; i < 38; i++)
                {
                    for (int j = 1; j < 37; j++)
                    {
                        Excel.Range currentCell = activeWorksheet.Cells[i, j];
                        var fontFamily = currentCell.Font.Name;
                        var italics = currentCell.Font.Italic;
                        var color = currentCell.Font.Color;
                        string colorval = Convert.ToString(color);
                        if (colorval.Contains("255") || j == 1)
                        {
                            Winnumbers.Add(Convert.ToInt32(currentCell.Value));
                        }
                    }
                    ConvertListToDataTable(Winnumbers,ExcelCount);
                    Winnumbers.Clear();
                }
                activeWorksheet = null;
                ex.Workbooks.Close();
                return table;
            }
            catch (Exception ex)
            {
                ViewBag.Error += ex.Message;
                return null;
            }
        }
        static void ConvertListToDataTable(List<int> list,int Count)
        {
            try
            {
                string[] ColumnValues = new string[Count];
                for (int i = 0; i < list.Count; i++)
                {
                    ColumnValues[i] = Convert.ToString(list[i]);
                }
                table.Rows.Add(ColumnValues);
            }
            catch (Exception ex)
            {
                throw;
            }
        }
    }
}
