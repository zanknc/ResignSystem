using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System.Diagnostics;
using Import_Freight_BOI.Models.TSQL;
using OfficeOpenXml;
using ClosedXML.Excel;
using System.IO;
using Microsoft.AspNetCore.Http;
using System.Net;
using System.Net.Http;
using MySql.Data.MySqlClient.Memcached;

namespace Import_Freight_BOI.Controllers
{
    public class Manage_FreightController : Controller
    {
        public const string SessionID = "";
        public const string Session_fullname = "";
        private readonly IConfiguration _configuration;
        public Manage_FreightController(IConfiguration configuration)
        {

         
            _configuration = configuration;
           
        }
        public IActionResult Import_Freight()
        {
            
            String Strsql = "";
            var ObjRun = new mgrSQLConnect(_configuration);
            DataTable dt = new DataTable();

            Strsql = "SELECT AccountNo ,Title FROM AccountTitle";
            ViewBag.AccountTitle = ObjRun.GetDatatables(Strsql);

            Strsql = "Select * FROM vewCountry";
            ViewBag.Country = ObjRun.GetDatatables(Strsql);

            Strsql = "Select * FROM vewDelivery";
            ViewBag.Delivery = ObjRun.GetDatatables(Strsql);

            Strsql = "Select * FROM vewForwarder";
            ViewBag.Forwarder = ObjRun.GetDatatables(Strsql);

            Strsql = "Select * FROM vewSupplier";
            ViewBag.Supplier = ObjRun.GetDatatables(Strsql);

            Strsql = "Select Value1 FROM Purpose where code= '002'";
            dt = ObjRun.GetDatatables(Strsql);
            ViewBag.Purpose = dt.Rows[0][0];

            Strsql = "Select * FROM vewTransportation";
            ViewBag.Transportation = ObjRun.GetDatatables(Strsql);

            string Session = HttpContext.Session.GetString(SessionID);
            ViewBag.SessionID = Session;

            
           
            if (Session == null)
            {
                return RedirectToAction("Index","Home");
            }
            else
            {
                return View();
            }
           
        }

        private string CheckGFDReport(int OPID)
        {
         
            string status = "";
            using (var httpClient = new HttpClient())
            {
                string url1 = "";
               
            }
            return "";
        }

        public IActionResult Report_Section()
        {
            string Session = HttpContext.Session.GetString(SessionID);

            if (Session == null)
            {
                return RedirectToAction("Index", "Home");
            }
            else
            {


                string contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                string fileName = "Report_Section.xlsx";
                try
                {
                    using (var workbook = new XLWorkbook())
                    {
                        IXLWorksheet worksheet =
                        workbook.Worksheets.Add("Section");
                        worksheet.Cell(1, 1).Value = "Export Frieght "+ '"' + "sort by Section"+'"' + "";
                        worksheet.Cell(1, 2).Value = "TYPE";
                        worksheet.Cell(1, 3).Value = "LastName";
                        worksheet.Cell(1, 1).Style.Font.FontSize = 20;
                        worksheet.Cell(1, 1).Style.Font.SetBold();



                        //required using System.IO;
                        using (var stream = new MemoryStream())
                        {
                            workbook.SaveAs(stream);
                            var content = stream.ToArray();
                            return File(content, contentType, fileName);
                        }
                    }




                }
                catch (Exception ex)
                {
                    throw ex;
                }



                return View();
            }
        }

        public IActionResult Login()
        {
      
            return View();
        }
        public IActionResult Export_data()
        {
            string Session = HttpContext.Session.GetString(SessionID);

            if (Session == null)
            {
                return RedirectToAction("Index", "Home");
            }
            else
            {

                string contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                string fileName = "Import freight.xlsx";
                try
                {
                    using (var workbook = new XLWorkbook())
                    {
                        IXLWorksheet worksheet =
                        workbook.Worksheets.Add("freight");
                        worksheet.Cell(1, 1).Value = "Id";
                        worksheet.Cell(1, 2).Value = "FirstName";
                        worksheet.Cell(1, 3).Value = "LastName";

                        workbook.Worksheets.Add("A");
                        worksheet.Cell(1, 1).Value = "Id";
                        worksheet.Cell(1, 2).Value = "FirstName";
                        worksheet.Cell(1, 3).Value = "LastName";

                        //required using System.IO;
                        using (var stream = new MemoryStream())
                        {
                            workbook.SaveAs(stream);
                            var content = stream.ToArray();
                            return File(content, contentType, fileName);
                        }
                    }
                }
                catch (Exception ex)
                {
                    throw ex;
                }

                return View();

            }
        }


        public IActionResult Report_by_DN()
        {

            string Session = HttpContext.Session.GetString(SessionID);

            if (Session == null)
            {
                return RedirectToAction("Index", "Home");
            }
            else
            {

                string contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                string fileName = "Report_DN.xlsx";
                try
                {
                    using (var workbook = new XLWorkbook())
                    {
                        IXLWorksheet worksheet =
                        workbook.Worksheets.Add("DN");
                        worksheet.Cell(1, 1).Value = "ID";
                        worksheet.Cell(1, 2).Value = "D/N";
                        worksheet.Cell(1, 3).Value = "TITLE";
                        worksheet.Cell(1, 4).Value = "SURSENSE";
                        worksheet.Cell(1, 5).Value = "FREIGHT";
                        worksheet.Cell(1, 6).Value = "SHIPPING";
                        worksheet.Cell(1, 7).Value = "ORIGINAL";
                        worksheet.Cell(1, 8).Value = "VAT1";
                        worksheet.Cell(1, 9).Value = "BASE TAX3";
                        worksheet.Cell(1, 10).Value = "BASE TAX1";
                        worksheet.Cell(1, 11).Value = "TYPE";

                        //workbook.Worksheets.Add("A");
                        //worksheet.Cell(1, 1).Value = "Id";
                        //worksheet.Cell(1, 2).Value = "FirstName";
                        //worksheet.Cell(1, 3).Value = "LastName";

                        //required using System.IO;
                        using (var stream = new MemoryStream())
                        {
                            workbook.SaveAs(stream);
                            var content = stream.ToArray();
                            return File(content, contentType, fileName);
                        }
                    }
                }
                catch (Exception ex)
                {
                    throw ex;
                }

                return View();

            }
        }


        public JsonResult GetValue_DropdowList()
        {
            String Strsql = "";
            var ObjRun = new mgrSQLConnect(_configuration);
            DataTable dt_Country = new DataTable();
            DataTable dt_Delivery = new DataTable();
            DataTable dt_Forwarder = new DataTable();
            DataTable dt_Purpose = new DataTable();
            DataTable dt_Supplier = new DataTable();


            Strsql = "Select * FROM vewCountry";
            dt_Country = ObjRun.GetDatatables(Strsql);

            Strsql = "Select * FROM vewDelivery";
            dt_Delivery = ObjRun.GetDatatables(Strsql);

            Strsql = "Select * FROM vewForwarder";
            dt_Forwarder = ObjRun.GetDatatables(Strsql);

            Strsql = "Select * FROM vewPurpose";
            dt_Purpose = ObjRun.GetDatatables(Strsql);

            Strsql = "Select * FROM vewSupplier";
            dt_Supplier = ObjRun.GetDatatables(Strsql);


            var JsonResult =  Json(new{ Country = dt_Country , Delivery = dt_Delivery , Forwarder = dt_Forwarder, Purpose = dt_Purpose, Supplier = dt_Supplier});
           return JsonResult;
        }

        public JsonResult CheckUser_Login(string OPID, string Password)
        {
            string status;
            var query = new mgrSQLConnect(_configuration);
          
            string strQuery = "select * from vewOperator where OperatorID = '" + OPID.Trim() + "' and Password = '" + Password.Trim() + "'";
            var checkUser = query.GetDatatables(strQuery);
            if (checkUser.Rows.Count != 0)
            {
                HttpContext.Session.SetString(SessionID, checkUser.Rows[0][0].ToString());
                HttpContext.Session.SetString(Session_fullname, checkUser.Rows[0][1].ToString());
                status = "True";
                //return RedirectToPage(nameof(HomeController.Index), "Home");
            }
            else
            {
                status = "False";
            }

            return Json(status);
        }

        public JsonResult Insert_Header(string DN , string YearMonth ,string Forwarder, string Invoice, string RentDay, string Vat, string Transport, string Delivery, string Supplier,
                                        string Vm, string ETA ,string Storage, string Country)
        {
            string Session = HttpContext.Session.GetString(SessionID);
            var Sqlquery = new mgrSQLConnect(_configuration);
            DataTable dt = new DataTable();
            string status;
            string lastID;
            try
            {
               
                string Sqlstr = "Exec sprOperation_InvoiceImportHeader '" + DN + "','" + YearMonth + "','" + Forwarder + "','" + Invoice + "','" + RentDay +  "','" + Transport + "','" + Delivery + "','" + Supplier + "','" + Vm + "','" + ETA + "','" + Storage + "','" + Country + "','" + Vat + "','" + Session + "'";
                var ExcuteProc = Sqlquery.ExcuteProc(Sqlstr);

                 dt = Sqlquery.GetDatatables("select Max(DNID) from ImportFreightHead");
                lastID = dt.Rows[0][0].ToString();

                status = "True";

            } catch (Exception ex){
                throw ex;

                status = "False";
            }

            var JsonResult = Json(new {status=  status , lastID = lastID });
            return JsonResult;
        }

        //public JsonResult Insert_Freight_detail(string title , string terminal, string License, string Freight, string  Origin, string Shipping, string Vat1 , string Btax3 , string Btax1, string Vat2 , string Advance, string APamount)
        //{

        //    
        //    return Json("");
        //}
        public JsonResult getDropdownTitle(string Data)
        {
            var Sqlquery = new mgrSQLConnect(_configuration);

            DataTable dt = new DataTable();
           string Strsql = "SELECT AccountNo ,Title FROM AccountTitle";
            dt = Sqlquery.GetDatatables(Strsql);
            return Json(dt);
        }

        public JsonResult Insert_Freight_detail(String[] RowData)
        {

            string status, lastID;
            var Sqlquery = new mgrSQLConnect(_configuration);
            DataTable dt = new DataTable();
            

            try
            {
                

                for (int i =0; i< RowData.Length; i++)
                {
                    dt = Sqlquery.GetDatatables("select Max(DNID) from ImportFreightHead");
                    lastID = dt.Rows[0][0].ToString().Trim();
                    string addchar = "'" + lastID + "','" + RowData[i].Replace(",", "','");
                    string Sqlstr = "Exec sprOperation_FreightImportDetail " + addchar + "'";
                    var ExcuteProc = Sqlquery.ExcuteProc(Sqlstr);


                }
                status = "True";
            

            }
            catch (Exception ex)
            {
                status = "False";
            }

            var JsonResult = Json(new { status = status });
            return JsonResult;
        }

    }

   
}
