using System.Web;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ClosedXML.Excel;
using System.IO;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Hosting;
using System.Text;
using System.Data;
using ClosedXML.Excel.Drawings;
using System.Net.Http;
using System.Net.Http.Headers;
using Microsoft.AspNetCore.Http.Extensions;
using System.Configuration;
using System.Data.SqlClient;
using Newtonsoft.Json;
using Import_Freight_BOI.Models.TSQL;
using Microsoft.Extensions.Configuration;
using Microsoft.AspNetCore.Mvc.Rendering;

namespace Import_Freight_BOI.Controllers
{
    
    public class ExcelViewModel
    {
        public string EnrollmentNo { get; set; }
        public string Semester { get; set; }
        public string Month { get; set; }
        public string Year { get; set; }
    }

    public class ResignController : Controller
    {
        private readonly IConfiguration configuration;
        public const string SessionID = "";
        public const string Session_fullname = "";

      
        public class ResignDateList
        {
                public string ResignDate { get; set; }
        }
        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Frm_Resign()
        {

            string fullname = HttpContext.Session.GetString(Session_fullname);
            string Session = HttpContext.Session.GetString(SessionID);
            ViewBag.Session_fullname = fullname;
            ViewBag.SessionID = Session;

            if(Session == "" || fullname == ""){
                return RedirectToAction("Index", "Home");
            }
            else
            {
                return View();
            }

           
        }

        private System.Data.DataTable datatb;

        [HttpPost]
        public IActionResult GetfileExcel(IFormFile File)
        {
            DataTable dt = new DataTable();
          
            if (File != null)
            {
              
                using var workbook = new XLWorkbook(File.OpenReadStream());
                var ws = workbook.Worksheet(1);

                foreach (IXLRow row in ws.Rows())
                {


                    int i = 0;
                  if(row.RowNumber() >= 4)
                    {
                        dt.Rows.Add();
                    }
                   
                    
                    
                    foreach (IXLCell cell in row.Cells())
                    {
                       
                        if (Convert.ToInt32(cell.Address.RowNumber) == 2)
                        {
                            
                            dt.Columns.Add(cell.Value.ToString());
                            
                        }
                        else if (Convert.ToInt32(cell.Address.RowNumber) >= 4)
                        {
                           
                           

                            if (i == 10)
                            {
                                
                                dt.Rows[dt.Rows.Count - 1][i] = cell.GetDateTime().ToString("d-MMM-yy");
                            
                            }
                            else if ( i < 10)
                            {
                            
                                dt.Rows[dt.Rows.Count - 1][i] = cell.Value.ToString();
                                
                            }


                           
                            i++;

                        }
                     

                    }


                    if (row.IsEmpty() && row.RowNumber() >= 4)
                    {
                        dt.Rows[dt.Rows.Count - 1].Delete();
                        break;
                    }

                }

                
            }
            DataTable DTSqlBulk;
            datatb = new DataTable();
            datatb = dt;
            DTSqlBulk = new DataTable();
            DTSqlBulk = dt;
            System.Data.DataColumn newColumn = new System.Data.DataColumn("ResignMaking", typeof(System.String));
            string month = DateTime.Now.AddMonths(-1).ToString("MMMM");
            string year = DateTime.Now.ToString("yyyy");
            string Shortmonth = DateTime.Now.AddMonths(-1).ToString("MM");
            string month_year = Shortmonth + "/" + year;
            newColumn.DefaultValue = month_year;

            DTSqlBulk.Columns.Add(newColumn);


            if (DTSqlBulk.Rows.Count > 0)
            {

                string consString = "Data Source=RTHSRV14;Initial Catalog=ImportExportDB;Persist Security Info=True;User ID=sa;Password=pwpolicy;";
                using (SqlConnection con = new SqlConnection(consString))
                {
                    using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(con))
                    {
                        //Set the database table name
                        sqlBulkCopy.DestinationTableName = "dbo.OperatorsResign";

                        sqlBulkCopy.ColumnMappings.Add("NO.", "NO");
                        sqlBulkCopy.ColumnMappings.Add("CODE", "OPID");
                        sqlBulkCopy.ColumnMappings.Add("NAME", "OPName");
                        sqlBulkCopy.ColumnMappings.Add("Column1", "OPSurName");
                        sqlBulkCopy.ColumnMappings.Add("POSITION", "OPPosition");
                        sqlBulkCopy.ColumnMappings.Add("LEVEL", "OPLevel");
                        sqlBulkCopy.ColumnMappings.Add("SECT.", "OPSect");
                        sqlBulkCopy.ColumnMappings.Add("DEPT.", "OPDept");
                        sqlBulkCopy.ColumnMappings.Add("DIV.", "OPDiv");
                        sqlBulkCopy.ColumnMappings.Add("HQ.", "OPHq");
                        sqlBulkCopy.ColumnMappings.Add("RESIGNED DATE", "ResignDate");
                        sqlBulkCopy.ColumnMappings.Add("ResignMaking", "ResignDateMaking");

                        con.Open();
                        sqlBulkCopy.WriteToServer(DTSqlBulk);
                        con.Close();
                    }
                }
            }
            //var insertLog = InstoServer_fromDT(datatb);
            DTSqlBulk.Columns.Remove(newColumn);
            //pathFile = Export_To_Excel(datatb,"Excel","");
            var pathFile = Export_To_Excel(datatb, "");

            //pathFile = "C://Users/010724/Desktop/Resignation_ 202207.xlsx";
            return Json(pathFile);
        }

        public JsonResult GetExcel_Ddl_resignDate(string resignDate)
        {
            DataTable dt = new DataTable();
            var mgrSQLConnect = new mgrSQLConnect(configuration);
            string query = "select NO, OPID, OPName, OPSurName, OPPosition, OPLevel, OPSect, OPDept, OPDiv, OPHq, ResignDate from[ImportExportDB].[dbo].[OperatorsResign]  where resignDateMaking  = '" + resignDate + "'";
            dt = mgrSQLConnect.GetDatatables(query);
            //var pathfile =  Export_To_Excel(dt, "Excel_DdlResignDate", resignDate);
            var pathfile = Export_To_Excel(dt, resignDate);
            return Json(pathfile);
        }



        public JsonResult Export_To_Excel(DataTable datatb , string DateRemake)
        {
            int lastRow = 0;
            string contentType , month, year , Shortmonth , fileName;
            if (DateRemake == "")
            {

                 contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                 month = DateTime.Now.AddMonths(-1).ToString("MMMM");
                 year = DateTime.Now.ToString("yyyy");
                 Shortmonth = DateTime.Now.AddMonths(-1).ToString("MM");
                 fileName = "Resignation_ " + year + Shortmonth + ".xlsx";
            }
            else
            {
                contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                month = Convert.ToDateTime(DateRemake).ToString("MMMM");
                year = Convert.ToDateTime(DateRemake).ToString("yyyy"); 
                Shortmonth = Convert.ToDateTime(DateRemake).ToString("MM");
                fileName = "Resignation_ " + year + Shortmonth + ".xlsx";
            }
            //string imagePathDomain = Server.MapPath(@"Content\img\DomainHeadPic.png"), imagePathSystem = Server.MapPath(@"Content\img\SystemHeadPic.png");
            string imagePathDomain = System.IO.Directory.GetCurrentDirectory() + "\\wwwroot\\Content\\img\\DomainHeadPic.png", imagePathSystem = System.IO.Directory.GetCurrentDirectory() + "\\wwwroot\\Content\\img\\SystemHeadPic.png";

            try
            {
                using (var Workbook = new XLWorkbook())
                {
                    IXLWorksheet worksheet =
                    Workbook.Worksheets.Add("Original");
                    //--------------------header sheet original----------------------------------
                    worksheet.Cell(1, 1).Value = "No.";
                    worksheet.Cell(1, 2).Value = "CODE";
                    worksheet.Cell(1, 3).Value = "NAME";
                    worksheet.Cell(1, 4).Value = "Sname";
                    worksheet.Cell(1, 4).Style.Alignment.SetTextRotation(90);
                    worksheet.Cell(1, 5).Value = "POSITION";
                    worksheet.Cell(1, 6).Value = "LEVEL";
                    worksheet.Cell(1, 6).Style.Alignment.SetTextRotation(90);
                    worksheet.Cell(1, 7).Value = "SECT.";
                    worksheet.Cell(1, 8).Value = "DEPT.";
                    worksheet.Cell(1, 9).Value = "DIV.";
                    worksheet.Cell(1, 10).Value = "HQ.";
                    worksheet.Cell(1, 11).Value = "RESIGNED DATE";
                    worksheet.Range("A1:K1").Style.Fill.BackgroundColor = XLColor.FromArgb(204, 255, 204);
                    worksheet.Range("A1:K1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    worksheet.Range("A1:K1").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    worksheet.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    worksheet.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                    worksheet.Column("A").Width = 2.86;
                    worksheet.Column("B").Width = 5.43;
                    worksheet.Column("B").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                    worksheet.Column("B").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                    worksheet.Column("C").Width = 18.43;
                    worksheet.Column("C").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                    worksheet.Column("C").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                    worksheet.Column("D").Width = 2.29;
                    worksheet.Column("E").Width = 13.71;
                    worksheet.Column("F").Width = 2.29;
                    worksheet.Column("G").Width = 38.29;
                    worksheet.Column("H").Width = 17.71;
                    worksheet.Column("I").Width = 29.86;
                    worksheet.Column("J").Width = 10;
                    worksheet.Column("K").Width = 11.71;

                    worksheet.Row(1).Height = 30;

                    worksheet.Style.Font.FontSize = 8;
                    worksheet.Style.Font.FontName = "Arial";
                    worksheet.Range("A1:K1").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    worksheet.Range("A1:K1").Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                    worksheet.Column(2).Style.NumberFormat.Format = "000000";
                    worksheet.SheetView.FreezeRows(1);

                    worksheet.PageSetup.Margins.Top = 0.393;
                    worksheet.PageSetup.Margins.Bottom = 0.984;
                    worksheet.PageSetup.Margins.Left = 0.157;
                    worksheet.PageSetup.Margins.Right = 0.157;
                    worksheet.PageSetup.Margins.Footer = 0.511;
                    worksheet.PageSetup.Margins.Header = 0.511;
                    worksheet.PageSetup.CenterHorizontally = true;
                    worksheet.PageSetup.AdjustTo(50);
                    worksheet.PageSetup.PageOrientation = XLPageOrientation.Landscape;

                    //--------------------header sheet System----------------------------------
                    IXLWorksheet worksheet2 =
                        Workbook.Worksheets.Add("System");

                    worksheet2.Cell(9, 1).Value = "No.";
                    worksheet2.Cell(9, 2).Value = "DIV.";
                    worksheet2.Cell(9, 3).Value = "CODE";
                    worksheet2.Range("D9:E9").Merge().Value = "NAME";
                    worksheet2.Cell(9, 5).Value = "";
                    worksheet2.Cell(9, 6).Value = "POSITION";
                    worksheet2.Style.Font.FontSize = 14;
                    //----------IXLColumn TRDI ------------
                    //------------Header------------------------------------    
                    worksheet2.RowHeight = 25.50;
                    worksheet2.Row(2).Height = 25.50; worksheet2.Row(3).Height = 25.50; worksheet2.Row(4).Height = 25.50; worksheet2.Row(5).Height = 25.50;
                    worksheet2.Row(6).Height = 46.50;
                    worksheet2.Row(7).Height = 25.50; worksheet2.Row(8).Height = 25.50; worksheet2.Row(9).Height = 25.50;

                    worksheet2.RowHeight = 25.50;

                    worksheet2.Range("A2:I5").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    worksheet2.Range("A7:AH9").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    worksheet2.Range("A2:I5").Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                    worksheet2.Range("A7:AH9").Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                    worksheet2.Range("A2:I5").Style.Font.FontName = "Tahoma";
                    worksheet2.Range("A2:I5").Style.Font.FontSize = 20;
                    worksheet2.Range("A2:I5").Merge().Value = "Resigned Person of " + month + " " + year + Environment.NewLine + "(System User)";
                    worksheet2.Range("A2:I5").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    worksheet2.Range("A2:I5").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                    worksheet2.Range("A2:I5").Style.Font.Bold = true;
                    worksheet2.Range("A6:F6").Merge();
                    worksheet2.Range("A7:F8").Merge().Value = "Resigned Person";
                    worksheet2.Range("A7:AH9").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    worksheet2.Range("A7:AH9").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    worksheet2.Style.Font.FontName = "Arial";

                    worksheet2.Range("A7:F7").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    worksheet2.Range("A7:F7").Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                    worksheet2.Column("A").Width = 6; worksheet2.Column("B").Width = 23.86; worksheet2.Column("C").Width = 12.86; worksheet2.Column("D").Width = 32.57; worksheet2.Column("E").Width = 4; worksheet2.Column("F").Width = 35.46;
                    worksheet2.Column("G").Width = 16; worksheet2.Column("H").Width = 14.75; worksheet2.Column("I").Width = 11.14; worksheet2.Column("J").Width = 15.43; worksheet2.Column("K").Width = 12.43; worksheet2.Column("L").Width = 18.86;
                    worksheet2.Column("M").Width = 11.14; worksheet2.Column("N").Width = 7.43; worksheet2.Column("O").Width = 11.14; worksheet2.Column("P").Width = 11.43; worksheet2.Column("Q").Width = 18.86; worksheet2.Column("R").Width = 15.43;
                    worksheet2.Column("S").Width = 11.14; worksheet2.Column("T").Width = 9.57; worksheet2.Column("U").Width = 10.29; worksheet2.Column("V").Width = 10.29; worksheet2.Column("W").Width = 10; worksheet2.Column("X").Width = 9.57;
                    worksheet2.Column("Y").Width = 11.14; worksheet2.Column("Z").Width = 9.57; worksheet2.Column("AA").Width = 11.14; worksheet2.Column("AB").Width = 10.71; worksheet2.Column("AC").Width = 11.71; worksheet2.Column("AD").Width = 15.43;
                    worksheet2.Column("AE").Width = 10.71; worksheet2.Column("AF").Width = 7.86; worksheet2.Column("AG").Width = 7.86; worksheet2.Column("AH").Width = 8.29;
                    worksheet2.RowHeight = 25.50;



                    worksheet2.Range("A7:F9").Style.Font.Bold = true;

                    //worksheet2.Range("R1:AG5").Style.Font.Bold = true;
                    //worksheet2.Range("R1:AG5").Style.Font.FontSize = 16;
                    //worksheet2.Range("R1:AG5").Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                    //worksheet2.Range("R1:AG5").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    //worksheet2.Range("R1:AG5").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    //worksheet2.Range("R1:AG5").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                    //worksheet2.Range("R1:AG5").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    worksheet2.Style.Font.FontName = "Tahoma";
                    worksheet2.Range("R8:AH9").Style.Font.Bold = false;

                    //worksheet2.Range("R1:S1").Merge().Value = "TRDI";
                    //worksheet2.Range("R2:S4").Merge();
                    //worksheet2.Range("R6:S6").Merge();


                    //worksheet2.Range("T1:U1").Merge().Value = "TC";
                    //worksheet2.Range("T2:U4").Merge();
                    //worksheet2.Range("T6:U6").Merge();

                    //worksheet2.Range("V1:W1").Merge().Value = "MCR";
                    //worksheet2.Range("V2:W4").Merge();
                    //worksheet2.Range("V6:W6").Merge();

                    //worksheet2.Range("X1:Y1").Merge().Value = "LSI";
                    //worksheet2.Range("X2:Y4").Merge();
                    //worksheet2.Range("X6:Y6").Merge();

                    //worksheet2.Range("Z1:AA1").Merge().Value = "ADMIN";
                    //worksheet2.Range("Z2:AA4").Merge();
                    //worksheet2.Range("Z6:AA6").Merge();

                    //worksheet2.Range("AB1:AC1").Merge().Value = "IS Sect.Mgr.";
                    //worksheet2.Range("AB2:AC4").Merge();
                    //worksheet2.Range("AB6:AC6").Merge();

                    //worksheet2.Range("AD1:AE1").Merge().Value = "IS Dept.Mgr.";
                    //worksheet2.Range("AD2:AE4").Merge();
                    //worksheet2.Range("AD6:AE6").Merge();

                    //worksheet2.Range("AF1:AG1").Merge().Value = "IS Div.Mgr.";
                    //worksheet2.Range("AF2:AG4").Merge();
                    //worksheet2.Range("AF6:AG6").Merge();

                    //------------Header------------------------------------
                    worksheet2.Range("G7:N7").Merge().Value = "TRDI";
                    worksheet2.Range("G7:N7").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    worksheet2.Range("G7:N7").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    worksheet2.Range("G7:N7").Style.Font.Bold = true;
                    worksheet2.Range("G7:N9").Style.Fill.BackgroundColor = XLColor.FromArgb(204, 255, 204);


                    worksheet2.Cell(8, 7).Value = "OEM";
                    worksheet2.Cell(8, 8).Value = "Stagnation";
                    worksheet2.Cell(8, 9).Value = "LotTraceability";
                    worksheet2.Cell(8, 10).Value = "Alarm";
                    worksheet2.Cell(8, 11).Value = "Non-conformance";
                    worksheet2.Cell(8, 12).Value = "SpareParts";
                    worksheet2.Cell(8, 13).Value = "Plating";
                    worksheet2.Cell(8, 14).Value = "NCIM";

                    worksheet2.Cell(9, 7).Value = "OEMOperator";
                    worksheet2.Cell(9, 8).Value = "Operator";
                    worksheet2.Cell(9, 9).Value = "Operator";
                    worksheet2.Cell(9, 10).Value = "T_Operator";
                    worksheet2.Cell(9, 11).Value = "Operator";
                    worksheet2.Cell(9, 12).Value = "Operator";
                    worksheet2.Cell(9, 13).Value = "OP";
                    worksheet2.Cell(9, 14).Value = "Operator";

                    //----------IXLColumn TC ------------
                    worksheet2.Range("O7:P7").Merge().Value = "TC";
                    worksheet2.Range("O7:P7").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    worksheet2.Range("O7:P7").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    worksheet2.Range("O7:P7").Style.Font.Bold = true;
                    worksheet2.Range("O7:P9").Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 153);

                    worksheet2.Cell(8, 15).Value = "TC System";
                    worksheet2.Cell(8, 16).Value = "Non-conformance";
                    worksheet2.Cell(9, 15).Value = "Operator";
                    worksheet2.Cell(9, 16).Value = "Operator";

                    //----------IXLColumn MCR ------------
                    worksheet2.Range("Q7:Y7").Merge().Value = "MCR";
                    worksheet2.Range("Q7:Y7").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    worksheet2.Range("Q7:Y7").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    worksheet2.Range("Q7:Y7").Style.Font.Bold = true;
                    worksheet2.Range("Q7:Y9").Style.Fill.BackgroundColor = XLColor.FromArgb(204, 255, 255);
                    worksheet2.Range("Q8:R9").Style.Fill.BackgroundColor = XLColor.FromArgb(216, 228, 188);


                    worksheet2.Cell(8, 17).Value = "LotTraceability";
                    worksheet2.Cell(8, 18).Value = "SpareParts";
                    worksheet2.Cell(8, 19).Value = "PI";
                    worksheet2.Cell(8, 20).Value = "B-Kanban";
                    worksheet2.Cell(8, 21).Value = "Shipment";
                    worksheet2.Cell(8, 22).Value = "Substrate";
                    worksheet2.Cell(8, 23).Value = "Paste";
                    worksheet2.Cell(8, 24).Value = "Screen";
                    worksheet2.Cell(8, 25).Value = "NCIM";


                    worksheet2.Cell(9, 17).Value = "Operator";
                    worksheet2.Cell(9, 18).Value = "Operator";
                    worksheet2.Cell(9, 19).Value = "Operator";
                    worksheet2.Cell(9, 20).Value = "Operator";
                    worksheet2.Cell(9, 21).Value = "Operator";
                    worksheet2.Cell(9, 22).Value = "Operator";
                    worksheet2.Cell(9, 23).Value = "Operator";
                    worksheet2.Cell(9, 24).Value = "Operator";
                    worksheet2.Cell(9, 25).Value = "Operator";
                    //----------IXLColumn OPM ------------

                    worksheet2.Range("Z7:AA7").Merge().Value = "OPM&WLM";
                    worksheet2.Range("Z7:AA7").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    worksheet2.Range("Z7:AA7").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    worksheet2.Range("Z7:AA7").Style.Font.Bold = true;
                    worksheet2.Range("Z7:AA9").Style.Fill.BackgroundColor = XLColor.FromArgb(204, 192, 218);

                    worksheet2.Cell(8, 26).Value = "spareParts";
                    worksheet2.Cell(8, 27).Value = "NCIM";

                    worksheet2.Cell(9, 26).Value = "Operator";
                    worksheet2.Cell(9, 27).Value = "Operator";

                    //----------IXLColumn LSI ------------
                    worksheet2.Range("AB7:AC7").Merge().Value = "LSI";
                    worksheet2.Range("AB7:AC7").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    worksheet2.Range("AB7:AC7").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    worksheet2.Range("AB7:AC7").Style.Font.Bold = true;
                    worksheet2.Range("AB7:AC9").Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 153);

                    worksheet2.Cell(8, 28).Value = "Ukebarai : FT";
                    worksheet2.Cell(8, 29).Value = "Web : FYI";

                    worksheet2.Cell(9, 28).Value = "OP";
                    worksheet2.Cell(9, 29).Value = "OP";


                    //worksheet2.AddPicture(imagePathDomain)
                    //  .MoveTo(worksheet2.Cell("B3"))
                    //  .Scale(0.5);
                    //----------IXLColumn Admin ------------
                    worksheet2.Range("AD7:AH7").Merge().Value = "Admin";
                    worksheet2.Range("AD7:AH7").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    worksheet2.Range("AD7:AH7").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    worksheet2.Range("AD7:AH7").Style.Font.Bold = true;
                    worksheet2.Range("AD7:AH9").Style.Fill.BackgroundColor = XLColor.FromArgb(255, 153, 204);

                    worksheet2.Cell(8, 30).Value = "GFDReport";
                    worksheet2.Cell(8, 31).Value = "MaterialLedger";
                    worksheet2.Cell(8, 32).Value = "OneWorld";
                    worksheet2.Cell(8, 33).Value = "TCCost";
                    worksheet2.Cell(8, 34).Value = "TRCost";

                    worksheet2.Cell(9, 30).Value = "Operator";
                    worksheet2.Cell(9, 31).Value = "Operator";
                    worksheet2.Cell(9, 32).Value = "Operator";
                    worksheet2.Cell(9, 33).Value = "OP";
                    worksheet2.Cell(9, 34).Value = "OP";
                    //worksheet2.AddPicture(imagePathSystem).MoveTo(worksheet2.Cell("R1").Address).Scale(1.4);
                    IXLPicture iXLPicture = worksheet2.AddPicture(imagePathSystem).MoveTo(worksheet2.Cell("R1")).Scale(1.4);

                    worksheet2.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    worksheet2.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                    worksheet2.Column("F").Style.Protection.SetLocked(true);
                    worksheet2.Column(3).Style.NumberFormat.Format = "000000";

                    worksheet2.SheetView.FreezeRows(9);
                    worksheet2.SheetView.FreezeColumns(6);
                    worksheet2.SheetView.ZoomScale = 75;
                    var Range = worksheet2.Range("F7");
                    Range.SetAutoFilter();

                    worksheet2.PageSetup.PageOrientation = XLPageOrientation.Landscape;
                    worksheet2.PageSetup.FitToPages(1, 1);

                    IXLWorksheet worksheet3 =
                         Workbook.Worksheets.Add("Domain");
                    //---------------header-------------
                    worksheet3.Range("A3:C6").Merge().Value = "Resigned Person of " + month + " " + year + Environment.NewLine + "(System User)";
                    worksheet3.Range("A3:C6").Style.Font.FontName = "Tahoma";
                    worksheet3.Range("A3:C6").Style.Font.FontSize = 12;
                    worksheet3.Style.Font.FontSize = 10;
                    worksheet3.Style.Font.FontName = "Arial";
                    worksheet3.Range("A3:C6").Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                    worksheet3.Range("A3:C6").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    worksheet3.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    worksheet3.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    worksheet3.Range("A3:C6").Style.Font.FontName = "Tahoma";

                    worksheet3.Column("A").Width = 7.71;
                    worksheet3.Column("B").Width = 20.29;
                    worksheet3.Column("C").Width = 12.57;
                    worksheet3.Column("D").Width = 21;
                    worksheet3.Column("E").Width = 3.29;
                    worksheet3.Column("F").Width = 23.57;
                    worksheet3.Column("G").Width = 11.71;

                    worksheet3.RowHeight = 12.75;

                    worksheet3.Range("A3:C6").Merge().Value = "Resigned Person of " + month + " " + year + Environment.NewLine + "( Domain User )";
                    //worksheet3.Range("D2").Value = "In Charge"; worksheet3.Range("E2").Value = "IS Sect.Mgr."; worksheet3.Range("F2").Value = "IS Dept.Mgr."; worksheet3.Range("G2").Value = "Admin Div.Mgr.";
                    //worksheet3.Range("D2:G7").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    //worksheet3.Range("D2:G7").Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                    //worksheet3.Range("D2:G7").Style.Font.Bold = true;
                    //worksheet3.Range("D3:D7").Merge();
                    //worksheet3.Range("E3:E7").Merge();
                    //worksheet3.Range("F3:F7").Merge();
                    //worksheet3.Range("G3:G7").Merge();


                    worksheet3.Cell(9, 1).Value = "NO.";
                    worksheet3.Cell(9, 2).Value = "DIV.";
                    worksheet3.Cell(9, 3).Value = "CODE";
                    worksheet3.Cell(9, 4).Value = "NAME";
                    worksheet3.Cell(9, 5).Value = "";
                    worksheet3.Cell(9, 6).Value = "POSITION";
                    worksheet3.Cell(9, 7).Value = "Domain User";
                    worksheet3.Range("A9:G9").Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                    worksheet3.Range("A9:G9").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                    worksheet3.Range("A3:C6").Style.Font.FontName = "Arial";
                    worksheet3.Range("A3:C6").Style.Font.FontSize = 12;
                    worksheet3.Range("A9:G9").Style.Font.FontName = "Tahoma";
                    worksheet3.Range("A9:G9").Style.Font.FontSize = 10;
                    worksheet3.Range("A9:G9").Style.Font.Bold = true;
                    worksheet3.Range("A3:C6").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    worksheet3.Range("A3:C6").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    worksheet3.Range("A3:C6").Style.Font.Bold = true;


                    worksheet3.Range("A9:G9").Style.Fill.BackgroundColor = XLColor.LightGray;
                    worksheet3.Range("D2:G7").Merge();
                    worksheet3.SheetView.FreezeRows(9);
                    //worksheet3.AddPicture(imagePathDomain).MoveTo(worksheet3.Cell("D2").Address).Scale(1.1);
                    IXLPicture iXLPicture3 = worksheet3.AddPicture(imagePathDomain).MoveTo(worksheet3.Cell("D2"), 50, 0).Scale(1.1);

                    worksheet3.Column(3).Style.NumberFormat.Format = "000000";

                    IXLWorksheet worksheet4 =
                       Workbook.Worksheets.Add("Mail");
                    //---------------header-------------
                    worksheet4.Column("A").Width = 3.71;
                    worksheet4.Column("B").Width = 20.71;
                    worksheet4.Column("C").Width = 12.57;
                    worksheet4.Column("D").Width = 21;
                    worksheet4.Column("E").Width = 3.29;
                    worksheet4.Column("F").Width = 23.57;
                    worksheet4.Column("G").Width = 11.86;
                    worksheet4.Column("H").Width = 12.43;
                    worksheet4.RowHeight = 12.75;


                    worksheet4.Range("A3:C6").Merge().Value = "Resigned Person of " + month + " " + year + Environment.NewLine + "( Mail (Internet & Internal) )";
                    worksheet4.Range("A3:C6").Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                    worksheet4.Range("A3:C6").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                    worksheet4.Range("A3:C6").Style.Font.Bold = true;
                    worksheet4.Range("A3:C6").Style.Font.FontSize = 12;
                    worksheet4.Range("A3:C6").Style.Font.FontName = "Tahoma";

                    //worksheet4.Range("D2").Value = "In Charge"; worksheet4.Range("E2").Value = "IS Sect.Mgr."; worksheet4.Range("F2").Value = "IS Dept.Mgr."; worksheet4.Range("G2").Value = "Admin Div.Mgr.";
                    //worksheet4.Range("D3:D7").Merge();
                    //worksheet4.Range("E3:E7").Merge();
                    //worksheet4.Range("F3:F7").Merge();
                    //worksheet4.Range("G3:G7").Merge();
                    //worksheet4.Range("G3:G7").Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                    //worksheet4.Range("E3:E7").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    //worksheet4.Range("E3:E7").Style.Font.Bold = true;

                    worksheet4.Cell(9, 1).Value = "NO.";
                    worksheet4.Cell(9, 2).Value = "DIV.";
                    worksheet4.Cell(9, 3).Value = "CODE";
                    worksheet4.Cell(9, 4).Value = "NAME";
                    worksheet4.Cell(9, 5).Value = "";
                    worksheet4.Cell(9, 6).Value = "POSITION";
                    worksheet4.Cell(9, 7).Value = "Internal Mail";
                    worksheet4.Cell(9, 8).Value = "Internet Mail";
                    worksheet4.Range("A9:H9").Style.Fill.BackgroundColor = XLColor.LightGray;
                    worksheet4.Range("A9:H9").Style.Font.Bold = true;
                    worksheet4.Range("A9:H9").Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                    worksheet4.Range("A9:H9").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    worksheet4.Range("A9:H9").Style.Font.FontSize = 10;
                    worksheet4.Range("A9:H9").Style.Font.FontName = "Tahoma";


                    worksheet4.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    worksheet4.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                    worksheet4.Column(3).Style.NumberFormat.Format = "000000";
                    worksheet4.SheetView.FreezeRows(9);
                    IXLPicture iXLPicture4 = worksheet4.AddPicture(imagePathDomain).MoveTo(worksheet4.Cell("E2")).Scale(1.05);
                    
                    for (int nrow = 0; nrow < datatb.Rows.Count; nrow++)
                    {
                        for (int ncol = 1; ncol < datatb.Columns.Count + 1; ncol++)
                        {
                                if (ncol == 2)
                                {

                                var OPID = datatb.Rows[nrow][ncol - 1].ToString();
                                string status_TROEM = Check_Operator("TROEM", OPID);
                                if (status_TROEM == "true")
                                {
                                    worksheet2.Cell(nrow + 10, ncol + 5).Value = "O";
                                    worksheet2.Cell(nrow + 10, ncol + 5).Style.Font.Bold = true;
                                }

                                string statusAlarmDB = Check_Operator("AlarmDB", OPID);
                                if (statusAlarmDB == "true")
                                {
                                    worksheet2.Cell(nrow + 10, ncol + 8).Value = "O";
                                    worksheet2.Cell(nrow + 10, ncol + 8).Style.Font.Bold = true;
                                }

                                string statusNonDb = Check_Operator("NonDb", OPID);
                                if (statusNonDb == "true")
                                {
                                    worksheet2.Cell(nrow + 10, ncol + 9).Value = "O";
                                    worksheet2.Cell(nrow + 10, ncol + 9).Style.Font.Bold = true;
                                }
                                string statusPlatingDB = Check_Operator("PlatingDB", OPID);
                                if (statusPlatingDB == "true")
                                {
                                    worksheet2.Cell(nrow + 10, ncol + 11).Value = "O";
                                    worksheet2.Cell(nrow + 10, ncol + 11).Style.Font.Bold = true;
                                }

                                string statusTRProcessControl = Check_Operator("TRProcessControl", OPID);
                                if (statusTRProcessControl == "true")
                                {
                                    worksheet2.Cell(nrow + 10, ncol + 12).Value = "O";
                                    worksheet2.Cell(nrow + 10, ncol + 12).Style.Font.Bold = true;
                                }

                                string statusTcOPOLLO = Check_Operator("TcOPOLLO", OPID);
                                if (statusTcOPOLLO == "true")
                                {
                                    worksheet2.Cell(nrow + 10, ncol + 13).Value = "O";
                                    worksheet2.Cell(nrow + 10, ncol + 13).Style.Font.Bold = true;
                                }
                                string statusTCNonDB = Check_Operator("TCNonDB", OPID);
                                if (statusTCNonDB == "true")
                                {
                                    worksheet2.Cell(nrow + 10, ncol + 14).Value = "O";
                                    worksheet2.Cell(nrow + 10, ncol + 14).Style.Font.Bold = true;
                                }
                                string statusMCRLT = Check_Operator("MCRLT", OPID);
                                if (statusMCRLT == "true")
                                {
                                    worksheet2.Cell(nrow + 10, ncol + 15).Value = "O";
                                    worksheet2.Cell(nrow + 10, ncol + 15).Style.Font.Bold = true;
                                }



                                string statusMCRSpareParts = Check_Operator("MCRSpareParts", OPID);
                                if (statusMCRSpareParts == "true")
                                {
                                    worksheet2.Cell(nrow + 10, ncol + 16).Value = "O";
                                    worksheet2.Cell(nrow + 10, ncol + 16).Style.Font.Bold = true;
                                }

                                string statusMCR = Check_Operator("MCR", OPID);
                                if (statusMCR == "true")
                                {
                                    worksheet2.Cell(nrow + 10, ncol + 17).Value = "O";
                                    worksheet2.Cell(nrow + 10, ncol + 17).Style.Font.Bold = true;
                                }

                                //string statusMCRTPProcessControl = Check_Operator("MCRTPProcessControl", OPID);
                                //if (statusMCRTPProcessControl == "true")
                                //{
                                //    worksheet2.Cell(nrow + 8, ncol + 18).Value = "O";
                                //    worksheet2.Cell(nrow + 8, ncol + 18).Style.Font.Bold = true;
                                //}

                                string statusMCRMaterialControl = Check_Operator("MCRMaterialControl", OPID);
                                if (statusMCRMaterialControl == "true")
                                {
                                    worksheet2.Cell(nrow + 10, ncol + 20).Value = "O";
                                    worksheet2.Cell(nrow + 10, ncol + 20).Style.Font.Bold = true;
                                                           
                                    worksheet2.Cell(nrow + 10, ncol + 21).Value = "O";
                                    worksheet2.Cell(nrow + 10, ncol + 21).Style.Font.Bold = true;
                                                           
                                    worksheet2.Cell(nrow + 10, ncol + 22).Value = "O";
                                    worksheet2.Cell(nrow + 10, ncol + 22).Style.Font.Bold = true;
                                }


                                string statusMCRProcessControl = Check_Operator("MCRProcessControl", OPID);
                                if (statusMCRProcessControl == "true")
                                {
                                    worksheet2.Cell(nrow + 10, ncol + 23).Value = "O";
                                    worksheet2.Cell(nrow + 10, ncol + 23).Style.Font.Bold = true;
                                }


                                string statusOPMSparepart = Check_Operator("OPMSparepart", OPID);
                                if (statusOPMSparepart == "true")
                                {
                                    worksheet2.Cell(nrow + 10, ncol + 24).Value = "O";
                                    worksheet2.Cell(nrow + 10, ncol + 24).Style.Font.Bold = true;
                                }

                                string statusOPMProcessControl = Check_Operator("OPMProcessControl", OPID);
                                if (statusOPMProcessControl == "true")
                                {
                                    worksheet2.Cell(nrow + 10, ncol + 25).Value = "O";
                                    worksheet2.Cell(nrow + 10, ncol + 25).Style.Font.Bold = true;
                                }


                                string statusGFDReport = Check_Operator("GFDReport", OPID);
                                if (statusGFDReport == "true")
                                {
                                    worksheet2.Cell(nrow + 10, ncol + 28).Value = "O";
                                    worksheet2.Cell(nrow + 10, ncol + 28).Style.Font.Bold = true;
                                }


                                string statusOneWord = Check_Operator("OneWord", OPID);
                                if (statusOneWord == "true")
                                {
                                    worksheet2.Cell(nrow + 10, ncol + 30).Value = "O";
                                    worksheet2.Cell(nrow + 10, ncol + 30).Style.Font.Bold = true;
                                }

                                string statusTCCostDb = Check_Operator("TCCostDb", OPID);
                                if (statusTCCostDb == "true")
                                {
                                    worksheet2.Cell(nrow + 10, ncol + 31).Value = "O";
                                    worksheet2.Cell(nrow + 10, ncol + 31).Style.Font.Bold = true;
                                }

                                string statusTRCostDb = Check_Operator("TRCostDb", OPID);
                                if (statusOPMProcessControl == "true")
                                {
                                    worksheet2.Cell(nrow + 10, ncol + 32).Value = "O";
                                    worksheet2.Cell(nrow + 10, ncol + 32).Style.Font.Bold = true;
                                }




                            }
                            if (ncol == 1)
                            {
                                 worksheet.Cell(nrow + 2, ncol).Value = nrow + 1;
                            }
                            else if (ncol == 11)
                            {
                                worksheet.Cell(nrow + 2, ncol).Value = Convert.ToDateTime(datatb.Rows[nrow][ncol - 1]).ToString("d-MMM-yy");
                                worksheet.Cell(nrow + 2, ncol).Style.DateFormat.Format = "d-MMM-yy";
                            }
                            else
                            {
                                worksheet.Cell(nrow + 2, ncol).Value = datatb.Rows[nrow][ncol - 1].ToString();
                            }

                            if (datatb.Rows[nrow][1].ToString() != "")
                            {
                                worksheet.Cell(nrow + 2, ncol).Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                                worksheet.Cell(nrow + 2, ncol).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                                worksheet.Row(nrow + 2).Height = 12.75;
                            }


                        }
                        worksheet2.Cell(nrow + 10, 1).Value = nrow + 1;
                        worksheet2.Cell(nrow + 10, 2).Value = datatb.Rows[nrow][8].ToString();
                        worksheet2.Cell(nrow + 10, 3).Value = datatb.Rows[nrow][1].ToString();
                        worksheet2.Cell(nrow + 10, 4).Value = datatb.Rows[nrow][2].ToString();
                        worksheet2.Cell(nrow + 10, 5).Value = datatb.Rows[nrow][3].ToString();
                        worksheet2.Cell(nrow + 10, 6).Value = datatb.Rows[nrow][4].ToString();

                        worksheet3.Cell(nrow + 10, 1).Value = nrow + 1;
                        worksheet3.Cell(nrow + 10, 2).Value = datatb.Rows[nrow][8].ToString();
                        worksheet3.Cell(nrow + 10, 3).Value = datatb.Rows[nrow][1].ToString();
                        worksheet3.Cell(nrow + 10, 4).Value = datatb.Rows[nrow][2].ToString();
                        worksheet3.Cell(nrow + 10, 5).Value = datatb.Rows[nrow][3].ToString();
                        worksheet3.Cell(nrow + 10, 6).Value = datatb.Rows[nrow][4].ToString();


                        worksheet4.Cell(nrow + 10, 1).Value = nrow + 1;
                        worksheet4.Cell(nrow + 10, 2).Value = datatb.Rows[nrow][8].ToString();
                        worksheet4.Cell(nrow + 10, 3).Value = datatb.Rows[nrow][1].ToString();
                        worksheet4.Cell(nrow + 10, 4).Value = datatb.Rows[nrow][2].ToString();
                        worksheet4.Cell(nrow + 10, 5).Value = datatb.Rows[nrow][3].ToString();
                        worksheet4.Cell(nrow + 10, 6).Value = datatb.Rows[nrow][4].ToString();
                        if (datatb.Rows[nrow][4].ToString() != "")
                        {
                            worksheet2.Cell(nrow + 10, 1).Style.Border.InsideBorder = XLBorderStyleValues.Thin; worksheet2.Cell(nrow + 10, 1).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                            worksheet2.Cell(nrow + 10, 2).Style.Border.InsideBorder = XLBorderStyleValues.Thin; worksheet2.Cell(nrow + 10, 2).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                            worksheet2.Cell(nrow + 10, 3).Style.Border.InsideBorder = XLBorderStyleValues.Thin; worksheet2.Cell(nrow + 10, 3).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                            worksheet2.Cell(nrow + 10, 4).Style.Border.InsideBorder = XLBorderStyleValues.Thin; worksheet2.Cell(nrow + 10, 4).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                            worksheet2.Cell(nrow + 10, 5).Style.Border.InsideBorder = XLBorderStyleValues.Thin; worksheet2.Cell(nrow + 10, 5).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                            worksheet2.Cell(nrow + 10, 6).Style.Border.InsideBorder = XLBorderStyleValues.Thin; worksheet2.Cell(nrow + 10, 6).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                                                                                                                                       
                                                                                                                                       
                            worksheet3.Cell(nrow + 10, 1).Style.Border.InsideBorder = XLBorderStyleValues.Thin; worksheet3.Cell(nrow + 10, 1).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                            worksheet3.Cell(nrow + 10, 2).Style.Border.InsideBorder = XLBorderStyleValues.Thin; worksheet3.Cell(nrow + 10, 2).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                            worksheet3.Cell(nrow + 10, 3).Style.Border.InsideBorder = XLBorderStyleValues.Thin; worksheet3.Cell(nrow + 10, 3).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                            worksheet3.Cell(nrow + 10, 4).Style.Border.InsideBorder = XLBorderStyleValues.Thin; worksheet3.Cell(nrow + 10, 4).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                            worksheet3.Cell(nrow + 10, 5).Style.Border.InsideBorder = XLBorderStyleValues.Thin; worksheet3.Cell(nrow + 10, 5).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                            worksheet3.Cell(nrow + 10, 6).Style.Border.InsideBorder = XLBorderStyleValues.Thin; worksheet3.Cell(nrow + 10, 6).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                            worksheet3.Cell(nrow + 10, 7).Style.Border.InsideBorder = XLBorderStyleValues.Thin; worksheet3.Cell(nrow + 10, 7).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                                                                                                                                       
                            worksheet4.Cell(nrow + 10, 1).Style.Border.InsideBorder = XLBorderStyleValues.Thin; worksheet4.Cell(nrow + 10, 1).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                            worksheet4.Cell(nrow + 10, 2).Style.Border.InsideBorder = XLBorderStyleValues.Thin; worksheet4.Cell(nrow + 10, 2).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                            worksheet4.Cell(nrow + 10, 3).Style.Border.InsideBorder = XLBorderStyleValues.Thin; worksheet4.Cell(nrow + 10, 3).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                            worksheet4.Cell(nrow + 10, 4).Style.Border.InsideBorder = XLBorderStyleValues.Thin; worksheet4.Cell(nrow + 10, 4).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                            worksheet4.Cell(nrow + 10, 5).Style.Border.InsideBorder = XLBorderStyleValues.Thin; worksheet4.Cell(nrow + 10, 5).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                            worksheet4.Cell(nrow + 10, 6).Style.Border.InsideBorder = XLBorderStyleValues.Thin; worksheet4.Cell(nrow + 10, 6).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                            worksheet4.Cell(nrow + 10, 7).Style.Border.InsideBorder = XLBorderStyleValues.Thin; worksheet4.Cell(nrow + 10, 7).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                            worksheet4.Cell(nrow + 10, 8).Style.Border.InsideBorder = XLBorderStyleValues.Thin; worksheet4.Cell(nrow + 10, 8).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                            lastRow = nrow + 10;
                        }

                    }
                    worksheet2.Range("A10:F" + lastRow).Style.Font.Bold = false;
                    worksheet2.Range("A10:F" + lastRow).Style.Font.FontName = "Arial";
                    worksheet2.Range("A10:F" + lastRow).Style.Font.FontSize = 10;

                    worksheet3.Range("A10:G" + lastRow).Style.Font.Bold = false;
                    worksheet3.Range("A10:G" + lastRow).Style.Font.FontName = "Arial";
                    worksheet3.Range("A10:G" + lastRow).Style.Font.FontSize = 10;

                    worksheet4.Column("A").Width = 7.71;

                    worksheet4.Range("A10:H" + lastRow).Style.Font.Bold = false;
                    worksheet4.Range("A10:H" + lastRow + 2).Style.Font.FontName = "Arial";
                    worksheet4.Range("A10:H" + lastRow + 2).Style.Font.FontSize = 10;

                    worksheet4.Range("A" + (lastRow + 2) + ":B" + (lastRow + 2)).Style.Font.FontName = "Tahoma";
                    worksheet4.Range("A" + (lastRow + 2) + ":B" + (lastRow + 2)).Style.Font.FontSize = 12;
                    worksheet4.Cell("A" + (lastRow + 2)).Value = "Note :";
                    worksheet4.Range("A" + (lastRow + 2)).Style.Font.Bold = true;
                    worksheet4.Range("B" + (lastRow + 2) + ":D" + (lastRow + 2)).Merge().Value = "O : Exist in System and already deleted.";
                    worksheet4.Range("B" + (lastRow + 2) + ":D" + (lastRow + 2)).Style.Font.Bold = false;
                    worksheet4.Range("B" + (lastRow + 2) + ":D" + (lastRow + 2)).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;




                    worksheet2.Range("G10:AH" + lastRow).Style.Border.BottomBorder = XLBorderStyleValues.Dotted;
                    worksheet2.Range("G10:AH" + lastRow).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                    worksheet2.Range("G" + lastRow + ":AH" + lastRow).Style.Border.BottomBorder = XLBorderStyleValues.Thin;

                   
                    string pathload =  "\\Content\\Excel\\" + fileName;
                    string pathfull = System.IO.Directory.GetCurrentDirectory() + "\\wwwroot\\Content\\Excel\\" + fileName;
                    using (var stream = new MemoryStream())
                    {
                        Workbook.SaveAs(stream);
                        Workbook.SaveAs(pathfull);
                        //stream.ToArray();
                        //return File(stream, "application/octet-stream", fileName);
                        FileStream fileStream = new FileStream(pathfull, FileMode.Open, FileAccess.Read);
                        //return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
                        return Json(pathload);
                    }



                }

            }
            catch (Exception ex)
            {
                throw ex;
            }


        }

        public JsonResult ExcelTemplate()
        {
            try
            {
                using (var Workbook = new XLWorkbook())
                {
                    IXLWorksheet worksheet =
                   Workbook.Worksheets.Add("sheet1");
                    //-----------------------header
                    worksheet.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    worksheet.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    worksheet.Range("A1:C1").Merge().Value = "RESIGNATION";
                    worksheet.Cell("A1").Style.Font.Bold = true;
                    worksheet.Style.Font.FontName = "Arial";
                    worksheet.Cell("A1").Style.Font.FontSize = 14;
                    worksheet.Range("A1:C1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                    worksheet.Range("A1:C1").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    worksheet.Row(1).Height = 34.50;
                   
                    //---------------------column
                    worksheet.Range("A2:A3").Merge().Value = "NO.";
                     worksheet.Range("B2:B3").Merge().Value = "CODE";
                     worksheet.Range("C2:D3").Merge().Value = "NAME";
                     worksheet.Range("E2:E3").Merge().Value = "POSITION";
                     worksheet.Range("F2:F3").Merge().Value = "LEVEL";
                     worksheet.Range("G2:G3").Merge().Value = "SECT.";
                     worksheet.Range("H2:H3").Merge().Value = "DEPT.";
                     worksheet.Range("I2:I3").Merge().Value = "DIV.";
                     worksheet.Range("J2:J3").Merge().Value = " HQ.";
                     worksheet.Range("K2:K3").Merge().Value = "RESIGNED";

                    //--------------------style Column
                    worksheet.Range("A2:K3").Style.Font.FontSize = 8;
                    worksheet.Range("A2:K3").Style.Fill.BackgroundColor = XLColor.FromArgb(204, 255, 204);
                   

                    worksheet.Range("A4:K5").Style.Font.FontSize = 9;
                    worksheet.Range("C4:D5").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                    worksheet.Range("C4:D5").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    worksheet.Range("A2:K5").Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                    worksheet.Range("A2:K5").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    worksheet.SheetView.FreezeColumns(3);
                    worksheet.SheetView.FreezeRows(3);
                    worksheet.Column(1).Width = 4.29;  worksheet.Column(2).Width = 8.43;
                    worksheet.Column(3).Width = 19.43;  worksheet.Column(4).Width = 4.71;
                    worksheet.Column(5).Width = 20.71;  worksheet.Column(6).Width = 3.86;
                    worksheet.Column(7).Width = 48.14;  worksheet.Column(8).Width = 34.71;
                    worksheet.Column(9).Width = 33.43;  worksheet.Column(10).Width = 15.57; worksheet.Column(11).Width = 10;
                    worksheet.Row(2).Height = 24.75; worksheet.Row(3).Height = 30;
                    worksheet.Row(4).Height = 30; worksheet.Row(5).Height = 30;
                    worksheet.RowHeight = 30;
                    worksheet.Range("F2:F3").Style.Alignment.SetTextRotation(90);
                    //---------------------


                    string pathload = "\\Content\\Excel\\OperatorResignTemplate.xlsx";
                    string pathfull = System.IO.Directory.GetCurrentDirectory() + "\\wwwroot\\Content\\Excel\\OperatorResignTemplate.xlsx";
                    using (var stream = new MemoryStream())
                    {
                        Workbook.SaveAs(stream);
                        Workbook.SaveAs(pathfull);
                        //stream.ToArray();
                        //return File(stream, "application/octet-stream", fileName);
                        FileStream fileStream = new FileStream(pathfull, FileMode.Open, FileAccess.Read);
                        //return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
                        return Json(pathload);
                    }
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }
             
        }

        public string Check_Operator(string Servername, string OPID)
        {

            try
            {
                using (var client = new HttpClient())
                {

                    var pathBase = HttpContext.Request.GetDisplayUrl();

                    var host = new Uri(pathBase).Authority;
                    var test = "http://" + host + "/" + Servername + "/" + OPID;
                    HttpResponseMessage resp = client.GetAsync("http://" + host + "/api/" + Servername + "/" + OPID).Result;
                    resp.EnsureSuccessStatusCode();

                    var result = resp.Content.ReadAsStringAsync().Result;

                    return result;
                }



            }
            catch (Exception ex)
            {
                return ex.ToString();
            }

        }

        public JsonResult GetDdl_Resign()
        {

            var query = "SELECT ResignDateMaking  FROM [ImportExportDB].[dbo].[OperatorsResign] group by ResignDateMaking order by ResignDateMaking DESC";
            var mgrSql = new mgrSQLConnect(configuration);

            var dt = mgrSql.GetDatatables(query);
            var session = HttpContext.Session.GetString("SessionID");
            var session_fullname = HttpContext.Session.GetString("Session_fullname");
            if (session == "" || session_fullname == "")
            {
                RedirectToAction("Index", "Home");
                return Json(data: "");
            }
            else
            {
                return Json(data: dt); 
            }
          
        }

        public IActionResult FrmTest()
        {
            return View();
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

        public JsonResult GetOperetorResign()
        {
            using (var client = new HttpClient())
            {

                var pathBase = HttpContext.Request.GetDisplayUrl();

                var host = new Uri(pathBase).Authority;
            
                HttpResponseMessage resp = client.GetAsync("http://" + host + "/api/GetOperatorsResign").Result;
                resp.EnsureSuccessStatusCode();

                var Dt = (DataTable)JsonConvert.DeserializeObject(resp.Content.ReadAsStringAsync().Result , (typeof(DataTable)));
                var JsonData = resp.Content.ReadAsStringAsync().Result;
                //var result = Json(new { data = Dt });
                return Json(data : Dt);
            }
          
           
        }

        public ActionResult Frm_NewLayout()
        {
            return View();
        }


    }


}


