using Import_Freight_BOI.Models.TSQL;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using RISTExamOnlineProject.Models.db;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Extensions.DependencyInjection.Extensions;

namespace Import_Freight_BOI.Api
{
   
    [ApiController]
    public class OperatorsController : ControllerBase
    {
        private readonly IConfiguration _configuration;
        private readonly IHttpContextAccessor httpContextAccessor;

        public OperatorsController( IConfiguration configuration,
           IHttpContextAccessor httpContextAccessor)
        {
            _configuration = configuration;
            this.httpContextAccessor = httpContextAccessor;

        }
        //[Route("api/[controller]")]
        //public int Get()
        //{
        //    //DataTable dt = new DataTable();
        //    string Strsql = "SELECT * FROM [MCRMaterialControl].[dbo].[Operator]";
        //    var ObjRun = new mgrSQLConnect(_configuration);
        //    int response = ObjRun.Get_RTHSRVM03(Strsql);
        //    return response;
        //}

        [Route("api/MCRBCP/{id}")]
        public DataTable MCRBCP(string id)
        {
            string Strsql = "select expire from [MCR].[dbo].[Operator] where OPID ='" + id + "'";
            var ObjRun = new mgrSQLConnect(_configuration);
            DataTable response = new DataTable();
            response = ObjRun.GetDatatables_SrvBCP(Strsql);
            return response;
        }

        [Route("api/MCRMaterialControl/{id}")]
        public string MCRMaterialControl(string id)
        {
            //DataTable dt = new DataTable();
            //string Strsql = "SELECT * FROM [MCRMaterialControl].[dbo].[Operator] where OPID ='" + id  + "'";
            //var ObjRun = new mgrSQLConnect(_configuration);
            //string response = ObjRun.Get_RTHSRVM03(Strsql);
            string Strsql = "EXEC sprUpdatePasswordResign 'MCRMaterialControl','" + id + "'";
            var ObjRun = new mgrSQLConnect(_configuration);
            string response = ObjRun.ExcuteStore(Strsql);
            return response;
        }

        [Route("api/MCR/{id}")]
        public string MCR(string id)
        {
            //DataTable dt = new DataTable();
            //string Strsql = "SELECT * FROM [MCR].[dbo].[Operator] where OPID ='" + id + "'";
            //var ObjRun = new mgrSQLConnect(_configuration);
            //string response = ObjRun.Get_RTHSRVM03(Strsql);
           
            string Strsql = "EXEC sprUpdatePasswordResign 'MCR','" + id + "'";
            var ObjRun = new mgrSQLConnect(_configuration);
            string response = ObjRun.ExcuteStore(Strsql);
            return response;
        }

        [Route("api/MCRLT/{id}")]
        public string MCRLT(string id)
        {
            //DataTable dt = new DataTable();
            //string Strsql = "SELECT * FROM [MCRLT].[dbo].[Operator] where OPID ='" + id + "'";
            //var ObjRun = new mgrSQLConnect(_configuration);
            //string response = ObjRun.Get_RTHSRVM03(Strsql);
       
            string Strsql = "EXEC sprUpdatePasswordResign 'MCRLT','" + id + "'";
            var ObjRun = new mgrSQLConnect(_configuration);
            string response = ObjRun.ExcuteStore(Strsql);
            return response;
        }
        [Route("api/MCRSpareParts/{id}")]
        public string MCRSpareParts(string id)
        {
            //DataTable dt = new DataTable();
            //string Strsql = "SELECT * FROM [MCRSpareParts].[dbo].[Operator] where UserID ='" + id + "'";
            //var ObjRun = new mgrSQLConnect(_configuration);
            //string response = ObjRun.Get_RTHSRVM03(Strsql);
            string Strsql = "EXEC sprUpdatePasswordResign 'MCRSpareParts','" + id + "'";
            var ObjRun = new mgrSQLConnect(_configuration);
            string response = ObjRun.ExcuteStore(Strsql);

            return response;
        }

        [Route("api/MCRProcessControl/{id}")]
        public string MCRProcessControl(string id)
        {
            //DataTable dt = new DataTable();
            //string Strsql = "SELECT * FROM [MCRProcessControl].[dbo].[Operator] where OperatorID ='" + id + "'";
            //var ObjRun = new mgrSQLConnect(_configuration);
            //string response = ObjRun.Get_Con_RTHSRVMCR01(Strsql);

            string Strsql = "EXEC sprUpdatePasswordResign 'MCRProcessControl','" + id + "'";
            var ObjRun = new mgrSQLConnect(_configuration);
            string response = ObjRun.ExcuteStore(Strsql);

            return response;
        }


        [Route("api/MCRTPProcessControl/{id}")]
        public string MCRTPProcessControl(string id)
        {
            //DataTable dt = new DataTable();
            //string Strsql = "SELECT * FROM [MCRProcessControl].[dbo].[Operator] where OperatorID ='" + id + "'";
            //var ObjRun = new mgrSQLConnect(_configuration);
            //string response = ObjRun.Get_Con_RTHSRVMCR01(Strsql);
            string Strsql = "EXEC sprUpdatePasswordResign 'MCRTPProcessControl','" + id + "'";
            var ObjRun = new mgrSQLConnect(_configuration);
            string response = ObjRun.ExcuteStore(Strsql);

            return response;
        }


        [Route("api/OPMSparepart/{id}")]
        public string OPMSparepart(string id)
        {
            //DataTable dt = new DataTable();
            //string Strsql = "SELECT * FROM [OPMSpareParts].[dbo].[Operator] where UserID ='" + id + "'";
            //var ObjRun = new mgrSQLConnect(_configuration);
            //string response = ObjRun.Get_RTHSRVOPM01(Strsql);

            string Strsql = "EXEC sprUpdatePasswordResign 'OPMSparepart','" + id + "'";
            var ObjRun = new mgrSQLConnect(_configuration);
            string response = ObjRun.ExcuteStore(Strsql);

            return response;
        }


        [Route("api/OPMProcessControl/{id}")]
        public string OPMProcessControl(string id)
        {
            //DataTable dt = new DataTable();
            //string Strsql = "SELECT * FROM [OPMProcessControl].[dbo].[Operator] where OperatorID ='" + id + "'";
            //var ObjRun = new mgrSQLConnect(_configuration);
            //string response = ObjRun.Get_RTHSRVOPM02(Strsql);

            string Strsql = "EXEC sprUpdatePasswordResign 'OPMSparepart','" + id + "'";
            var ObjRun = new mgrSQLConnect(_configuration);
            string response = ObjRun.ExcuteStore(Strsql);

            return response;
        }


        [Route("api/TcOPOLLO/{id}")]
        public string TcOPOLLO(string id)
        {
            //DataTable dt = new DataTable();
            //string Strsql = "SELECT * FROM [APOLLO].[dbo].[Operator] where OPID ='" + id + "'";
            //var ObjRun = new mgrSQLConnect(_configuration);
            //string response = ObjRun.Get_RTHSRV07(Strsql);

            string Strsql = "EXEC sprUpdatePasswordResign 'OPMSparepart','" + id + "'";
            var ObjRun = new mgrSQLConnect(_configuration);
            string response = ObjRun.ExcuteStore(Strsql);

            return response;
        }


        [Route("api/TCNonDB/{id}")]
        public string TCNonDB(string id)
        {
            //DataTable dt = new DataTable();
            //string Strsql = "SELECT * FROM [TCNonDB].[dbo].[Operator] where OPID ='" + id + "'";
            //var ObjRun = new mgrSQLConnect(_configuration);
            //string response = ObjRun.Get_RTHSRV07(Strsql);

            string Strsql = "EXEC sprUpdatePasswordResign 'TCNonDB','" + id + "'";
            var ObjRun = new mgrSQLConnect(_configuration);
            string response = ObjRun.ExcuteStore(Strsql);
            return response;
        }
        /// Not Using
        [Route("api/TCMaterialControl/{id}")]
        public string TCMaterialControl(string id)
        {
            //DataTable dt = new DataTable();
            //string Strsql = "SELECT * FROM [TCMaterialControl].[dbo].[Operator] where OPID ='" + id + "'";
            //var ObjRun = new mgrSQLConnect(_configuration);
            //string response = ObjRun.Get_RTHSRV07(Strsql);

            string Strsql = "EXEC sprUpdatePasswordResign 'TCMaterialControl','" + id + "'";
            var ObjRun = new mgrSQLConnect(_configuration);
            string response = ObjRun.ExcuteStore(Strsql);

            return response;
        }


        [Route("api/TROEM/{id}")]
        public string TROEM(string id)
        {
            //DataTable dt = new DataTable();
            //string Strsql = "SELECT * FROM [TROEM].[dbo].[SanyoOperator]  where OPID ='" + id + "'";
            //var ObjRun = new mgrSQLConnect(_configuration);
            //string response = ObjRun.Get_RTHSRV11(Strsql);
            string Strsql = "EXEC sprUpdatePasswordResign 'TROEM','" + id + "'";
            var ObjRun = new mgrSQLConnect(_configuration);
            string response = ObjRun.ExcuteStore(Strsql);
            return response;
        }
       
        /// Not Using
        
        [Route("api/OEM/{id}")]
        public string OEM(string id)
        {
            //DataTable dt = new DataTable();
            //string Strsql = "SELECT * FROM [OEM].[dbo].[OEMOperator]  where OperatorID ='" + id + "'";
            //var ObjRun = new mgrSQLConnect(_configuration);
            //string response = ObjRun.Get_RTHSRV11(Strsql);

            string Strsql = "EXEC sprUpdatePasswordResign 'OEM','" + id + "'";
            var ObjRun = new mgrSQLConnect(_configuration);
            string response = ObjRun.ExcuteStore(Strsql);
            return response;
        }

        ///Sqlserver 2000

        [Route("api/StagnationLot/{id}")]
        public string StagnationLot(string id)
        {
            //DataTable dt = new DataTable();
            string Strsql = "SELECT * FROM [StagnationLot].[dbo].[OEMOperator]  where OperatorID ='" + id + "'";
            var ObjRun = new mgrSQLConnect(_configuration);
            string response = ObjRun.Get_RTHSRV04(Strsql);
            return response;
        }
        [Route("api/TRLT/{id}")]
        public string TRLT(string id)
        {
            //DataTable dt = new DataTable();
            string Strsql = "SELECT * FROM [TRLT].[dbo].[Operator]  where OPID ='" + id + "'";
            var ObjRun = new mgrSQLConnect(_configuration);
            string response = ObjRun.Get_RTHSRV04(Strsql);
            return response;
        }
        [Route("api/AlarmDB/{id}")]
        public string AlarmDB(string id)
        {
            //DataTable dt = new DataTable();
            //string Strsql = "SELECT * FROM [AlarmDB].[dbo].[T_Operator]  where UserID ='" + id + "'";
            //var ObjRun = new mgrSQLConnect(_configuration);
            //string response = ObjRun.Get_RTHSRVWSlip(Strsql);

            string Strsql = "EXEC sprUpdatePasswordResign 'AlarmDB','" + id + "'";
            var ObjRun = new mgrSQLConnect(_configuration);
            string response = ObjRun.ExcuteStore(Strsql);

            return response;
        }


        [Route("api/NonDB/{id}")]
        public string NonDB(string id)
        {
            //DataTable dt = new DataTable();
            //string Strsql = "SELECT * FROM [NonDB].[dbo].[Operator]  where OPID ='" + id + "'";
            //var ObjRun = new mgrSQLConnect(_configuration);
            //string response = ObjRun.Get_RTHSRVWSlip(Strsql);


            string Strsql = "EXEC sprUpdatePasswordResign 'NonDB','" + id + "'";
            var ObjRun = new mgrSQLConnect(_configuration);
            string response = ObjRun.ExcuteStore(Strsql);

            return response;
        }

        [Route("api/Parts/{id}")]
        public string Parts(string id)
        {
            //DataTable dt = new DataTable();
            string Strsql = "SELECT * FROM [Parts].[dbo].[Operator]  where Emp_No ='" + id + "'";
            var ObjRun = new mgrSQLConnect(_configuration);
            string response = ObjRun.Get_RTHSRVWSlip(Strsql);
            return response;
        }

        [Route("api/PlatingDB/{id}")]
        public string PlatingDB(string id)
        {
            //DataTable dt = new DataTable();
            //string Strsql = "SELECT * FROM [PlatingDB].[dbo].[OP]  where OPID ='" + id + "'";
            //var ObjRun = new mgrSQLConnect(_configuration);
            //string response = ObjRun.Get_RTHSRVWSlip(Strsql);

            string Strsql = "EXEC sprUpdatePasswordResign 'PlatingDB','" + id + "'";
            var ObjRun = new mgrSQLConnect(_configuration);
            string response = ObjRun.ExcuteStore(Strsql);

            return response;
        }

        [Route("api/TRProcessControl/{id}")]
        public string TRProcessControl(string id)
        {
            //DataTable dt = new DataTable();
            //string Strsql = "SELECT * FROM [TRProcessControl].[dbo].[Operator]  where OperatorID ='" + id + "'";
            //var ObjRun = new mgrSQLConnect(_configuration);
            //string response = ObjRun.Get_RTHSRVTR01(Strsql);

            string Strsql = "EXEC sprUpdatePasswordResign 'TRProcessControl','" + id + "'";
            var ObjRun = new mgrSQLConnect(_configuration);
            string response = ObjRun.ExcuteStore(Strsql);

            return response;
        }


        [Route("api/GFDReport/{id}")]
        public string GFDReport(string id)
        {
           
            string Strsql = "EXEC sprUpdatePasswordResign 'GFDReport','" + id + "'";
            var ObjRun = new mgrSQLConnect(_configuration);
            string response = ObjRun.ExcuteStore(Strsql);

            return response;
        }

        [Route("api/OneWord/{id}")]
        public string OneWord(string id)
        {
          

            string Strsql = "EXEC sprUpdatePasswordResign 'OneWord','" + id + "'";
            var ObjRun = new mgrSQLConnect(_configuration);
            string response = ObjRun.ExcuteStore(Strsql);

            return response;
        }

        [Route("api/TCCostDb/{id}")]
        public string TCCostDb(string id)
        {
            
            string Strsql = "EXEC sprUpdatePasswordResign 'TCCostDb','" + id + "'";
            var ObjRun = new mgrSQLConnect(_configuration);
            string response = ObjRun.ExcuteStore(Strsql);

            return response;
        }

        [Route("api/TRCostDb/{id}")]
        public string TRCostDb(string id)
        {
            
            string Strsql = "EXEC sprUpdatePasswordResign 'TRCostDb','" + id + "'";
            var ObjRun = new mgrSQLConnect(_configuration);
            string response = ObjRun.ExcuteStore(Strsql);

            return response;
        }


        [Route("api/GetOperatorsResign")]
        public DataTable GetOperatorsResign(string id)
        {
            DataTable dt = new DataTable();
            string Strsql = "SELECT * FROM [ImportExportDB].[dbo].[OperatorsResign]";
            var ObjRun = new mgrSQLConnect(_configuration);
            dt= ObjRun.GetDatatables(Strsql);
           
            return dt;
        
        }



        [Route("api/GetOperatorResignChart")]
        public DataTable GetOperatorResignChart()
        {
          
            string Strsql = "select * from vewOperatorResignChart";
            var ObjRun = new mgrSQLConnect(_configuration);
            DataTable response = new DataTable();
            response =  ObjRun.GetDatatables(Strsql);
            return response;
        }


        [Route("api/GetGroupOperatorResign")]
        public DataTable GetGroupOperatorResign()
        {
            string Strsql = "select * from vewGroupHqResign";
            var ObjRun = new mgrSQLConnect(_configuration);
            DataTable response = new DataTable();
            response = ObjRun.GetDatatables(Strsql);
            return response;
        }






    }

}
