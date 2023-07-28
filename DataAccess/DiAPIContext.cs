
using SAPbobsCOM;
using System;
using Xeneff.SAPB1.DiAPI.Common;
using System.Configuration;
namespace Xeneff.SAPB1.DiAPI.DataAccess
{

    public abstract class DiAPIContext
    {
        private Company company = null;
        private int errorCode = 0;
        private string errorMessage = "";
        string _sapServer = ConfigurationManager.AppSettings["Server"];
        string _dbUser = ConfigurationManager.AppSettings["DBUser"];
        string _dbPassword = ConfigurationManager.AppSettings["DBPassword"];
        string _companyDB = ConfigurationManager.AppSettings["DevDatabase"];
        string _userName = ConfigurationManager.AppSettings["SapUser"];
        string _password = ConfigurationManager.AppSettings["SapUserPassword"];
        string _licenseServer = ConfigurationManager.AppSettings["License"];
        string _SLDServer = ConfigurationManager.AppSettings["SLD"];
        public int Connect()
        {
            int connectionResult = Constants.DefaultDiApiResult;
            if (company == null)
            {

                company = new Company();
                company.Server = Environment.GetEnvironmentVariable(_sapServer);
                company.DbUserName = Environment.GetEnvironmentVariable(_dbUser);
                company.DbPassword = Environment.GetEnvironmentVariable(_dbPassword);
                company.CompanyDB = Environment.GetEnvironmentVariable(_companyDB);
                company.UserName = Environment.GetEnvironmentVariable(_userName);
                company.Password = Environment.GetEnvironmentVariable(_password);
                company.LicenseServer = Environment.GetEnvironmentVariable(_licenseServer);
                company.SLDServer = Environment.GetEnvironmentVariable(_SLDServer);
                company.DbServerType = BoDataServerTypes.dst_MSSQL2019;
                company.language = BoSuppLangs.ln_Turkish_Tr;
                connectionResult = company.Connect();
            }
            else if (company.Connected == false)
            {
                connectionResult = company.Connect();
            }
            else
            {
                connectionResult = (company.Connected == true ? Constants.DiApiSuccess : Constants.DefaultDiApiResult);
            }
            if (connectionResult != Constants.DiApiSuccess)
                company.GetLastError(out errorCode, out errorMessage);

            return connectionResult;
        }

        public void DiApiDisconnect()
        {
            if (company != null)
            {
                company.Disconnect();
                company = null;
                GC.Collect();
            }
        }
        public Company GetCompany()
        {
            return company;
        }
        public int GetErrorCode()
        {
            return errorCode;
        }
        public string GetErrorMessage()
        {
            return errorMessage;
        }
        public void ReleaseSapBobsObject(object diApiObject)
        {
            if (diApiObject != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(diApiObject);
                diApiObject = null;
            }
        }
        public Recordset RecordSetSqlRun(string query)
        {
            Recordset oRecordSet = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
            oRecordSet.DoQuery(query);
            return oRecordSet;
        }
    }

}
