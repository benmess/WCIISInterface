using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.ServiceModel.Web;
using System.ServiceModel.Activation;
using System.Text;
using System.Web;
using System.Web.Services;
using System.Web.Services.Protocols;
using System.Security.Cryptography.X509Certificates;
using System.IO;
using System.Data.SqlClient;
using System.Data;
using System.Net;
using word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using DocumentFormat.OpenXml.Packaging;

namespace WcfWCService
{
    // NOTE: You can use the "Rename" command on the "Refactor" menu to change the class name "Service1" in code, svc and config file together.
    // NOTE: In order to launch WCF Test Client for testing this service, please select Service1.svc or Service1.svc.cs at the Solution Explorer and start debugging.
    [ServiceBehavior(Namespace = "http://regain.com/rest")]
    public class Service1 : IService1
    {
        public class rtnInt
        {
            public int iReturnValue;
            public bool bReturnValue;
        }
        public class rtnString
        {
            public string sReturnValue;
            public string sReturnValueExtra1;
            public bool bReturnValue;
            public int iLineNumber = 0;
        }
        public class rtnTerms
        {
            public bool bReturnValue;
            public int iCoreNo;
            public string sCoreLabel;
            public string sWireNo;
            public string sFromTermination;
            public int iFromLineNumber = 0;
            public string sToTermination;
            public int iToLineNumber = 0;
        }

        public string CookieLogin(string sUsername, string sPassword, string sWebAppId)
        {
            string sSessionId;
            //                string sNothing = "";
            try
            {
                Environment env = new Environment();
                string sServerRoot = env.Get_Environment_String_Value("ServerWithPort");
                if (sServerRoot.EndsWith("/"))
                    sServerRoot = sServerRoot.Substring(0, sServerRoot.Length - 1);
                string uri = String.Format("http://" + sServerRoot + "/Regain/rest/cookielogin/" + sUsername + "/" + sPassword + "/" + sWebAppId);
                HttpWebRequest myRequest = (HttpWebRequest)WebRequest.Create(uri);
                myRequest.Method = "GET";
                myRequest.Timeout = 15000;
                WebResponse thePage = myRequest.GetResponse();
                using (var reader = new StreamReader(thePage.GetResponseStream()))
                {
                    sSessionId = reader.ReadToEnd(); // do something fun...
                }
            }
            catch (Exception ex)
            {
                sSessionId = ex.Message;
            }

            return sSessionId;

            //ArrayList arrUser = GetUserDetails(sUsername);

            //bool bPassCheck = BCrypt.Verify(arrUser[1].ToString(), sPassword);

            //if (bPassCheck && sUsername == arrUser[0].ToString())
            //{
            //    sSessionId = HttpContext.Current.Session.SessionID;
            //    SetUser_SessionId(sUsername, sSessionId);
            //    return sSessionId;
            //}
            //else
            //{
            //    return "";
            //}
        }

        public bool IsUserLoggedIn(string sSessionId, string sUsername, int iWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUsername, iWebAppId))
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        public bool IsExternalUserValid(string sSessionId, string sUsername, int iWebAppId)
        {

            if (Get_User_From_SessionID(sSessionId, iWebAppId) != sUsername)
            {
                return false;
            }
            else
            {
                return true;
            }

        }

        public ArrayList GetUserDetails(string sUser)
        {
            RecordSet rst = new RecordSet();
            ArrayList rtnArray = new ArrayList();
            string sSQL = "select UserId, Username, Password, Email, isnull(Fullname, '') as Fullname from tblUser where Username = '" + sUser + "'";
            DataSet ds = rst.OpenRecordset(sSQL, rst.SqlConnectionStr());

            if (rst.m_RecordCount > 0)
            {
                rtnArray.Add(rst.Get_NVarchar(ds, "Username", 0));
                rtnArray.Add(rst.Get_NVarchar(ds, "Password", 0));
                rtnArray.Add(rst.Get_NVarchar(ds, "Fullname", 0));
                rtnArray.Add(rst.Get_NVarchar(ds, "Email", 0));
            }

            ds.Dispose();

            return rtnArray;

        }
        public string Get_User_From_SessionID(string sSessionId, int iWebAppId)
        {
            RecordSet rst = new RecordSet();
            rst.SetWebApp(iWebAppId);
            string sRtn = "";
            string sSQL = "SELECT Username FROM tblUserSession  Where SessionId = '" + sSessionId + "'";
            DataSet ds = rst.OpenRecordset(sSQL, rst.SqlConnectionStr());

            if (rst.m_RecordCount > 0)
            {
                sRtn = rst.Get_NVarchar(ds, "Username", 0);
            }

            ds.Dispose();

            return sRtn;


        }


        public bool SetUser_SessionId(string sUserId, string sSessionId, int iWebAppId)
        {
            string sSQL;
            RecordSet rst = new RecordSet();
            rst.SetWebApp(iWebAppId);

            sSQL = "UPDATE tblUser SET SessionId = '" + sSessionId + "' WHERE Username = '" + sUserId + "'";
            bool bRtn = rst.ExecuteSQL(sSQL);
            return bRtn;
        }

        public void Update_User_Time(string sUserId, string sSessionId)
        {
            StoredProc SP = new StoredProc();
            SP.SetProcName("SP_UpdateUserSession");
            SP.SetParam("@pvchUsername", sUserId);
            SP.SetParam("@pvchSession", sSessionId);
            SP.RunStoredProc();
        }

        public string GetConstantValue(string sConstantName, int iWebAppId)
        {
            RecordSet rst = new RecordSet();
            rst.SetWebApp(iWebAppId);
            string sRtn = "";
            string sSQL = "SELECT isnull(Value,'') as Value FROM tblConstants WHERE Name = '" + sConstantName + "'";
            DataSet ds = rst.OpenRecordset(sSQL, rst.SqlConnectionStr());

            if (rst.m_RecordCount > 0)
            {
                sRtn = rst.Get_NVarchar(ds, "Value", 0);
            }

            ds.Dispose();

            return sRtn;


        }

        public string GetTemplateName(int iTemplateCode, int iWebAppId)
        {
            RecordSet rst = new RecordSet();
            rst.SetWebApp(iWebAppId);
            string sRtn = "";
            string sSQL = "select TemplateFilename from tblDocTemplates where Code = " + iTemplateCode;
            DataSet ds = rst.OpenRecordset(sSQL, rst.SqlConnectionStr());

            if (rst.m_RecordCount > 0)
            {
                sRtn = rst.Get_NVarchar(ds, "TemplateFilename", 0);
            }

            ds.Dispose();

            return sRtn;


        }

        public string add(string sSessionId, string sUserId, string a, string b, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                int ia = Convert.ToInt16(a);
                int ib = Convert.ToInt16(b);
                int c = client2.add(ia, ib);
                return c.ToString();
            }
        }

        public string simpleadd(string a, string b)
        {
            try
            {
                ExampleService.MyJavaService3Client client2 = GetWCService();
                int ia = Convert.ToInt16(a);
                int ib = Convert.ToInt16(b);
                int c = client2.add(ia, ib);
                return c.ToString();
            }
            catch
            {
                WebOperationContext ctx = WebOperationContext.Current;
                ctx.OutgoingResponse.StatusCode = System.Net.HttpStatusCode.NotFound;
                return null;
            }
        }

        public string CreateWCDoc(string sSessionId, string sUserId, string sDocNo, string sDocName, string sProductName, string sDocType, string sFolderNameAndPath,
                                  string sLongDesc, string sOriginator, string sOriginatorDocId, string sJobCode, string sRevision, string sCheckInComments, string iProdOrLibrary, string sWebAppId)
        {

            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                int iiProdOrLibrary = Convert.ToInt16(iProdOrLibrary);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                string[] sAttributeNames = new string[2];
                string[] sAttributeValues = new string[2];
                string[] sAttributeTypes = new string[2];

                sAttributeNames[0] = "LongDescription";
                sAttributeNames[1] = "Originator";


                sAttributeValues[0] = sLongDesc;
                sAttributeValues[1] = sOriginator;

                sAttributeTypes[0] = "string";
                sAttributeTypes[1] = "string";

                if (sOriginatorDocId != "")
                {
                    Array.Resize<string>(ref sAttributeNames, 3);
                    Array.Resize<string>(ref sAttributeValues, 3);
                    Array.Resize<string>(ref sAttributeTypes, 3);
                    sAttributeNames[2] = "OrigDocId";
                    sAttributeValues[2] = sOriginatorDocId;
                    sAttributeTypes[2] = "string";

                    if (sJobCode != "")
                    {
                        Array.Resize<string>(ref sAttributeNames, 4);
                        Array.Resize<string>(ref sAttributeValues, 4);
                        Array.Resize<string>(ref sAttributeTypes, 4);
                        sAttributeNames[3] = "JobCode";
                        sAttributeValues[3] = sJobCode;
                        sAttributeTypes[3] = "string";
                    }

                }
                else
                {
                    if (sJobCode != "")
                    {
                        Array.Resize<string>(ref sAttributeNames, 3);
                        Array.Resize<string>(ref sAttributeValues, 3);
                        Array.Resize<string>(ref sAttributeTypes, 3);
                        sAttributeNames[2] = "JobCode";
                        sAttributeValues[2] = sJobCode;
                        sAttributeTypes[2] = "string";
                    }

                }

                //                int iChangeRevision = Convert.ToInt16(sChangeRevision);
                return client2.doccreate(sDocNo, sDocName, sProductName, sDocType, sFolderNameAndPath, sRevision, sAttributeNames, sAttributeValues, sCheckInComments, iiProdOrLibrary, Convert.ToInt16(sWebAppId));
            }
        }

        public string CreateWCDoc2(string sSessionId, string sUserId, string sDocNo, string sDocName, string sProductName, string sDocType, string sFolderNameAndPath,
                                  string sDesc, string sOriginator, string sOriginatorDocId, string sJobCode, string sRevision, string sCheckInComments, string iProdOrLibrary, string sWebAppId)
        {

            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                int iiProdOrLibrary = Convert.ToInt16(iProdOrLibrary);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                string[] sAttributeNames = new string[2];
                string[] sAttributeValues = new string[2];

                sAttributeNames[0] = "description";
                sAttributeNames[1] = "Originator";


                sAttributeValues[0] = sDesc;
                sAttributeValues[1] = sOriginator;

                if (sOriginatorDocId != "")
                {
                    Array.Resize<string>(ref sAttributeNames, 3);
                    Array.Resize<string>(ref sAttributeValues, 3);
                    sAttributeNames[2] = "OrigDocId";
                    sAttributeValues[2] = sOriginatorDocId;

                    if (sJobCode != "")
                    {
                        Array.Resize<string>(ref sAttributeNames, 4);
                        Array.Resize<string>(ref sAttributeValues, 4);
                        sAttributeNames[3] = "JobCode";
                        sAttributeValues[3] = sJobCode;
                    }

                }
                else
                {
                    if (sJobCode != "")
                    {
                        Array.Resize<string>(ref sAttributeNames, 3);
                        Array.Resize<string>(ref sAttributeValues, 3);
                        sAttributeNames[2] = "JobCode";
                        sAttributeValues[2] = sJobCode;
                    }

                }


                return client2.doccreate(sDocNo, sDocName, sProductName, sDocType, sFolderNameAndPath, sRevision, sAttributeNames, sAttributeValues, sCheckInComments, iiProdOrLibrary, Convert.ToInt16(sWebAppId));
            }
        }

        public string CreateRequirementDoc(string sSessionId, string sUserId, string sDocNo, string sDocName, string sProductName, string sDocType, string sFolderNameAndPath,
                                          string sDesc, string sOriginator, string sOriginatorDocId, string sJobCode, string sRevision,
                                          string sTargetDate, string sForecastDate, string sActualDate, string sDateBasis, string sComments,
                                          string sCheckInComments, string iProdOrLibrary, string sWebAppId)
        {

            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                int iiProdOrLibrary = Convert.ToInt16(iProdOrLibrary);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                string[] sAttributeNames = new string[2];
                string[] sAttributeValues = new string[2];

                sAttributeNames[0] = "description";
                sAttributeNames[1] = "Originator";


                sAttributeValues[0] = sDesc;
                sAttributeValues[1] = sOriginator;

                if (sOriginatorDocId != "")
                {
                    Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                    Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                    sAttributeNames[sAttributeNames.Length - 1] = "OrigDocId";
                    sAttributeValues[sAttributeValues.Length - 1] = sOriginatorDocId;
                }

                if (sJobCode != "")
                {
                    Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                    Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                    sAttributeNames[sAttributeNames.Length - 1] = "JobCode";
                    sAttributeValues[sAttributeValues.Length - 1] = sJobCode;
                }

                if (sTargetDate != "")
                {
                    Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                    Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                    sAttributeNames[sAttributeNames.Length - 1] = "TargetDate";
                    sAttributeValues[sAttributeValues.Length - 1] = sTargetDate;
                }

                if (sForecastDate != "")
                {
                    Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                    Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                    sAttributeNames[sAttributeNames.Length - 1] = "ForecastDate";
                    sAttributeValues[sAttributeValues.Length - 1] = sForecastDate;
                }

                if (sActualDate != "")
                {
                    Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                    Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                    sAttributeNames[sAttributeNames.Length - 1] = "ActualDate";
                    sAttributeValues[sAttributeValues.Length - 1] = sActualDate;
                }

                if (sDateBasis != "")
                {
                    Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                    Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                    sAttributeNames[sAttributeNames.Length - 1] = "DateBasis";
                    sAttributeValues[sAttributeValues.Length - 1] = sDateBasis;
                }

                if (sComments != "")
                {
                    Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                    Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                    sAttributeNames[sAttributeNames.Length - 1] = "Comments";
                    sAttributeValues[sAttributeValues.Length - 1] = sComments;
                }

                return client2.doccreate(sDocNo, sDocName, sProductName, sDocType, sFolderNameAndPath, sRevision, sAttributeNames, sAttributeValues, sCheckInComments, iiProdOrLibrary, Convert.ToInt16(sWebAppId));
            }
        }

        public string CreatePlantEquipMaterialDoc(string sSessionId, string sUserId, string sDocNo, string sDocName, string sProductName, string sDocType, string sFolderNameAndPath,
                                                  string sDesc, string sOriginator, string sOriginatorDocId, string sJobCode, string sRevision,
                                                  string sFirstIssueDate, string sIssueForUseDate, string sFinalIssueDate, string sStatusComments, string sComments,
                                                  string sCheckInComments, string iProdOrLibrary, string sWebAppId)
        {

            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                int iiProdOrLibrary = Convert.ToInt16(iProdOrLibrary);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                string[] sAttributeNames = new string[2];
                string[] sAttributeValues = new string[2];

                sAttributeNames[0] = "description";
                sAttributeNames[1] = "Originator";


                sAttributeValues[0] = sDesc;
                sAttributeValues[1] = sOriginator;

                if (sOriginatorDocId != "")
                {
                    Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                    Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                    sAttributeNames[sAttributeNames.Length - 1] = "OrigDocId";
                    sAttributeValues[sAttributeValues.Length - 1] = sOriginatorDocId;
                }

                if (sJobCode != "")
                {
                    Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                    Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                    sAttributeNames[sAttributeNames.Length - 1] = "JobCode";
                    sAttributeValues[sAttributeValues.Length - 1] = sJobCode;
                }

                if (sFirstIssueDate != "")
                {
                    Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                    Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                    sAttributeNames[sAttributeNames.Length - 1] = "FirstIssueDate";
                    sAttributeValues[sAttributeValues.Length - 1] = sFirstIssueDate;
                }

                if (sIssueForUseDate != "")
                {
                    Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                    Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                    sAttributeNames[sAttributeNames.Length - 1] = "IssueForUseDate";
                    sAttributeValues[sAttributeValues.Length - 1] = sIssueForUseDate;
                }

                if (sFinalIssueDate != "")
                {
                    Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                    Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                    sAttributeNames[sAttributeNames.Length - 1] = "FinalIssueDate";
                    sAttributeValues[sAttributeValues.Length - 1] = sFinalIssueDate;
                }

                if (sStatusComments != "")
                {
                    Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                    Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                    sAttributeNames[sAttributeNames.Length - 1] = "StatusComments";
                    sAttributeValues[sAttributeValues.Length - 1] = sStatusComments;
                }

                if (sComments != "")
                {
                    Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                    Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                    sAttributeNames[sAttributeNames.Length - 1] = "Comments";
                    sAttributeValues[sAttributeValues.Length - 1] = sComments;
                }

                return client2.doccreate(sDocNo, sDocName, sProductName, sDocType, sFolderNameAndPath, sRevision, sAttributeNames, sAttributeValues, sCheckInComments, iiProdOrLibrary, Convert.ToInt16(sWebAppId));
            }
        }

        public string UpdateRequirementDoc(string sSessionId, string sUserId, string sDocNo, string sDocName,
                                          string sDesc, string sOriginator, string sOriginatorDocId,
                                          string sTargetDate, string sForecastDate, string sActualDate, string sDateBasis,
                                          string sComments, string sCheckInComments, string sWebAppId)
        {

            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                string[] sAttributeNames = new string[2];
                string[] sAttributeValues = new string[2];
                string[] sAttributeTypes = new string[2];

                sAttributeNames[0] = "description";
                sAttributeNames[1] = "Originator";


                sAttributeValues[0] = sDesc;
                sAttributeValues[1] = sOriginator;

                sAttributeTypes[0] = "string";
                sAttributeTypes[1] = "string";

                if (sOriginatorDocId != "")
                {
                    Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                    Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                    Array.Resize<string>(ref sAttributeTypes, sAttributeTypes.Length + 1);
                    sAttributeNames[sAttributeNames.Length - 1] = "OrigDocId";
                    sAttributeValues[sAttributeValues.Length - 1] = sOriginatorDocId;
                    sAttributeTypes[sAttributeTypes.Length - 1] = "string";
                }

                //if (sTargetDate != "")
                //{
                Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                Array.Resize<string>(ref sAttributeTypes, sAttributeTypes.Length + 1);
                sAttributeNames[sAttributeNames.Length - 1] = "TargetDate";
                sAttributeValues[sAttributeValues.Length - 1] = sTargetDate;
                sAttributeTypes[sAttributeTypes.Length - 1] = "string";
                //}

                //if (sForecastDate != "")
                //{
                Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                Array.Resize<string>(ref sAttributeTypes, sAttributeTypes.Length + 1);
                sAttributeNames[sAttributeNames.Length - 1] = "ForecastDate";
                sAttributeValues[sAttributeValues.Length - 1] = sForecastDate;
                sAttributeTypes[sAttributeTypes.Length - 1] = "string";
                //}

                //if (sActualDate != "")
                //{
                Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                Array.Resize<string>(ref sAttributeTypes, sAttributeTypes.Length + 1);
                sAttributeNames[sAttributeNames.Length - 1] = "ActualDate";
                sAttributeValues[sAttributeValues.Length - 1] = sActualDate;
                sAttributeTypes[sAttributeTypes.Length - 1] = "string";
                //}

                //if (sDateBasis != "")
                //{
                Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                Array.Resize<string>(ref sAttributeTypes, sAttributeTypes.Length + 1);
                sAttributeNames[sAttributeNames.Length - 1] = "DateBasis";
                sAttributeValues[sAttributeValues.Length - 1] = sDateBasis;
                sAttributeTypes[sAttributeTypes.Length - 1] = "string";
                //}

                //if (sComments != "")
                //{
                Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                Array.Resize<string>(ref sAttributeTypes, sAttributeTypes.Length + 1);
                sAttributeNames[sAttributeNames.Length - 1] = "Comments";
                sAttributeValues[sAttributeValues.Length - 1] = sComments;
                sAttributeTypes[sAttributeTypes.Length - 1] = "string";
                //}

                return client2.setdocattributes(sDocNo, sDocName, sAttributeNames, sAttributeValues, sAttributeTypes, sCheckInComments, Convert.ToInt16(sWebAppId));
            }
        }

        public string UpdatePlantEquipMaterialDoc(string sSessionId, string sUserId, string sDocNo, string sDocName,
                                          string sDesc, string sOriginator, string sOriginatorDocId,
                                          string sFirstIssueDate, string sIssueForUseDate, string sFinalIssueDate, string sStatusComments,
                                          string sComments, string sCheckInComments, string sWebAppId)
        {

            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                string[] sAttributeNames = new string[2];
                string[] sAttributeValues = new string[2];
                string[] sAttributeTypes = new string[2];

                sAttributeNames[0] = "description";
                sAttributeNames[1] = "Originator";


                sAttributeValues[0] = sDesc;
                sAttributeValues[1] = sOriginator;

                sAttributeTypes[0] = "string";
                sAttributeTypes[1] = "string";

                if (sOriginatorDocId != "")
                {
                    Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                    Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                    Array.Resize<string>(ref sAttributeTypes, sAttributeTypes.Length + 1);
                    sAttributeNames[sAttributeNames.Length - 1] = "OrigDocId";
                    sAttributeValues[sAttributeValues.Length - 1] = sOriginatorDocId;
                    sAttributeTypes[sAttributeTypes.Length - 1] = "string";
                }

                //if (sFirstIssueDate != "")
                //{
                Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                Array.Resize<string>(ref sAttributeTypes, sAttributeTypes.Length + 1);
                sAttributeNames[sAttributeNames.Length - 1] = "FirstIssueDate";
                sAttributeValues[sAttributeValues.Length - 1] = sFirstIssueDate;
                sAttributeTypes[sAttributeTypes.Length - 1] = "string";
                //}

                //if (sIssueForUseDate != "")
                //{
                Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                Array.Resize<string>(ref sAttributeTypes, sAttributeTypes.Length + 1);
                sAttributeNames[sAttributeNames.Length - 1] = "IssueForUseDate";
                sAttributeValues[sAttributeValues.Length - 1] = sIssueForUseDate;
                sAttributeTypes[sAttributeTypes.Length - 1] = "string";
                //}

                //if (sFinalIssueDate != "")
                //{
                Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                Array.Resize<string>(ref sAttributeTypes, sAttributeTypes.Length + 1);
                sAttributeNames[sAttributeNames.Length - 1] = "FinalIssueDate";
                sAttributeValues[sAttributeValues.Length - 1] = sFinalIssueDate;
                sAttributeTypes[sAttributeTypes.Length - 1] = "string";
                //}

                //if (sStatusComments != "")
                //{
                Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                Array.Resize<string>(ref sAttributeTypes, sAttributeTypes.Length + 1);
                sAttributeNames[sAttributeNames.Length - 1] = "StatusComments";
                sAttributeValues[sAttributeValues.Length - 1] = sStatusComments;
                sAttributeTypes[sAttributeTypes.Length - 1] = "string";
                //}

                //if (sComments != "")
                //{
                Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                Array.Resize<string>(ref sAttributeTypes, sAttributeTypes.Length + 1);
                sAttributeNames[sAttributeNames.Length - 1] = "Comments";
                sAttributeValues[sAttributeValues.Length - 1] = sComments;
                sAttributeTypes[sAttributeTypes.Length - 1] = "string";
                //}

                return client2.setdocattributes(sDocNo, sDocName, sAttributeNames, sAttributeValues, sAttributeTypes, sCheckInComments, Convert.ToInt16(sWebAppId));
            }
        }

        public string CreateWorkExecutionPackage(string sSessionId, string sUserId, string sWorkItemId, string sAssignedActivityId, string sRoute, string sPlannedWorkPackageNo, string sWEDName, string sProductName, string sDocType, string sFolderNameAndPath,
                                                 string sOriginator, string sJobCode, string sNew, string sExistingWEDNo, string sWebAppId, string sSkipCompleteTask)
        {

            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                int iiProdOrLibrary = 0;
                bool bNew = Convert.ToBoolean(sNew);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                string[] sAttributeNames = new string[2];
                string[] sAttributeValues = new string[2];
                string[] sAttributeTypes = new string[2];
                string sDocName;
                string sDocNo;
                string sCheckInComments;
                string sRtn2 = "";
                string[] sVariableNames = new string[0];
                string[] sVariableValues = new string[0];
                string[] sVariableTypes = new string[0];
                int iSkipCompleteTask = Convert.ToInt32(sSkipCompleteTask);


                if (sRoute.Contains("Terminate"))
                {
                    sRtn2 = client2.completetask(Convert.ToInt32(sWorkItemId), Convert.ToInt32(sAssignedActivityId), sRoute, sVariableNames, sVariableTypes, sVariableValues, Convert.ToInt16(sWebAppId));
                    if (sRtn2 != "Success")
                        return sRtn2;
                }
                else
                {
                    string sRtn1 = "Success";
                    if (bNew)
                    {
                        sAttributeNames[0] = "Originator";
                        sAttributeValues[0] = sOriginator;
                        sAttributeNames[1] = "JobCode";
                        sAttributeValues[1] = sJobCode;

                        sDocName = sWEDName;
                        sCheckInComments = "Auto creation of work execution package related to planned work package " + sPlannedWorkPackageNo;

                        sRtn1 = client2.doccreate2("", sDocName, sProductName, sDocType, sFolderNameAndPath, "A", sAttributeNames, sAttributeValues, sCheckInComments, iiProdOrLibrary, Convert.ToInt16(sWebAppId));
                    }

                    if (sRtn1.StartsWith("Success"))
                    {
                        string sSuccess = "";
                        if (bNew)
                        {
                            //Get the new document number
                            string[] sarrSuccess = Extract_Values(sRtn1);
                            sDocNo = sarrSuccess[1];
                            sSuccess = sarrSuccess[0];

                        }
                        else
                        {
                            sDocNo = sExistingWEDNo;
                            sSuccess = "Success";
                        }

                        if (sSuccess == "Success")
                        {
                            sCheckInComments = "Creating link between " + sDocNo + " and " + sPlannedWorkPackageNo;
                            sRtn2 = client2.setdoctopartdescribedby(sUserId, sDocNo, sPlannedWorkPackageNo, sCheckInComments, Convert.ToInt16(sWebAppId));

                            if (sRtn2 != "Success")
                                return sRtn2;
                            else
                            {
                                if (iSkipCompleteTask == 1)
                                    sRtn1 = "Success^" + sDocNo + "^";
                                else
                                {
                                    sRtn2 = client2.completetask(Convert.ToInt32(sWorkItemId), Convert.ToInt32(sAssignedActivityId), sRoute, sVariableNames, sVariableTypes, sVariableValues, Convert.ToInt16(sWebAppId));
                                    if (sRtn2 != "Success")
                                        return sRtn2;
                                    else
                                        sRtn1 = "Success^" + sDocNo + "^";
                                }
                            }
                        }
                    }

                    return sRtn1;
                }

                return "";

            }
        }

        public string SetDocPartDescribedByLink(string sSessionId, string sUserId, string sDocNo, string sPartNo, string sWebAppId)
        {

            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                string sCheckInComments;
                string sRtn2 = "";

                sCheckInComments = "Creating link between " + sDocNo + " and " + sPartNo;
                sRtn2 = client2.setdoctopartdescribedby(sUserId, sDocNo, sPartNo, sCheckInComments, Convert.ToInt16(sWebAppId));

                return sRtn2;
            }
        }

        public string CreateProjectMaterialItem(string sSessionId, string sUserId, string sFullName, string sPartNo, string sPartName,
                                                string sProductName, string sPartType,  string sFolderNameAndPath,
                                                string sCheckInComments, string sPartDescription, string sComments,
                                                string iProdOrLibrary, string sWebAppId)
        {

            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                int iiProdOrLibrary = Convert.ToInt16(iProdOrLibrary);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                string[] sAttributeNames = new string[3];
                string[] sAttributeValues = new string[3];
                string[] sAttributeTypes = new string[3];
                string sReturn = "";

                sAttributeNames[0] = "Originator";
                sAttributeNames[1] = "LongDescription";
                sAttributeNames[2] = "Comments";

                sAttributeValues[0] = sFullName;
                sAttributeValues[1] = sPartDescription;
                sAttributeValues[2] = sComments;

                sAttributeTypes[0] = "string";
                sAttributeTypes[1] = "string";
                sAttributeTypes[2] = "string";


                sReturn = client2.createpart(sPartNo, sPartName, sProductName, sPartType, sFolderNameAndPath, sFullName, sAttributeNames, sAttributeValues, sAttributeTypes, sCheckInComments, iiProdOrLibrary, Convert.ToInt16(sWebAppId));

                return sReturn;
            }
        }

        public string CreateFunctionalLocationBasePart(string sSessionId, string sUserId, string sFullName, string sPartNo, string sPartName,
                                                        string sProductName, string sPartType, string sFolderNameAndPath,
                                                        string sCheckInComments, string sPartDescription, string sComments,
                                                        string iProdOrLibrary, string sWebAppId)
        {

            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                int iiProdOrLibrary = Convert.ToInt16(iProdOrLibrary);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                string[] sAttributeNames = new string[3];
                string[] sAttributeValues = new string[3];
                string[] sAttributeTypes = new string[3];
                string sReturn = "";

                sAttributeNames[0] = "Originator";
                sAttributeNames[1] = "LongDescription";
                sAttributeNames[2] = "Comments";

                sAttributeValues[0] = sFullName;
                sAttributeValues[1] = sPartDescription;
                sAttributeValues[2] = sComments;

                sAttributeTypes[0] = "string";
                sAttributeTypes[1] = "string";
                sAttributeTypes[2] = "string";


                sReturn = client2.createpart(sPartNo, sPartName, sProductName, sPartType, sFolderNameAndPath, sFullName, sAttributeNames, sAttributeValues, sAttributeTypes, sCheckInComments, iiProdOrLibrary, Convert.ToInt16(sWebAppId));

                return sReturn;
            }
        }

        public string UpdateProjectMaterialItem(string sSessionId, string sUserId, string sFullName, string sPartNo, string sPartName,
                                                string sCheckInComments, string sPartDescription, string sComments,
                                                string sWebAppId)
        {

            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                string[] sAttributeNames = new string[3];
                string[] sAttributeValues = new string[3];
                string[] sAttributeTypes = new string[3];
                string sReturn = "";

                sAttributeNames[0] = "Originator";
                sAttributeNames[1] = "LongDescription";
                sAttributeNames[2] = "Comments";

                sAttributeValues[0] = sFullName;
                sAttributeValues[1] = sPartDescription;
                sAttributeValues[2] = sComments;

                sAttributeTypes[0] = "string";
                sAttributeTypes[1] = "string";
                sAttributeTypes[2] = "string";


                sReturn = client2.setpartattributes(sPartNo, sPartName, sFullName, sAttributeNames, sAttributeValues, sAttributeTypes, sCheckInComments, Convert.ToInt16(sWebAppId));

                return sReturn;
            }
        }

        public string CreateProjectWorkItem(string sSessionId, string sUserId, string sFullName, string sParentPartNo, string sPartNo, string sPartName,
                                            string sProductName, string sPartType, string sPartUsageType, string sPartUsageUnit, string sFolderNameAndPath,
                                            string sCheckInComments, string sLineNumber, string sPartDescription,
                                            string sReqirementsInfo, string sPreparationInfo, string sReviewInfo, string iProdOrLibrary, string sWebAppId)
        {

            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                int iiProdOrLibrary = Convert.ToInt16(iProdOrLibrary);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                string[] sAttributeNames = new string[2];
                string[] sAttributeValues = new string[2];
                string[] sAttributeTypes = new string[2];
                string[] sAttributeNamesLink = new string[0];
                string[] sAttributeValuesLink = new string[0];
                string[] sAttributeTypesLink = new string[0];
                string sReturn = "";
                string sReturn2 = "";

                sAttributeNames[0] = "Originator";
                sAttributeNames[1] = "PartDesc";

                sAttributeValues[0] = sFullName;
                sAttributeValues[1] = sPartDescription;

                sAttributeTypes[0] = "string";
                sAttributeTypes[1] = "string";

                if (sReqirementsInfo != "")
                {
                    Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                    Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                    Array.Resize<string>(ref sAttributeTypes, sAttributeTypes.Length + 1);
                    sAttributeNames[sAttributeNames.Length - 1] = "RequirementsInfo";
                    sAttributeValues[sAttributeValues.Length - 1] = sReqirementsInfo;
                    sAttributeTypes[sAttributeTypes.Length - 1] = "string";
                }

                if (sPreparationInfo != "")
                {
                    Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                    Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                    Array.Resize<string>(ref sAttributeTypes, sAttributeTypes.Length + 1);
                    sAttributeNames[sAttributeNames.Length - 1] = "PreparationInfo";
                    sAttributeValues[sAttributeValues.Length - 1] = sPreparationInfo;
                    sAttributeTypes[sAttributeTypes.Length - 1] = "string";
                }

                if (sReviewInfo != "")
                {
                    Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                    Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                    Array.Resize<string>(ref sAttributeTypes, sAttributeTypes.Length + 1);
                    sAttributeNames[sAttributeNames.Length - 1] = "ReviewInfo";
                    sAttributeValues[sAttributeValues.Length - 1] = sReviewInfo;
                    sAttributeTypes[sAttributeTypes.Length - 1] = "string";
                }

                sReturn = client2.createpart(sPartNo, sPartName, sProductName, sPartType, sFolderNameAndPath, sFullName, sAttributeNames, sAttributeValues, sAttributeTypes, sCheckInComments, iiProdOrLibrary, Convert.ToInt16(sWebAppId));
                if (sReturn.StartsWith("Success"))
                {
                    sCheckInComments = "Set link between project work item " + sParentPartNo + " and project work item " + sPartNo;
                    sReturn2 = client2.setpartpartlinkwithattributes(sFullName, sParentPartNo, sPartNo, 1, sCheckInComments, sPartUsageType, sPartUsageUnit, Convert.ToInt16(sLineNumber),
                                                                     sAttributeNamesLink, sAttributeValuesLink, sAttributeTypesLink, Convert.ToInt16(sWebAppId));
                    if (sReturn2 != "Success")
                        sReturn = sReturn2;
                }

                return sReturn;
            }
        }

        public string CreateProject(string sSessionId, string sUserId, string sFullName, string sPartNo, string sPartName,
                                            string sProductName, string sPartType, string sFolderNameAndPath,
                                            string sCheckInComments, string sPartDescription,
                                            string sReqirementsInfo, string sPreparationInfo, string sReviewInfo, string iProdOrLibrary, string sWebAppId)
        {

            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                int iiProdOrLibrary = Convert.ToInt16(iProdOrLibrary);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                string[] sAttributeNames = new string[2];
                string[] sAttributeValues = new string[2];
                string[] sAttributeTypes = new string[2];
                string sReturn = "";

                sAttributeNames[0] = "Originator";
                sAttributeNames[1] = "PartDesc";

                sAttributeValues[0] = sFullName;
                sAttributeValues[1] = sPartDescription;

                sAttributeTypes[0] = "string";
                sAttributeTypes[1] = "string";

                if (sReqirementsInfo != "")
                {
                    Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                    Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                    Array.Resize<string>(ref sAttributeTypes, sAttributeTypes.Length + 1);
                    sAttributeNames[sAttributeNames.Length - 1] = "RequirementsInfo";
                    sAttributeValues[sAttributeValues.Length - 1] = sReqirementsInfo;
                    sAttributeTypes[sAttributeTypes.Length - 1] = "string";
                }

                if (sPreparationInfo != "")
                {
                    Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                    Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                    Array.Resize<string>(ref sAttributeTypes, sAttributeTypes.Length + 1);
                    sAttributeNames[sAttributeNames.Length - 1] = "PreparationInfo";
                    sAttributeValues[sAttributeValues.Length - 1] = sPreparationInfo;
                    sAttributeTypes[sAttributeTypes.Length - 1] = "string";
                }

                if (sReviewInfo != "")
                {
                    Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                    Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                    Array.Resize<string>(ref sAttributeTypes, sAttributeTypes.Length + 1);
                    sAttributeNames[sAttributeNames.Length - 1] = "ReviewInfo";
                    sAttributeValues[sAttributeValues.Length - 1] = sReviewInfo;
                    sAttributeTypes[sAttributeTypes.Length - 1] = "string";
                }

                sReturn = client2.createpart(sPartNo, sPartName, sProductName, sPartType, sFolderNameAndPath, sFullName, sAttributeNames, sAttributeValues, sAttributeTypes, sCheckInComments, iiProdOrLibrary, Convert.ToInt16(sWebAppId));

                return sReturn;
            }
        }

        public string CreateParentChildPartLink(string sSessionId, string sUserId, string sFullName, string sParentPartNo, string sChildPartNo, string sQty,
                                               string sPartUsageType, string sPartUsageUnit,
                                               string sCheckInComments, string sLineNumber, string sWebAppId)
        {

            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                double dQty = Convert.ToDouble(sQty);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                string[] sAttributeNames = new string[1];
                string[] sAttributeValues = new string[1];
                string[] sAttributeTypes = new string[1];
                string[] sAttributeNamesLink = new string[0];
                string[] sAttributeValuesLink = new string[0];
                string[] sAttributeTypesLink = new string[0];
                string sReturn2 = "";

                sAttributeNames[0] = "Originator";
                sAttributeValues[0] = sFullName;
                sAttributeTypes[0] = "string";

                sReturn2 = client2.setpartpartlink(sFullName, sParentPartNo, sChildPartNo, dQty, sCheckInComments, sPartUsageType, sPartUsageUnit, Convert.ToInt16(sWebAppId));

                return sReturn2;
            }
        }

        public string InsertExistingProjectWorkItem(string sSessionId, string sUserId, string sFullName, string sParentPartNo, string sExistingPWIPartNo,
                                                    string sPartUsageType, string sPartUsageUnit, string sCheckInComments, string sLineNumber, string sWebAppId)
        {

            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                string[] sAttributeNamesLink = new string[0];
                string[] sAttributeValuesLink = new string[0];
                string[] sAttributeTypesLink = new string[0];
                string sReturn = "";

                sCheckInComments = "Set link between project work item " + sParentPartNo + " and project work item " + sExistingPWIPartNo;
                sReturn = client2.setpartpartlinkwithattributes(sFullName, sParentPartNo, sExistingPWIPartNo, 1, sCheckInComments, sPartUsageType, sPartUsageUnit, Convert.ToInt16(sLineNumber),
                                                                    sAttributeNamesLink, sAttributeValuesLink, sAttributeTypesLink, Convert.ToInt16(sWebAppId));

                return sReturn;
            }
        }

        public string CreateFronesisProject(string sSessionId, string sUserId, string sProjNo, string sProjDesc, string sProductName, string sDocType, string sPartType, string sFolderNameAndPath,
                                  string sClientDesc, string sOriginator, string sClientProjNo, string sRevision, string sCheckInComments, string iProdOrLibrary, string sWebAppId, string sProjType)
        {

            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ArrayList arrUser = GetUserDetails(sUserId);
                string sFullName = arrUser[2].ToString();

                int iiProdOrLibrary = Convert.ToInt16(iProdOrLibrary);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                string[] sAttributeNames = new string[3];
                string[] sAttributeValues = new string[3];
                string[] sAttributeTypes = new string[3];
                string sReturn = "";
                string sReturn2 = "";

                sAttributeNames[0] = "ClientId";
                sAttributeNames[1] = "ClientDesc";
                sAttributeNames[2] = "ProjectType";

                sAttributeValues[0] = sClientProjNo;
                sAttributeValues[1] = sClientDesc;
                sAttributeValues[2] = sProjType;

                sAttributeTypes[0] = "string";
                sAttributeTypes[1] = "string";
                sAttributeTypes[2] = "string";

                if (sOriginator != "")
                {
                    Array.Resize<string>(ref sAttributeNames, 4);
                    Array.Resize<string>(ref sAttributeValues, 4);
                    Array.Resize<string>(ref sAttributeTypes, 4);
                    sAttributeNames[3] = "Originator";
                    sAttributeValues[3] = sOriginator;
                    sAttributeTypes[3] = "string";
                }

                sReturn = client2.doccreate(sProjNo, sProjDesc, sProductName, sDocType, sFolderNameAndPath, sRevision, sAttributeNames, sAttributeValues, sCheckInComments, iiProdOrLibrary, Convert.ToInt16(sWebAppId));
                if (sReturn == "Success")
                {
                    sReturn = client2.createpart(sProjNo, sProjDesc, sProductName, sPartType, sFolderNameAndPath, sFullName, sAttributeNames, sAttributeValues, sAttributeTypes, sCheckInComments, iiProdOrLibrary, Convert.ToInt16(sWebAppId));
                    if (sReturn.StartsWith("Success"))
                    {
                        sCheckInComments = "Set referenced by link between project document and project part with the identical number " + sProjNo;
                        sReturn2 = client2.setdoctopartref(sOriginator, sProjNo, sProjNo, sCheckInComments, "wt.part.WTPartReferenceLink", Convert.ToInt16(sWebAppId));
                        if (sReturn2 != "Success")
                            sReturn = sReturn2;
                    }
                }

                return sReturn;
            }
        }

        public string CreateFronesisProjectChildDoc(string sSessionId, string sUserId, string sProjNo, string sChildDocNo, string sChildDocName, string sProductName, string sDocType, string sFolderNameAndPath,
                                                    string sOriginator, string sRevision, string sCheckInComments, string iProdOrLibrary, string sWebAppId)
        {

            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                int iiProdOrLibrary = Convert.ToInt16(iProdOrLibrary);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                string[] sAttributeNames = new string[1];
                string[] sAttributeValues = new string[1];
                string[] sAttributeTypes = new string[1];
                string sReturn = "";
                string sReturn2 = "";

                if (sOriginator != "")
                {
                    sAttributeNames[0] = "Originator";
                    sAttributeValues[0] = sOriginator;
                    sAttributeTypes[0] = "string";
                }

                sReturn = client2.doccreate(sChildDocNo, sChildDocName, sProductName, sDocType, sFolderNameAndPath, sRevision, sAttributeNames, sAttributeValues, sCheckInComments, iiProdOrLibrary, Convert.ToInt16(sWebAppId));
                if (sReturn == "Success")
                {
                    sCheckInComments = "Set link between project document " + sProjNo + " and child document " + sChildDocNo;
                    sReturn2 = client2.setdocdoclink(sOriginator, sProjNo, sChildDocNo, sCheckInComments, "wt.doc.WTDocumentUsageLink", Convert.ToInt16(sWebAppId));
                    if (sReturn2 != "Success")
                        sReturn = sReturn2;
                }

                return sReturn;
            }
        }

        public string AttachWCDoc(string sSessionId, string sUserId, string sFullName, string sDocNo, string sAttachDesc, string sAttachPath, string bSecondary, string sAttachComments, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                bool bbSecondary = Convert.ToBoolean(bSecondary);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                sAttachPath = sAttachPath.Replace("||", "/");
                return client2.attachdoc(sFullName, sDocNo, sAttachDesc, sAttachPath, bbSecondary, sAttachComments, Convert.ToInt16(sWebAppId));
            }
        }

        public string AttachURL(string sSessionId, string sUserId, string sFullName, string sDocNo, string sURLDesc, string sURL, string bSecondary, string sAttachComments, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                bool bbSecondary = Convert.ToBoolean(bSecondary);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                return client2.attachurl(sFullName, sDocNo, sURLDesc, sURLDesc, sURL, bbSecondary, sAttachComments, Convert.ToInt16(sWebAppId));
            }
        }

        public string DeleteWCDoc(string sSessionId, string sUserId, string sFullName, string sDocNo, string sAttachFileName, string bSecondary, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                bool bbSecondary = Convert.ToBoolean(bSecondary);
                ExampleService.MyJavaService3Client client2 = GetWCService();

                return client2.deleteattachment(sFullName, sDocNo, sAttachFileName, bbSecondary, Convert.ToInt16(sWebAppId));
            }
        }

        public string DeleteURL(string sSessionId, string sUserId, string sFullName, string sDocNo, string sURL, string bSecondary, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                bool bbSecondary = Convert.ToBoolean(bSecondary);
                ExampleService.MyJavaService3Client client2 = GetWCService();

                return client2.deleteurl(sFullName, sDocNo, sURL, bbSecondary, Convert.ToInt16(sWebAppId));
            }
        }

        public string SetDocAttributeStrings(string sSessionId, string sUserId, string sDocNo, string sDocName, string sLongDesc, string sOriginator, string sOriginatorDocId, string sJobCode, string sCheckInComments, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                string[] sAttributeNames = new string[3];
                string[] sAttributeValues = new string[3];

                sAttributeNames[0] = "LongDescription";
                sAttributeNames[1] = "Originator";
                sAttributeNames[2] = "OrigDocId";

                sAttributeValues[0] = sLongDesc;
                sAttributeValues[1] = sOriginator;
                sAttributeValues[2] = sOriginatorDocId;

                if (sJobCode != "")
                {
                    Array.Resize<string>(ref sAttributeNames, 4);
                    Array.Resize<string>(ref sAttributeValues, 4);
                    sAttributeNames[3] = "JobCode";
                    sAttributeValues[3] = sJobCode;
                }

                return client2.setdocattributestrings(sDocNo, sDocName, sAttributeNames, sAttributeValues, sCheckInComments, Convert.ToInt16(sWebAppId));
            }
        }

        public string SetDocAttributeStrings2(string sSessionId, string sUserId, string sDocNo, string sDocName, string sDesc, string sOriginator, string sOriginatorDocId, string sJobCode, string sCheckInComments, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                string[] sAttributeNames = new string[3];
                string[] sAttributeValues = new string[3];

                sAttributeNames[0] = "description";
                sAttributeNames[1] = "Originator";
                sAttributeNames[2] = "OrigDocId";

                sAttributeValues[0] = sDesc;
                sAttributeValues[1] = sOriginator;
                sAttributeValues[2] = sOriginatorDocId;

                if (sJobCode != "")
                {
                    Array.Resize<string>(ref sAttributeNames, 4);
                    Array.Resize<string>(ref sAttributeValues, 4);
                    sAttributeNames[3] = "JobCode";
                    sAttributeValues[3] = sJobCode;
                }

                return client2.setdocattributestrings(sDocNo, sDocName, sAttributeNames, sAttributeValues, sCheckInComments, Convert.ToInt16(sWebAppId));
            }
        }

        public string SetDocToDocRef(string sSessionId, string sUserId, string sFullName, string sDocNo, string sReferencedDocNo, string sCheckinComments, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                return client2.setdoctodocref(sFullName, sDocNo, sReferencedDocNo, sCheckinComments, Convert.ToInt16(sWebAppId));
            }
        }

        public string SetDocToDocRefs(string sSessionId, string sUserId, string sFullName, string sDocNo, string sReferencedDocNos, string sCheckinComments, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                char[] charSeparators = new char[] { '^' };
                string[] sReferencedDocNos2 = sReferencedDocNos.Split(charSeparators, StringSplitOptions.RemoveEmptyEntries);

                return client2.setdoctodocrefs(sFullName, sDocNo, sReferencedDocNos2, sCheckinComments, Convert.ToInt16(sWebAppId));
            }
        }

        public string SetDocReviewer(string sSessionId, string sUserId, string sFullName, string sDocNo, string sReviewerNo, string sCheckinComments, string sReviewerTypeName, string sCompletionDate, string sCompletionStatus, string sComments, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                string[] sAttributeNames = new string[2];
                string[] sAttributeValues = new string[2];
                string[] sAttributeTypes = new string[2];

                sAttributeNames[0] = "ReviewerTypeName";
                sAttributeNames[1] = "CompletionStatus";

                sAttributeValues[0] = sReviewerTypeName;
                sAttributeValues[1] = sCompletionStatus;

                sAttributeTypes[0] = "string";
                sAttributeTypes[1] = "long";

                if (sCompletionDate != "")
                {
                    Array.Resize<string>(ref sAttributeNames, 3);
                    Array.Resize<string>(ref sAttributeValues, 3);
                    Array.Resize<string>(ref sAttributeTypes, 3);
                    sAttributeNames[2] = "CompletedDate";
                    sAttributeValues[2] = sCompletionDate;
                    sAttributeTypes[2] = "date";

                    if (sComments != "")
                    {
                        Array.Resize<string>(ref sAttributeNames, 4);
                        Array.Resize<string>(ref sAttributeValues, 4);
                        Array.Resize<string>(ref sAttributeTypes, 4);
                        sAttributeNames[3] = "Comments";
                        sAttributeValues[3] = sComments;
                        sAttributeTypes[3] = "string";
                    }
                }


                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                return client2.setdocdoclinkwithattributes(sFullName, sDocNo, sReviewerNo, sCheckinComments, "local.rs.vsrs05.Regain.ReviewerDocLink", sAttributeNames, sAttributeValues, sAttributeTypes, Convert.ToInt16(sWebAppId));
            }
        }

        public string UpdateDocReviewer(string sSessionId, string sUserId, string sFullName, string sDocNo, string sReviewerNo, string sCheckinComments, string sReviewerTypeName, string sCompletionDate, string sCompletionStatus, string sComments, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                string[] sAttributeNames = new string[2];
                string[] sAttributeValues = new string[2];
                string[] sAttributeTypes = new string[2];

                sAttributeNames[0] = "ReviewerTypeName";
                sAttributeNames[1] = "CompletionStatus";

                sAttributeValues[0] = sReviewerTypeName;
                sAttributeValues[1] = sCompletionStatus;

                sAttributeTypes[0] = "string";
                sAttributeTypes[1] = "long";

                if (sCompletionDate != "")
                {
                    Array.Resize<string>(ref sAttributeNames, 3);
                    Array.Resize<string>(ref sAttributeValues, 3);
                    Array.Resize<string>(ref sAttributeTypes, 3);
                    sAttributeNames[2] = "CompletedDate";
                    sAttributeValues[2] = sCompletionDate;
                    sAttributeTypes[2] = "date";

                    if (sComments != "")
                    {
                        Array.Resize<string>(ref sAttributeNames, 4);
                        Array.Resize<string>(ref sAttributeValues, 4);
                        Array.Resize<string>(ref sAttributeTypes, 4);
                        sAttributeNames[3] = "Comments";
                        sAttributeValues[3] = sComments;
                        sAttributeTypes[3] = "string";
                    }
                }

                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                return client2.updatedocdoclinkwithattributes(sFullName, sDocNo, sReviewerNo, sCheckinComments, "local.rs.vsrs05.Regain.ReviewerDocLink", sAttributeNames, sAttributeValues, sAttributeTypes, Convert.ToInt16(sWebAppId));
            }
        }

        public string SetPartReviewer(string sSessionId, string sUserId, string sFullName, string sPartNo, string sReviewerNo, string sPartRefLinkType, string sCheckinComments,
                                      string sReviewerTypeName, string sCompletionDate, string sCompletionStatus, string sComments, string sAccountableFlag, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                string[] sAttributeNames = new string[2];
                string[] sAttributeValues = new string[2];
                string[] sAttributeTypes = new string[2];

                sAttributeNames[0] = "ReviewerTypeName";
                sAttributeNames[1] = "CompletionStatus";
                sAttributeValues[0] = sReviewerTypeName;
                sAttributeValues[1] = sCompletionStatus;

                sAttributeTypes[0] = "string";
                sAttributeTypes[1] = "long";

                if (sCompletionDate != "")
                {
                    Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                    Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                    Array.Resize<string>(ref sAttributeTypes, sAttributeTypes.Length + 1);
                    sAttributeNames[2] = "CompletedDate";
                    sAttributeValues[2] = sCompletionDate;
                    sAttributeTypes[2] = "date";
                }


                if (sComments != "")
                {
                    Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                    Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                    Array.Resize<string>(ref sAttributeTypes, sAttributeTypes.Length + 1);
                    sAttributeNames[3] = "Comments";
                    sAttributeValues[3] = sComments;
                    sAttributeTypes[3] = "string";
                }

                if (sAccountableFlag != "")
                {
                    Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                    Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                    Array.Resize<string>(ref sAttributeTypes, sAttributeTypes.Length + 1);
                    sAttributeNames[sAttributeNames.Length - 1] = "AccountableFlag";
                    sAttributeValues[sAttributeValues.Length - 1] = sAccountableFlag;
                    sAttributeTypes[sAttributeTypes.Length - 1] = "long";
                }



                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                return client2.setpartreferencedbydoclinkwithattributes(sFullName, sReviewerNo, sPartNo, sAttributeNames, sAttributeValues, sAttributeTypes, sCheckinComments, sPartRefLinkType, Convert.ToInt16(sWebAppId));
            }
        }

        public string UpdatePartReviewer(string sSessionId, string sUserId, string sFullName, string sPartNo, string sReviewerNo, string sCheckinComments,
                                         string sReviewerTypeName, string sCompletionDate, string sCompletionStatus, string sComments, string sAccountableFlag, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                string[] sAttributeNames = new string[2];
                string[] sAttributeValues = new string[2];
                string[] sAttributeTypes = new string[2];

                sAttributeNames[0] = "ReviewerTypeName";
                sAttributeNames[1] = "CompletionStatus";

                sAttributeValues[0] = sReviewerTypeName;
                sAttributeValues[1] = sCompletionStatus;

                sAttributeTypes[0] = "string";
                sAttributeTypes[1] = "long";

                if (sCompletionDate != "")
                {
                    Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                    Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                    Array.Resize<string>(ref sAttributeTypes, sAttributeTypes.Length + 1);
                    sAttributeNames[2] = "CompletedDate";
                    sAttributeValues[2] = sCompletionDate;
                    sAttributeTypes[2] = "date";
                }


                Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                Array.Resize<string>(ref sAttributeTypes, sAttributeTypes.Length + 1);
                sAttributeNames[3] = "Comments";
                sAttributeValues[3] = sComments;
                sAttributeTypes[3] = "string";

                if (sAccountableFlag != "")
                {
                    Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                    Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                    Array.Resize<string>(ref sAttributeTypes, sAttributeTypes.Length + 1);
                    sAttributeNames[sAttributeNames.Length - 1] = "AccountableFlag";
                    sAttributeValues[sAttributeValues.Length - 1] = sAccountableFlag;
                    sAttributeTypes[sAttributeTypes.Length - 1] = "long";
                }

                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                return client2.updatepartreferencedbydoclinkwithattributes(sFullName, sReviewerNo, sPartNo, sAttributeNames, sAttributeValues, sAttributeTypes, sCheckinComments, Convert.ToInt16(sWebAppId));
            }
        }

        public string DeleteDocToDocUsageLink(string sSessionId, string sUserId, string sFullName, string sParentDocNo, string sChildDocNo, string sCheckinComments, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                return client2.deletedoctodocusagelink(sFullName, sParentDocNo, sChildDocNo, sCheckinComments, Convert.ToInt16(sWebAppId));
            }
        }

        public string DeleteDocToDocRef(string sSessionId, string sUserId, string sFullName, string sDocNo, string sReferencedDocNo, string sCheckinComments, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                return client2.deletedoctodocref(sFullName, sDocNo, sReferencedDocNo, sCheckinComments, Convert.ToInt16(sWebAppId));
            }
        }

        public string DeleteDocToDocRefs(string sSessionId, string sUserId, string sFullName, string sDocNo, string sReferencedDocNos, string sCheckinComments, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                char[] charSeparators = new char[] { '^' };
                string[] sReferencedDocNos2 = sReferencedDocNos.Split(charSeparators, StringSplitOptions.RemoveEmptyEntries);

                return client2.deletedoctodocrefs(sFullName, sDocNo, sReferencedDocNos2, sCheckinComments, Convert.ToInt16(sWebAppId));
            }
        }

        public string SetDocToPartRef(string sSessionId, string sUserId, string sFullName, string sDocNo, string sPartNo, string sCheckinComments, string sPartRefLinkType, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                return client2.setdoctopartref(sFullName, sDocNo, sPartNo, sCheckinComments, sPartRefLinkType, Convert.ToInt16(sWebAppId));
            }
        }

        public string SetDeliverableDocToPartRef(string sSessionId, string sUserId, string sFullName, string sDocNo, string sPartNo, string sCheckinComments, string sPartRefLinkType, string sWebAppId)
        {
            string[] sAttributeNames = new string[1];
            string[] sAttributeValues = new string[1];
            string[] sAttributeTypes = new string[1];

            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                sAttributeNames[0] = "Deliverable";

                sAttributeValues[0] = "1";

                sAttributeTypes[0] = "long";

                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                return client2.setpartreferencedbydoclinkwithattributes(sFullName, sDocNo, sPartNo, sAttributeNames, sAttributeValues, sAttributeTypes, sCheckinComments, sPartRefLinkType, Convert.ToInt16(sWebAppId));
            }
        }

        public string SetDeliverablePartToPartLink(string sSessionId, string sUserId, string sFullName, string sParentPart, string sChildPart, string sCheckinComments, string sPartUsageLinkType, string sLineNumber, string sWebAppId)
        {
            string[] sAttributeNames = new string[1];
            string[] sAttributeValues = new string[1];
            string[] sAttributeTypes = new string[1];
            long lLineNumber = 1;

            lLineNumber = Convert.ToInt64(sLineNumber);

            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                sAttributeNames[0] = "Deliverable";

                sAttributeValues[0] = "1";

                sAttributeTypes[0] = "long";

                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();

                return client2.setpartpartlinkwithattributes(sFullName, sParentPart, sChildPart, 1.0, sCheckinComments, sPartUsageLinkType, "ea", lLineNumber, sAttributeNames, sAttributeValues, sAttributeTypes, Convert.ToInt16(sWebAppId));
            }
        }

        public string SetFuncDocToPartRef(string sSessionId, string sUserId, string sFullName, string sFuncDocNo, string sPartNo, string sSequenceNo, string sPrimaryPart, string sPartDocRefLinkType, string sCheckinComments, string sWebAppId)
        {
            string[] sAttributeNames = new string[2];
            string[] sAttributeValues = new string[2];
            string[] sAttributeTypes = new string[2];


            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();

                sAttributeNames[0] = "SequenceNo";
                sAttributeNames[1] = "PrimaryPart";

                sAttributeValues[0] = sSequenceNo;
                sAttributeValues[1] = sPrimaryPart;

                sAttributeTypes[0] = "long";
                sAttributeTypes[1] = "bool";

                return client2.setpartreferencedbydoclinkwithattributes(sFullName, sFuncDocNo, sPartNo, sAttributeNames, sAttributeValues, sAttributeTypes, sCheckinComments, sPartDocRefLinkType, Convert.ToInt16(sWebAppId));
            }
        }

        //THis uses the special part reference type that takes attributes. Attributes for relationships cannot be updated in native Windchill with the exception of the Part to Part link.
        public string UpdateFuncDocToPartRef(string sSessionId, string sUserId, string sFullName, string sFuncDocNo, string sPartNo, string sSequenceNo, string sPrimaryPart, string sCheckinComments, string sWebAppId)
        {
            string[] sAttributeNames = new string[2];
            string[] sAttributeValues = new string[2];
            string[] sAttributeTypes = new string[2];


            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();

                sAttributeNames[0] = "SequenceNo";
                sAttributeNames[1] = "PrimaryPart";

                sAttributeValues[0] = sSequenceNo;

                if (sPrimaryPart == "1" || sPrimaryPart == "Yes" || sPrimaryPart == "Y")
                    sPrimaryPart = "true";
                else
                    sPrimaryPart = "false";

                sAttributeValues[1] = sPrimaryPart;

                sAttributeTypes[0] = "long";
                sAttributeTypes[1] = "bool";

                return client2.updatepartreferencedbydoclinkwithattributes(sFullName, sFuncDocNo, sPartNo, sAttributeNames, sAttributeValues, sAttributeTypes, sCheckinComments, Convert.ToInt16(sWebAppId));
            }
        }

        public string SetSupplierToPartRef(string sSessionId, string sUserId, string sFullName, string sSupplierNo, string sPartNo, string sSupplierPartNo, string sPartDocRefLinkType, string sCheckinComments, string sWebAppId, string sManufacturerFlag)
        {
            string[] sAttributeNames = new string[2];
            string[] sAttributeValues = new string[2];
            string[] sAttributeTypes = new string[2];


            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();

                sAttributeNames[0] = "SupplierPartNo";
                sAttributeValues[0] = sSupplierPartNo;
                sAttributeTypes[0] = "string";

                sAttributeNames[1] = "ManufacturerFlag";
                sAttributeValues[1] = sManufacturerFlag;
                sAttributeTypes[1] = "long";

                return client2.setpartreferencedbydoclinkwithattributes(sFullName, sSupplierNo, sPartNo, sAttributeNames, sAttributeValues, sAttributeTypes, sCheckinComments, sPartDocRefLinkType, Convert.ToInt16(sWebAppId));
            }
        }

        //THis uses the special part reference type that takes attributes. Attributes for relationships cannot be updated in native Windchill with the exception of the Part to Part link.
        public string UpdateSupplierToPartRef(string sSessionId, string sUserId, string sFullName, string sSupplierNo, string sPartNo, string sSupplierPartNo, string sCheckinComments, string sWebAppId, string sManufacturerFlag)
        {
            string[] sAttributeNames = new string[2];
            string[] sAttributeValues = new string[2];
            string[] sAttributeTypes = new string[2];


            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();

                sAttributeNames[0] = "SupplierPartNo";
                sAttributeValues[0] = sSupplierPartNo;
                sAttributeTypes[0] = "string";

                sAttributeNames[1] = "ManufacturerFlag";
                sAttributeValues[1] = sManufacturerFlag;
                sAttributeTypes[1] = "long";

                return client2.updatepartreferencedbydoclinkwithattributes(sFullName, sSupplierNo, sPartNo, sAttributeNames, sAttributeValues, sAttributeTypes, sCheckinComments, Convert.ToInt16(sWebAppId));
            }
        }

        public string SetDocToPartRefs(string sSessionId, string sUserId, string sFullName, string sDocNo, string sPartNos, string sCheckinComments, string sPartDocRefType, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                char[] charSeparators = new char[] { '^' };
                string[] sPartNos2 = sPartNos.Split(charSeparators, StringSplitOptions.RemoveEmptyEntries);

                return client2.setdoctopartrefs(sFullName, sDocNo, sPartNos2, sCheckinComments, sPartDocRefType, Convert.ToInt16(sWebAppId));
            }
        }

        public string DeleteDocToPartRef(string sSessionId, string sUserId, string sFullName, string sDocNo, string sPartNo, string sCheckinComments, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                return client2.deletedoctopartref(sFullName, sDocNo, sPartNo, sCheckinComments, Convert.ToInt16(sWebAppId));
            }
        }

        public string DeleteDocToPartRefs(string sSessionId, string sUserId, string sFullName, string sDocNo, string sPartNos, string sCheckinComments, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                char[] charSeparators = new char[] { '^' };
                string[] sPartNos2 = sPartNos.Split(charSeparators, StringSplitOptions.RemoveEmptyEntries);

                return client2.deletedoctopartrefs(sFullName, sDocNo, sPartNos2, sCheckinComments, Convert.ToInt16(sWebAppId));
            }
        }

        public string DeleteDocToPartRefWithAttribute(string sSessionId, string sUserId, string sFullName, string sDocNo, string sPartNo, string sAttributeName, string sAttributeValue, string sCheckinComments, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                return client2.deletedoctopartrefwithattribute(sFullName, sDocNo, sPartNo, sAttributeName, sAttributeValue, sCheckinComments, Convert.ToInt16(sWebAppId));
            }
        }

        public string DeleteDocToPartDescribeBy(string sSessionId, string sUserId, string sFullName, string sDocNo, string sPartNo, string sCheckinComments, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                return client2.deletedoctopartdescribeby(sFullName, sDocNo, sPartNo, sCheckinComments, Convert.ToInt16(sWebAppId));
            }
        }

        public string DeleteDocToPartDescribeBys(string sSessionId, string sUserId, string sFullName, string sDocNo, string sPartNos, string sCheckinComments, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                char[] charSeparators = new char[] { '^' };
                string[] sPartNos2 = sPartNos.Split(charSeparators, StringSplitOptions.RemoveEmptyEntries);

                return client2.deletedoctopartdescribebys(sFullName, sDocNo, sPartNos2, sCheckinComments, Convert.ToInt16(sWebAppId));
            }
        }

        public string UpdateActionRequest(string sSessionId, string sUserId, string sFullName, string sARCode, string sARName, string sARCategory, string sARCause, string sARComments,
                                          string sARLongDesc, string sARDate, string sRequestActionType, string sCheckinComments, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                string[] sAttributeNames = new string[6];
                string[] sAttributeValues = new string[6];
                string[] sAttributeTypes = new string[6];

                sAttributeNames[0] = "Originator";
                sAttributeNames[1] = "ActionCategory";
                sAttributeNames[2] = "ARCause";
                sAttributeNames[3] = "LongDescription";
                sAttributeNames[4] = "Comments";
                sAttributeNames[5] = "RequestDate";

                sAttributeValues[0] = sFullName;
                sAttributeValues[1] = sARCategory;
                sAttributeValues[2] = sARCause;
                sAttributeValues[3] = sARLongDesc;
                sAttributeValues[4] = sARComments;
                sAttributeValues[5] = sARDate;

                sAttributeTypes[0] = "string";
                sAttributeTypes[1] = "string";
                sAttributeTypes[2] = "string";
                sAttributeTypes[3] = "string";
                sAttributeTypes[4] = "string";
                sAttributeTypes[5] = "datetime";

                if (sRequestActionType != "")
                {
                    Array.Resize<string>(ref sAttributeNames, 7);
                    Array.Resize<string>(ref sAttributeValues, 7);
                    Array.Resize<string>(ref sAttributeTypes, 7);
                    sAttributeNames[6] = "RequestedActionType";
                    sAttributeValues[6] = sRequestActionType;
                    sAttributeTypes[6] = "string";
                }

                return client2.setpartattributes(sARCode, sARName, sFullName, sAttributeNames, sAttributeValues, sAttributeTypes, sCheckinComments, Convert.ToInt16(sWebAppId));
            }
        }

        public string UpdateProjectStatus(string sSessionId, string sUserId, string sFullName, string sProjCode, string sProjName, string sProjStatus, string sCheckinComments, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                string[] sAttributeNames = new string[1];
                string[] sAttributeValues = new string[1];
                string[] sAttributeTypes = new string[1];

                sAttributeNames[0] = "ProjectStatus";
                sAttributeValues[0] = sProjStatus;
                sAttributeTypes[0] = "string";

                return client2.setpartattributes(sProjCode, sProjName, sFullName, sAttributeNames, sAttributeValues, sAttributeTypes, sCheckinComments, Convert.ToInt16(sWebAppId));
            }
        }

        public string SetPartToPartLink(string sSessionId, string sUserId, string sFullName, string sParentPartNo, string sChildPartNumber, string dQty, string sCheckInComments, string sPartUsageType, string sUnit, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                double ddQty = Convert.ToDouble(dQty);
                ExampleService.MyJavaService3Client client2 = GetWCService();

                return client2.setpartpartlink(sFullName, sParentPartNo, sChildPartNumber, ddQty, sCheckInComments, sPartUsageType, sUnit, Convert.ToInt16(sWebAppId));
            }
        }

        public string SetPartUsageLinkQty(string sSessionId, string sUserId, string sParentPartNo, string sChildPartNo, string dQty, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ArrayList arrUser = GetUserDetails(sUserId);

                string sFullName = arrUser[2].ToString();
                double ddQty = Convert.ToDouble(dQty);
                ExampleService.MyJavaService3Client client2 = GetWCService();

                return client2.setpartusagelinkqty(sParentPartNo, sChildPartNo, sFullName, ddQty, Convert.ToInt16(sWebAppId));
            }
        }


        public string SetMBATransaction(string sSessionId, string sUserId, string sFullName, string sParentPartNo, string sChildPartNumber, string dQty,
                                        long lLineNumber, string sDDno, string sDDDate, string sComments, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                double ddQty = Convert.ToDouble(dQty);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                string sCheckInComments = "Created usage link between parent " + sParentPartNo + " and child " + sChildPartNumber;
                string[] sAttributeNames = new string[6];
                string[] sAttributeValues = new string[6];
                string[] sAttributeTypes = new string[6];

                sAttributeNames[0] = "Originator";
                sAttributeNames[1] = "DispatchDocketNo";
                sAttributeNames[2] = "TransactionDate";
                sAttributeNames[3] = "UsageComments";

                sAttributeValues[0] = sFullName;
                sAttributeValues[1] = sDDno;
                sAttributeValues[2] = sDDDate;
                sAttributeValues[3] = sComments;

                sAttributeTypes[0] = "string";
                sAttributeTypes[1] = "string";
                sAttributeTypes[2] = "datetime";
                sAttributeTypes[3] = "string";

                return client2.setpartpartlinkwithattributes(sFullName, sParentPartNo, sChildPartNumber, ddQty, sCheckInComments, "local.rs.vsrs05.Regain.MBAUsageLink", "tonne",
                                                             lLineNumber, sAttributeNames, sAttributeValues, sAttributeTypes, Convert.ToInt16(sWebAppId));
            }
        }

        public string UpdateMBATransaction(string sSessionId, string sUserId, string sFullName, string sParentPartNo, string sChildPartNumber,
                                           string dQty, long lOldLineNumber, long lNewLineNumber, string sDDno, string sDDDate, string sComments, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                double ddQty = Convert.ToDouble(dQty);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                string sCheckInComments = "Created usage link between parent " + sParentPartNo + " and child " + sChildPartNumber;
                string[] sAttributeNames = new string[6];
                string[] sAttributeValues = new string[6];
                string[] sAttributeTypes = new string[6];

                sAttributeNames[0] = "Originator";
                sAttributeNames[1] = "DispatchDocketNo";
                sAttributeNames[2] = "TransactionDate";
                sAttributeNames[3] = "UsageComments";

                sAttributeValues[0] = sFullName;
                sAttributeValues[1] = sDDno;
                sAttributeValues[2] = sDDDate;
                sAttributeValues[3] = sComments;

                sAttributeTypes[0] = "string";
                sAttributeTypes[1] = "string";
                sAttributeTypes[2] = "datetime";
                sAttributeTypes[3] = "string";

                return client2.updatedispatchdocketpartpartlinkwithattributes(sFullName, sParentPartNo, sChildPartNumber, ddQty, sDDno, sCheckInComments, "local.rs.vsrs05.Regain.MBAUsageLink",
                                                                              "tonne", lOldLineNumber, lNewLineNumber, sAttributeNames, sAttributeValues, sAttributeTypes, Convert.ToInt16(sWebAppId));
            }
        }

        public string UpdateMBATransactionInvoiceStatus(string sSessionId, string sUserId, string sFullName, string sParentPartNo, string sChildPartNo,
                                                        string sLineNumber, string sInvoiceStatus, string sInvoiceNo, string sBatchList, string sCutoffDate, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                long lLineNumber = Convert.ToInt64(sLineNumber);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                string sCheckInComments = "Changed invoice status between parent " + sParentPartNo + " and child " + sChildPartNo + " on line number " + sLineNumber;
                if (sInvoiceNo == "")
                {
                    string[] sAttributeNames = new string[2];
                    string[] sAttributeValues = new string[2];
                    string[] sAttributeTypes = new string[2];

                    sAttributeNames[0] = "Originator";
                    sAttributeNames[1] = "InvoiceStatus";

                    sAttributeValues[0] = sFullName;
                    sAttributeValues[1] = sInvoiceStatus;

                    sAttributeTypes[0] = "string";
                    sAttributeTypes[1] = "long";

                    return client2.setpartusageattributesfromlinenumber(sParentPartNo, sChildPartNo, sFullName, lLineNumber, sAttributeNames, sAttributeValues, sAttributeTypes, Convert.ToInt16(sWebAppId));

                }
                else
                {
                    string[] sAttributeNames = new string[5];
                    string[] sAttributeValues = new string[5];
                    string[] sAttributeTypes = new string[5];

                    sAttributeNames[0] = "Originator";
                    sAttributeNames[1] = "InvoiceStatus";
                    sAttributeNames[2] = "InvoiceNo";
                    sAttributeNames[3] = "BatchList";
                    sAttributeNames[4] = "CutoffDate";

                    sAttributeValues[0] = sFullName;
                    sAttributeValues[1] = sInvoiceStatus;
                    sAttributeValues[2] = sInvoiceNo;
                    sAttributeValues[3] = sBatchList;
                    sAttributeValues[4] = sCutoffDate;

                    sAttributeTypes[0] = "string";
                    sAttributeTypes[1] = "long";
                    sAttributeTypes[2] = "string";
                    sAttributeTypes[3] = "string";
                    sAttributeTypes[4] = "string";

                    return client2.setpartusageattributesfromlinenumber(sParentPartNo, sChildPartNo, sFullName, lLineNumber, sAttributeNames, sAttributeValues, sAttributeTypes, Convert.ToInt16(sWebAppId));

                }
            }
        }

        public string UpdateMBAMultipleTransactionInvoiceStatus(string sSessionId, string sUserId, string sFullName, string sParentPartNo, string sChildPartNo,
                                                                string sLineNumber, string sInvoiceStatus, string sQtyInvoiced, string sWebAppId)
        {
            int i;
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                //long[] lLineNumber = Array.ConvertAll(sLineNumber.Split(','), long.Parse);

                //long?[] lLineNumbers = new long?[lLineNumber.Length];

                //for (i = 0; i < lLineNumber.Length; i++)
                //    lLineNumbers[i] = lLineNumber[i];

                string[] sChildParts = sChildPartNo.Split(',');
                string[] sInvoiceStatuses = sInvoiceStatus.Split(',');
                string[] sQtysInvoiced = sQtyInvoiced.Split(',');
                ExampleService.MyJavaService3Client client2 = GetWCService();
                string sCheckInComments = "Changed invoice status between parent " + sParentPartNo + " and child " + sChildPartNo + " on line number " + sLineNumber;
                string[] sAttributeNames = new string[sChildParts.Length * 3];
                string[] sAttributeValues = new string[sChildParts.Length * 3];
                string[] sAttributeTypes = new string[sChildParts.Length * 3];

                for (i = 0; i < sInvoiceStatuses.Length; i++)
                {
                    sAttributeNames[i * 3] = "Originator";
                    sAttributeNames[i * 3 + 1] = "InvoiceStatus";
                    sAttributeNames[i * 3 + 2] = "QtyInvoiced";

                    sAttributeValues[i * 3] = sFullName;
                    sAttributeValues[i * 3 + 1] = sInvoiceStatuses[i];
                    sAttributeValues[i * 3 + 2] = sQtysInvoiced[i];

                    sAttributeTypes[i * 3] = "string";
                    sAttributeTypes[i * 3 + 1] = "long";
                    sAttributeTypes[i * 3 + 2] = "double";
                }

                //Set it up to split at the Windchill end because for some reason sending through nullable long arrays does not work. Everything in java is nullable
                return client2.setpartmultipleusageattributes(sFullName, sParentPartNo, sChildParts, sLineNumber, sAttributeNames, sAttributeValues, sAttributeTypes, 3, Convert.ToInt16(sWebAppId));
            }
        }

        public string DeletePartToPartLink(string sSessionId, string sUserId, string sFullName, string sParentPartNo, string sChildPartNumber, string sCheckInComments, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                return client2.deletepartpartlink(sFullName, sParentPartNo, sChildPartNumber, sCheckInComments, Convert.ToInt16(sWebAppId));
            }
        }

        public string CreateActionRequest(string sSessionId, string sUserId, string sFullName, string sProductName, string sFolder, string sARName, string sARCategory, string sARCause, string sARComments,
                                          string sARLongDesc, string sARDate, string sRequestActionType, string sCheckInComments, string iProdOrLibrary, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                int iiProdOrLibrary = Convert.ToInt16(iProdOrLibrary);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                string[] sAttributeNames = new string[6];
                string[] sAttributeValues = new string[6];
                string[] sAttributeTypes = new string[6];

                sAttributeNames[0] = "Originator";
                sAttributeNames[1] = "ActionCategory";
                sAttributeNames[2] = "ARCause";
                sAttributeNames[3] = "LongDescription";
                sAttributeNames[4] = "Comments";
                sAttributeNames[5] = "RequestDate";

                sAttributeValues[0] = sFullName;
                sAttributeValues[1] = sARCategory;
                sAttributeValues[2] = sARCause;
                sAttributeValues[3] = sARLongDesc;
                sAttributeValues[4] = sARComments;
                sAttributeValues[5] = sARDate;

                sAttributeTypes[0] = "string";
                sAttributeTypes[1] = "string";
                sAttributeTypes[2] = "string";
                sAttributeTypes[3] = "string";
                sAttributeTypes[4] = "string";
                sAttributeTypes[5] = "datetime";

                if (sRequestActionType != "")
                {
                    Array.Resize<string>(ref sAttributeNames, 7);
                    Array.Resize<string>(ref sAttributeValues, 7);
                    Array.Resize<string>(ref sAttributeTypes, 7);
                    sAttributeNames[6] = "RequestedActionType";
                    sAttributeValues[6] = sRequestActionType;
                    sAttributeTypes[6] = "string";
                }

                return client2.createpart("", sARName, sProductName, "local.rs.vsrs05.Regain.RequestedAction", sFolder, sFullName, sAttributeNames, sAttributeValues, sAttributeTypes, sCheckInComments, iiProdOrLibrary, Convert.ToInt16(sWebAppId));
            }
        }

        public string CreateBatch(string sSessionId, string sUserId, string sFullName, string sBatchNo, string sBatchName, string sProductName, string sFolder, string sBatchType,
                                  string sCheckInComments, string iProdOrLibrary,
                                  string dTargetQty, string dActualQty, string dMoisturePercentage, string sQualityStatus,
                                  string dTargetAl2O3, string dActualAl2O3, string dTargetCaO, string dActualCaO, string dTargetF, string dActualF,
                                  string dTargetFe2O3, string dActualFe2O3, string dTargetK2O, string dActualK2O, string dTargetMgO, string dActualMgO,
                                  string dTargetMnO, string dActualMnO, string dTargetNa2O3, string dActualNa2O3, string dTargetSiO2, string dActualSiO2,
                                  string dTargetC, string dActualC, string dTargetSO3, string dActualSO3, string dTargetCN, string dActualCN, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                int iiProdOrLibrary = Convert.ToInt16(iProdOrLibrary);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                string[] sAttributeNames = new string[29];
                string[] sAttributeValues = new string[29];
                string[] sAttributeTypes = new string[29];

                sAttributeNames[0] = "Originator";
                sAttributeNames[1] = "QtyTarget";
                sAttributeNames[2] = "QtyActual";
                sAttributeNames[3] = "MoistureContent";
                sAttributeNames[4] = "QualityStatus";
                sAttributeNames[5] = "Al2O3Target";
                sAttributeNames[6] = "Al2O3Actual";
                sAttributeNames[7] = "CaOTarget";
                sAttributeNames[8] = "CaOActual";
                sAttributeNames[9] = "FTarget";
                sAttributeNames[10] = "FActual";
                sAttributeNames[11] = "Fe2O3Target";
                sAttributeNames[12] = "Fe2O3Actual";
                sAttributeNames[13] = "K2OTarget";
                sAttributeNames[14] = "K2OActual";
                sAttributeNames[15] = "MgOTarget";
                sAttributeNames[16] = "MgOActual";
                sAttributeNames[17] = "MnOTarget";
                sAttributeNames[18] = "MnOActual";
                sAttributeNames[19] = "Na2O3Target";
                sAttributeNames[20] = "Na2O3Actual";
                sAttributeNames[21] = "SiO2Target";
                sAttributeNames[22] = "SiO2Actual";
                sAttributeNames[23] = "CTarget";
                sAttributeNames[24] = "CActual";
                sAttributeNames[25] = "SO3Target";
                sAttributeNames[26] = "SO3Actual";
                sAttributeNames[27] = "CNTarget";
                sAttributeNames[28] = "CNActual";

                sAttributeValues[0] = sFullName;
                sAttributeValues[1] = dTargetQty;
                sAttributeValues[2] = dActualQty;
                sAttributeValues[3] = dMoisturePercentage;
                sAttributeValues[4] = sQualityStatus;
                sAttributeValues[5] = dTargetAl2O3;
                sAttributeValues[6] = dActualAl2O3;
                sAttributeValues[7] = dTargetCaO;
                sAttributeValues[8] = dActualCaO;
                sAttributeValues[9] = dTargetF;
                sAttributeValues[10] = dActualF;
                sAttributeValues[11] = dTargetFe2O3;
                sAttributeValues[12] = dActualFe2O3;
                sAttributeValues[13] = dTargetK2O;
                sAttributeValues[14] = dActualK2O;
                sAttributeValues[15] = dTargetMgO;
                sAttributeValues[16] = dActualMgO;
                sAttributeValues[17] = dTargetMnO;
                sAttributeValues[18] = dActualMnO;
                sAttributeValues[19] = dTargetNa2O3;
                sAttributeValues[20] = dActualNa2O3;
                sAttributeValues[21] = dTargetSiO2;
                sAttributeValues[22] = dActualSiO2;
                sAttributeValues[23] = dTargetC;
                sAttributeValues[24] = dActualC;
                sAttributeValues[25] = dTargetSO3;
                sAttributeValues[26] = dActualSO3;
                sAttributeValues[27] = dTargetCN;
                sAttributeValues[28] = dActualCN;

                sAttributeTypes[0] = "string";
                sAttributeTypes[1] = "double";
                sAttributeTypes[2] = "double";
                sAttributeTypes[3] = "double";
                sAttributeTypes[4] = "string";
                sAttributeTypes[5] = "double";
                sAttributeTypes[6] = "double";
                sAttributeTypes[7] = "double";
                sAttributeTypes[8] = "double";
                sAttributeTypes[9] = "double";
                sAttributeTypes[10] = "double";
                sAttributeTypes[11] = "double";
                sAttributeTypes[12] = "double";
                sAttributeTypes[13] = "double";
                sAttributeTypes[14] = "double";
                sAttributeTypes[15] = "double";
                sAttributeTypes[16] = "double";
                sAttributeTypes[17] = "double";
                sAttributeTypes[18] = "double";
                sAttributeTypes[19] = "double";
                sAttributeTypes[20] = "double";
                sAttributeTypes[21] = "double";
                sAttributeTypes[22] = "double";
                sAttributeTypes[23] = "double";
                sAttributeTypes[24] = "double";
                sAttributeTypes[25] = "double";
                sAttributeTypes[26] = "double";
                sAttributeTypes[27] = "double";
                sAttributeTypes[28] = "double";


                return client2.createpart(sBatchNo, sBatchName, sProductName, sBatchType, sFolder, sFullName, sAttributeNames, sAttributeValues, sAttributeTypes, sCheckInComments, iiProdOrLibrary, Convert.ToInt16(sWebAppId));
            }
        }

        public string CreateMBA(string sSessionId, string sUserId, string sFullName, string sBatchNo, string sBatchName, string sProductName, string sFolder, string sBatchType,
                                  string sCheckInComments, string iProdOrLibrary, string dMoisturePercentage, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                int iiProdOrLibrary = Convert.ToInt16(iProdOrLibrary);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                string[] sAttributeNames = new string[2];
                string[] sAttributeValues = new string[2];
                string[] sAttributeTypes = new string[2];

                sAttributeNames[0] = "Originator";
                sAttributeNames[1] = "MoistureContent";

                sAttributeValues[0] = sFullName;
                sAttributeValues[1] = dMoisturePercentage;

                sAttributeTypes[0] = "string";
                sAttributeTypes[1] = "double";

                return client2.createpart(sBatchNo, sBatchName, sProductName, sBatchType, sFolder, sFullName, sAttributeNames, sAttributeValues, sAttributeTypes, sCheckInComments, iiProdOrLibrary, Convert.ToInt16(sWebAppId));
            }
        }

        public string CopyPart(string sSessionId, string sUserId, string sSourcePartNo, string sTargetPartNo, string sTargetPartName, string sProductName, string sFolder,
                               string sPartType, string iProdOrLibrary, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                int iiProdOrLibrary = Convert.ToInt16(iProdOrLibrary);
                ExampleService.MyJavaService3Client client2 = GetWCService();

                return client2.copypart(sSourcePartNo, sTargetPartNo, sTargetPartName, sProductName, sPartType, sFolder, iiProdOrLibrary, Convert.ToInt16(sWebAppId));
            }
        }



        public string UpdateBatch(string sSessionId, string sUserId, string sFullName, string sBatchNo, string sBatchName, string sCheckinComments,
                                  string dTargetQty, string dActualQty, string dMoisturePercentage, string sQualityStatus,
                                  string dTargetAl2O3, string dActualAl2O3, string dTargetCaO, string dActualCaO, string dTargetF, string dActualF,
                                  string dTargetFe2O3, string dActualFe2O3, string dTargetK2O, string dActualK2O, string dTargetMgO, string dActualMgO,
                                  string dTargetMnO, string dActualMnO, string dTargetNa2O3, string dActualNa2O3, string dTargetSiO2, string dActualSiO2,
                                  string dTargetC, string dActualC, string dTargetSO3, string dActualSO3, string dTargetCN, string dActualCN, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                string[] sAttributeNames = new string[29];
                string[] sAttributeValues = new string[29];
                string[] sAttributeTypes = new string[29];

                sAttributeNames[0] = "Originator";
                sAttributeNames[1] = "QtyTarget";
                sAttributeNames[2] = "QtyActual";
                sAttributeNames[3] = "MoistureContent";
                sAttributeNames[4] = "QualityStatus";
                sAttributeNames[5] = "Al2O3Target";
                sAttributeNames[6] = "Al2O3Actual";
                sAttributeNames[7] = "CaOTarget";
                sAttributeNames[8] = "CaOActual";
                sAttributeNames[9] = "FTarget";
                sAttributeNames[10] = "FActual";
                sAttributeNames[11] = "Fe2O3Target";
                sAttributeNames[12] = "Fe2O3Actual";
                sAttributeNames[13] = "K2OTarget";
                sAttributeNames[14] = "K2OActual";
                sAttributeNames[15] = "MgOTarget";
                sAttributeNames[16] = "MgOActual";
                sAttributeNames[17] = "MnOTarget";
                sAttributeNames[18] = "MnOActual";
                sAttributeNames[19] = "Na2O3Target";
                sAttributeNames[20] = "Na2O3Actual";
                sAttributeNames[21] = "SiO2Target";
                sAttributeNames[22] = "SiO2Actual";
                sAttributeNames[23] = "CTarget";
                sAttributeNames[24] = "CActual";
                sAttributeNames[25] = "SO3Target";
                sAttributeNames[26] = "SO3Actual";
                sAttributeNames[27] = "CNTarget";
                sAttributeNames[28] = "CNActual";

                sAttributeValues[0] = sFullName;
                sAttributeValues[1] = dTargetQty;
                sAttributeValues[2] = dActualQty;
                sAttributeValues[3] = dMoisturePercentage;
                sAttributeValues[4] = sQualityStatus;
                sAttributeValues[5] = dTargetAl2O3;
                sAttributeValues[6] = dActualAl2O3;
                sAttributeValues[7] = dTargetCaO;
                sAttributeValues[8] = dActualCaO;
                sAttributeValues[9] = dTargetF;
                sAttributeValues[10] = dActualF;
                sAttributeValues[11] = dTargetFe2O3;
                sAttributeValues[12] = dActualFe2O3;
                sAttributeValues[13] = dTargetK2O;
                sAttributeValues[14] = dActualK2O;
                sAttributeValues[15] = dTargetMgO;
                sAttributeValues[16] = dActualMgO;
                sAttributeValues[17] = dTargetMnO;
                sAttributeValues[18] = dActualMnO;
                sAttributeValues[19] = dTargetNa2O3;
                sAttributeValues[20] = dActualNa2O3;
                sAttributeValues[21] = dTargetSiO2;
                sAttributeValues[22] = dActualSiO2;
                sAttributeValues[23] = dTargetC;
                sAttributeValues[24] = dActualC;
                sAttributeValues[25] = dTargetSO3;
                sAttributeValues[26] = dActualSO3;
                sAttributeValues[27] = dTargetCN;
                sAttributeValues[28] = dActualCN;

                sAttributeTypes[0] = "string";
                sAttributeTypes[1] = "double";
                sAttributeTypes[2] = "double";
                sAttributeTypes[3] = "double";
                sAttributeTypes[4] = "string";
                sAttributeTypes[5] = "double";
                sAttributeTypes[6] = "double";
                sAttributeTypes[7] = "double";
                sAttributeTypes[8] = "double";
                sAttributeTypes[9] = "double";
                sAttributeTypes[10] = "double";
                sAttributeTypes[11] = "double";
                sAttributeTypes[12] = "double";
                sAttributeTypes[13] = "double";
                sAttributeTypes[14] = "double";
                sAttributeTypes[15] = "double";
                sAttributeTypes[16] = "double";
                sAttributeTypes[17] = "double";
                sAttributeTypes[18] = "double";
                sAttributeTypes[19] = "double";
                sAttributeTypes[20] = "double";
                sAttributeTypes[21] = "double";
                sAttributeTypes[22] = "double";
                sAttributeTypes[23] = "double";
                sAttributeTypes[24] = "double";
                sAttributeTypes[25] = "double";
                sAttributeTypes[26] = "double";
                sAttributeTypes[27] = "double";
                sAttributeTypes[28] = "double";

                return client2.setpartattributes(sBatchNo, sBatchName, sFullName, sAttributeNames, sAttributeValues, sAttributeTypes, sCheckinComments, Convert.ToInt16(sWebAppId));
            }
        }

        public string UpdateMBA(string sSessionId, string sUserId, string sFullName, string sBatchNo, string sBatchName, string sCheckinComments, string dMoisturePercentage, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                string[] sAttributeNames = new string[2];
                string[] sAttributeValues = new string[2];
                string[] sAttributeTypes = new string[2];

                sAttributeNames[0] = "Originator";
                sAttributeNames[1] = "MoistureContent";

                sAttributeValues[0] = sFullName;
                sAttributeValues[1] = dMoisturePercentage;

                sAttributeTypes[0] = "string";
                sAttributeTypes[1] = "double";

                return client2.setpartattributes(sBatchNo, sBatchName, sFullName, sAttributeNames, sAttributeValues, sAttributeTypes, sCheckinComments, Convert.ToInt16(sWebAppId));
            }
        }

        public string CreateProductionLoss(string sSessionId, string sUserId, string sFullName, string sProdLossNo, string sProdLossName, string sProductName, string sPRType, string sFolderNameAndPath,
                                           string sPlant, string sRegainCategory, string sRegainSubCategory, string sStartDateAndTime, string sEndDateAndTime,
                                           string dDurationInHours, string sSuspectedFailureReason, string sComments, string iProdOrLibrary, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                double ddDurationInHours = Convert.ToDouble(dDurationInHours);
                int iiProdOrLibrary = Convert.ToInt16(iProdOrLibrary);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                string[] sAttributeNames = new string[9];
                string[] sAttributeValues = new string[9];
                string[] sAttributeTypes = new string[9];

                sAttributeNames[0] = "Originator";
                sAttributeNames[1] = "ProdLossCategory";
                sAttributeNames[2] = "ProdLossSubCategory";
                sAttributeNames[3] = "PlantCode";
                sAttributeNames[4] = "StartDate";
                sAttributeNames[5] = "EndDate";
                sAttributeNames[6] = "DurationInHours";
                sAttributeNames[7] = "SuspectedProblem";
                sAttributeNames[8] = "Comments";

                sAttributeValues[0] = sFullName;
                sAttributeValues[1] = sRegainCategory;
                sAttributeValues[2] = sRegainSubCategory;
                sAttributeValues[3] = sPlant;
                sAttributeValues[4] = sStartDateAndTime;
                sAttributeValues[5] = sEndDateAndTime;
                sAttributeValues[6] = ddDurationInHours.ToString();
                sAttributeValues[7] = sSuspectedFailureReason;
                sAttributeValues[8] = sComments;

                sAttributeTypes[0] = "string";
                sAttributeTypes[1] = "string";
                sAttributeTypes[2] = "string";
                sAttributeTypes[3] = "string";
                sAttributeTypes[4] = "string";
                sAttributeTypes[5] = "string";
                sAttributeTypes[6] = "double";
                sAttributeTypes[7] = "string";
                sAttributeTypes[8] = "string";

                return client2.createproblemreport(sProdLossNo, sProdLossName, sProductName, sPRType, sFolderNameAndPath, sAttributeNames, sAttributeValues, sAttributeTypes, iiProdOrLibrary, "", Convert.ToInt16(sWebAppId));
            }
        }

        public string UpdateProductionLoss(string sSessionId, string sUserId, string sFullName, string sProdLossNo, string sProdLossName, string sPlant, string sRegainCategory, string sRegainSubCategory,
                                           string sStartDateAndTime, string sEndDateAndTime, string dDurationInHours, string sSuspectedFailureReason, string sComments, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                double ddDurationInHours = Convert.ToDouble(dDurationInHours);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                string[] sAttributeNames = new string[9];
                string[] sAttributeValues = new string[9];
                string[] sAttributeTypes = new string[9];

                sAttributeNames[0] = "Originator";
                sAttributeNames[1] = "ProdLossCategory";
                sAttributeNames[2] = "ProdLossSubCategory";
                sAttributeNames[3] = "PlantCode";
                sAttributeNames[4] = "StartDate";
                sAttributeNames[5] = "EndDate";
                sAttributeNames[6] = "DurationInHours";
                sAttributeNames[7] = "SuspectedProblem";
                sAttributeNames[8] = "Comments";

                sAttributeValues[0] = sFullName;
                sAttributeValues[1] = sRegainCategory;
                sAttributeValues[2] = sRegainSubCategory;
                sAttributeValues[3] = sPlant;
                sAttributeValues[4] = sStartDateAndTime;
                sAttributeValues[5] = sEndDateAndTime;
                sAttributeValues[6] = ddDurationInHours.ToString();
                sAttributeValues[7] = sSuspectedFailureReason;
                sAttributeValues[8] = sComments;

                sAttributeTypes[0] = "string";
                sAttributeTypes[1] = "string";
                sAttributeTypes[2] = "string";
                sAttributeTypes[3] = "string";
                sAttributeTypes[4] = "string";
                sAttributeTypes[5] = "string";
                sAttributeTypes[6] = "double";
                sAttributeTypes[7] = "string";
                sAttributeTypes[8] = "string";

                return client2.setproblemreportattributes(sProdLossNo, sProdLossName, sAttributeNames, sAttributeValues, sAttributeTypes, "", Convert.ToInt16(sWebAppId));
            }
        }

        public string CreateTechnicalAction(string sSessionId, string sUserId, string sFullName, string sTechActionNo, string sTechActionName, string sProductName, string sPRType, string sFolderNameAndPath,
                                           string sPlantCode, string sTechActionDesc, string sComments, string sNeedDate, string iProdOrLibrary, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                int iiProdOrLibrary = Convert.ToInt16(iProdOrLibrary);
                ExampleService.MyJavaService3Client client2 = GetWCService();

                string[] sAttributeNames = new string[3];
                string[] sAttributeValues = new string[3];
                string[] sAttributeTypes = new string[3];

                sAttributeNames[0] = "Originator";
                sAttributeNames[1] = "PlantCode";
                sAttributeNames[2] = "Comments";

                sAttributeValues[0] = sFullName;
                sAttributeValues[1] = sPlantCode;
                sAttributeValues[2] = sComments;

                sAttributeTypes[0] = "string";
                sAttributeTypes[1] = "string";
                sAttributeTypes[2] = "string";

                return client2.createproblemreport2(sTechActionNo, sTechActionName, sTechActionDesc, sProductName, sPRType, sFolderNameAndPath, sAttributeNames, sAttributeValues, sAttributeTypes, iiProdOrLibrary, sNeedDate, Convert.ToInt16(sWebAppId));
            }
        }

        public string UpdateTechnicalAction(string sSessionId, string sUserId, string sFullName, string sTechActionNo, string sTechActionName, string sTechActionDesc,
                                            string sComments, string sNeedDate, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                string[] sAttributeNames = new string[3];
                string[] sAttributeValues = new string[3];
                string[] sAttributeTypes = new string[3];

                sAttributeNames[0] = "Originator";
                sAttributeNames[1] = "description";
                sAttributeNames[2] = "Comments";

                sAttributeValues[0] = sFullName;
                sAttributeValues[1] = sTechActionDesc;
                sAttributeValues[2] = sComments;

                sAttributeTypes[0] = "string";
                sAttributeTypes[1] = "string";
                sAttributeTypes[2] = "string";

                return client2.setproblemreportattributes(sTechActionNo, sTechActionName, sAttributeNames, sAttributeValues, sAttributeTypes, sNeedDate, Convert.ToInt16(sWebAppId));
            }
        }

        public string CreateIssueReport(string sSessionId, string sUserId, string sFullName, string sIssueRptNo, string sIssueRptName, string sPlant, string sProductName,
                                        string sPRType, string sFolderNameAndPath, string sComments, string iProdOrLibrary, string sNeedDate, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                int iiProdOrLibrary = Convert.ToInt16(iProdOrLibrary);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                string[] sAttributeNames = new string[3];
                string[] sAttributeValues = new string[3];
                string[] sAttributeTypes = new string[3];

                sAttributeNames[0] = "Originator";
                sAttributeNames[1] = "PlantCode";
                sAttributeNames[2] = "Comments";

                sAttributeValues[0] = sFullName;
                sAttributeValues[1] = sPlant;
                sAttributeValues[2] = sComments;

                sAttributeTypes[0] = "string";
                sAttributeTypes[1] = "string";
                sAttributeTypes[2] = "string";

                return client2.createproblemreport(sIssueRptNo, sIssueRptName, sProductName, sPRType, sFolderNameAndPath, sAttributeNames, sAttributeValues, sAttributeTypes, iiProdOrLibrary, sNeedDate, Convert.ToInt16(sWebAppId));
            }
        }

        public string CreateImprovementReport(string sSessionId, string sUserId, string sFullName, string sImprovementRptNo, string sImprovementRptName,
                                              string sPlant, string sProductName, string sPRType, string sFolderNameAndPath, string sComments,
                                              string iProdOrLibrary, string sNeedDate, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                int iiProdOrLibrary = Convert.ToInt16(iProdOrLibrary);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                string[] sAttributeNames = new string[0];
                string[] sAttributeValues = new string[0];
                string[] sAttributeTypes = new string[0];

                //sAttributeNames[0] = "Originator";
                //sAttributeNames[1] = "PlantCode";
                //sAttributeNames[2] = "Comments";

                //sAttributeValues[0] = sFullName;
                //sAttributeValues[1] = sPlant;
                //sAttributeValues[2] = sComments;

                //sAttributeTypes[0] = "string";
                //sAttributeTypes[1] = "string";
                //sAttributeTypes[2] = "string";

                return client2.createproblemreport(sImprovementRptNo, sImprovementRptName, sProductName, sPRType, sFolderNameAndPath, sAttributeNames, sAttributeValues, sAttributeTypes, iiProdOrLibrary, sNeedDate, Convert.ToInt16(sWebAppId));
            }
        }

        public string CreateBatchEvent(string sSessionId, string sUserId, string sFullName, string sBatchEventNo, string sBatchEventName, string sProductName,
                                       string sPRType, string sFolderNameAndPath, string sComments, string iProdOrLibrary, string sTransDate, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                int iiProdOrLibrary = Convert.ToInt16(iProdOrLibrary);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                string[] sAttributeNames = new string[3];
                string[] sAttributeValues = new string[3];
                string[] sAttributeTypes = new string[3];

                sAttributeNames[0] = "Originator";
                sAttributeNames[1] = "Comments";
                sAttributeNames[2] = "DispatchDocketDate"; //For some reason this has to be the underlying global attribute name

                sAttributeValues[0] = sFullName;
                sAttributeValues[1] = sComments;
                sAttributeValues[2] = sTransDate;


                sAttributeTypes[0] = "string";
                sAttributeTypes[1] = "string";
                sAttributeTypes[2] = "date";

                return client2.createproblemreport(sBatchEventNo, sBatchEventName, sProductName, sPRType, sFolderNameAndPath, sAttributeNames, sAttributeValues, sAttributeTypes, iiProdOrLibrary, "", Convert.ToInt16(sWebAppId));
            }
        }

        public string UpdateIssueReport(string sSessionId, string sUserId, string sFullName, string sIssueRptNo, string sIssueRptName, string sPlant, string sComments, string sNeedDate, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                string[] sAttributeNames = new string[3];
                string[] sAttributeValues = new string[3];
                string[] sAttributeTypes = new string[3];

                sAttributeNames[0] = "Originator";
                sAttributeNames[1] = "PlantCode";
                sAttributeNames[2] = "Comments";

                sAttributeValues[0] = sFullName;
                sAttributeValues[1] = sPlant;
                sAttributeValues[2] = sComments;

                sAttributeTypes[0] = "string";
                sAttributeTypes[1] = "string";
                sAttributeTypes[2] = "string";


                return client2.setproblemreportattributes(sIssueRptNo, sIssueRptName, sAttributeNames, sAttributeValues, sAttributeTypes, sNeedDate, Convert.ToInt16(sWebAppId));
            }
        }

        public string UpdateImprovementReport(string sSessionId, string sUserId, string sFullName, string sImprovementRptNo, string sImprovementRptName, string sPlant, string sComments, string sNeedDate, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                string[] sAttributeNames = new string[3];
                string[] sAttributeValues = new string[3];
                string[] sAttributeTypes = new string[3];

                sAttributeNames[0] = "Originator";
                sAttributeNames[1] = "PlantCode";
                sAttributeNames[2] = "Comments";

                sAttributeValues[0] = sFullName;
                sAttributeValues[1] = sPlant;
                sAttributeValues[2] = sComments;

                sAttributeTypes[0] = "string";
                sAttributeTypes[1] = "string";
                sAttributeTypes[2] = "string";


                return client2.setproblemreportattributes(sImprovementRptNo, sImprovementRptName, sAttributeNames, sAttributeValues, sAttributeTypes, sNeedDate, Convert.ToInt16(sWebAppId));
            }
        }

        public string UpdateBatchEvent(string sSessionId, string sUserId, string sFullName, string sBatchEventNo, string sBatchEventName, string sComments, string sTransDate, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                string[] sAttributeNames = new string[3];
                string[] sAttributeValues = new string[3];
                string[] sAttributeTypes = new string[3];

                sAttributeNames[0] = "Originator";
                sAttributeNames[1] = "Comments";
                sAttributeNames[2] = "DispatchDocketDate"; //For some reason this has to be the underlying global attribute name

                sAttributeValues[0] = sFullName;
                sAttributeValues[1] = sComments;
                sAttributeValues[2] = sTransDate;

                sAttributeTypes[0] = "string";
                sAttributeTypes[1] = "string";
                sAttributeTypes[2] = "date";

                return client2.setproblemreportattributes(sBatchEventNo, sBatchEventName, sAttributeNames, sAttributeValues, sAttributeTypes, "", Convert.ToInt16(sWebAppId));
            }
        }

        public string SetTaskNextElapsedDateOnCompletion(string sSessionId, string sUserId, string sWorkItemId, string sAssignedActivityId, string sRoute, string sNextElapsedDate, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                string[] sVariableNames = new string[1];
                string[] sVariableValues = new string[1];
                string[] sVariableTypes = new string[1];

                sVariableNames[0] = "gdtElapsedNextDateLocal"; //THis is the variable in the task. A local variable gdtElapsedNextDateLocal
                sVariableValues[0] = sNextElapsedDate;
                sVariableTypes[0] = "date";

                return client2.completetask(Convert.ToInt32(sWorkItemId), Convert.ToInt32(sAssignedActivityId), sRoute, sVariableNames, sVariableTypes, sVariableValues, Convert.ToInt16(sWebAppId));
            }
        }

        public string SetTaskOperationalHoursOnCompletion(string sSessionId, string sUserId, string sWorkItemId, string sAssignedActivityId, string sHoursOnCompletion, string sRoute, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                string[] sVariableNames = new string[1];
                string[] sVariableValues = new string[1];
                string[] sVariableTypes = new string[1];

                sVariableNames[0] = "giAccumThreshold";
                sVariableValues[0] = sHoursOnCompletion;
                sVariableTypes[0] = "int";

                return client2.completetask(Convert.ToInt32(sWorkItemId), Convert.ToInt32(sAssignedActivityId), sRoute, sVariableNames, sVariableTypes, sVariableValues, Convert.ToInt16(sWebAppId));
            }
        }

        //This is a function to simply progress a WO and set the completion date.
        public string SetTaskWOCompletionDate(string sSessionId, string sUserId, string sWorkItemId, string sAssignedActivityId, string sRoute, string sDateOnCompletion, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                string[] sVariableNames = new string[1];
                string[] sVariableValues = new string[1];
                string[] sVariableTypes = new string[1];

                sVariableNames[0] = "dtCompletedDate";
                sVariableValues[0] = sDateOnCompletion;
                sVariableTypes[0] = "date";

                return client2.completetask(Convert.ToInt32(sWorkItemId), Convert.ToInt32(sAssignedActivityId), sRoute, sVariableNames, sVariableTypes, sVariableValues, Convert.ToInt16(sWebAppId));
            }
        }

        public string ProgressTask(string sSessionId, string sUserId, string sWorkItemId, string sAssignedActivityId, string sRoute, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                string[] sVariableNames = new string[0];
                string[] sVariableValues = new string[0];
                string[] sVariableTypes = new string[0];

                return client2.completetask(Convert.ToInt32(sWorkItemId), Convert.ToInt32(sAssignedActivityId), sRoute, sVariableNames, sVariableTypes, sVariableValues, Convert.ToInt16(sWebAppId));
            }
        }

        public string SetProbRptAffectedObjects(string sSessionId, string sUserId, string sProdLossNo, string sAffectdObjectsString, string sAffectdObjectTypesString, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                char[] charSeparators = new char[] { '^' };
                string[] sAffectedParts = sAffectdObjectsString.Split(charSeparators, StringSplitOptions.RemoveEmptyEntries);
                string[] sAffectedParts2 = sAffectdObjectTypesString.Split(charSeparators, StringSplitOptions.RemoveEmptyEntries);
                int i;
                int?[] iAffectedPartTypes = new int?[sAffectedParts2.Length];

                for (i = 0; i < sAffectedParts2.Length; i++)
                {
                    iAffectedPartTypes[i] = Convert.ToInt16(sAffectedParts2[i]);
                }
                return client2.setpraffectedobjects(sProdLossNo, sAffectedParts, iAffectedPartTypes, Convert.ToInt16(sWebAppId));
            }
        }

        public string SetProbRptState(string sSessionId, string sUserId, string sFullName, string sProbRptNo, string sProbRptName, string sLifecycleState, string sComments, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();

                string[] sAttributeNames = new string[1];
                string[] sAttributeValues = new string[1];
                string[] sAttributeTypes = new string[2];

                sAttributeNames[0] = "Originator";

                sAttributeValues[0] = sFullName;

                sAttributeTypes[0] = "string";

                if (sComments != "")
                {
                    Array.Resize<string>(ref sAttributeNames, 2);
                    Array.Resize<string>(ref sAttributeValues, 2);
                    Array.Resize<string>(ref sAttributeTypes, 2);

                    sAttributeNames[1] = "Comments";
                    sAttributeValues[1] = sComments;
                    sAttributeTypes[1] = "string";
                }

                //Send empty string for need date so it is ignored at the Windchill end
                string sRtn = client2.setproblemreportattributes(sProbRptNo, sProbRptName, sAttributeNames, sAttributeValues, sAttributeTypes, "", Convert.ToInt16(sWebAppId));
                if (sRtn == "Success")
                {
                    sRtn = client2.setproblemreportstate(sProbRptNo, sLifecycleState, Convert.ToInt16(sWebAppId));
                }

                return sRtn;

            }
        }


        public string DeleteProbRptAffectedObjects(string sSessionId, string sUserId, string sProdLossNo, string sAffectdPartsString, string sAffectdObjectTypesString, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                char[] charSeparators = new char[] { '^' };
                string[] sAffectedParts = sAffectdPartsString.Split(charSeparators, StringSplitOptions.RemoveEmptyEntries);
                string[] sAffectedParts2 = sAffectdObjectTypesString.Split(charSeparators, StringSplitOptions.RemoveEmptyEntries);
                int i;
                int?[] iAffectedPartTypes = new int?[sAffectedParts2.Length];

                for (i = 0; i < sAffectedParts2.Length; i++)
                {
                    iAffectedPartTypes[i] = Convert.ToInt16(sAffectedParts2[i]);
                }

                return client2.deletepraffectedobjects(sProdLossNo, sAffectedParts, iAffectedPartTypes, Convert.ToInt16(sWebAppId));
            }
        }

        public string AttachProductionLossDoc(string sSessionId, string sUserId, string sProdLossNo, string sAttachDesc, string sAttachPath, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                return client2.attachprdoc(sProdLossNo, sAttachDesc, sAttachPath, Convert.ToInt16(sWebAppId));
            }
        }

        public string DeleteProductionLossAttachment(string sSessionId, string sUserId, string sProdLossNo, string sAttachFileName, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                return client2.deleteprattachment(sProdLossNo, sAttachFileName, Convert.ToInt16(sWebAppId));
            }
        }

        public string DeleteProblemReport(string sSessionId, string sUserId, string sProbReportNo, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                return client2.deleteprobreport(sProbReportNo, Convert.ToInt16(sWebAppId));
            }
        }

        public string ReviseDocument(string sSessionId, string sUserId, string sDocNo, string sRevision, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                return client2.setdocrevision(sDocNo, sRevision, Convert.ToInt16(sWebAppId));
            }
        }

        public string ReviseDocumentAndRemoveAttachments(string sSessionId, string sUserId, string sFullname, string sDocNo, string sDocName, string sRevision,
                                                         string sLongDesc, string sOriginator, string sOriginatorDocId, string sJobCode,
                                                         string sCheckInComments, string sIncludeHyperlinks, string sWebAppId)
        {

            //String sDocNumber, String sDocName, String sRevision, String sUser,  
            //				  String[] sAttributeName, String[] sAttributeValue, String[] sAttributeType,
            //                                                      String sCheckInComments, int iWebAppId
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();

                string[] sAttributeNames = new string[3];
                string[] sAttributeValues = new string[3];
                string[] sAttributeTypes = new string[3];

                sAttributeNames[0] = "LongDescription";
                sAttributeNames[1] = "Originator";
                sAttributeNames[2] = "OrigDocId";

                sAttributeValues[0] = sLongDesc;
                sAttributeValues[1] = sOriginator;
                sAttributeValues[2] = sOriginatorDocId;

                sAttributeTypes[0] = "string";
                sAttributeTypes[1] = "string";
                sAttributeTypes[2] = "string";

                if (sJobCode != "")
                {
                    Array.Resize<string>(ref sAttributeNames, 4);
                    Array.Resize<string>(ref sAttributeValues, 4);
                    Array.Resize<string>(ref sAttributeTypes, 4);

                    sAttributeNames[3] = "JobCode";
                    sAttributeValues[3] = sJobCode;
                    sAttributeTypes[3] = "string";
                }

                //                return client2.setdocrevision(sDocNo, sRevision, Convert.ToInt16(sWebAppId));
                int iIncludeHyperlinks = Convert.ToInt16(sIncludeHyperlinks);

                string sRtn = client2.setdocrevremoveattachs(sDocNo, sDocName, sRevision, sFullname,
                                                                       sAttributeNames, sAttributeValues, sAttributeTypes,
                                                                       sCheckInComments, iIncludeHyperlinks, Convert.ToInt16(sWebAppId));
                return sRtn;
            }
        }
        // sAttributeType can have values
        //		Boolean - bool or boolean
        //		Date & Time - date or datetime
        //		Integer Number - int or integer
        //		Real Number - real or doub or double or float
        //		String - string or the default
        public string UpdateDocAttributes(string sSessionId, string sUserId, string sDocNumber, string sDocName,
                                          string sAttributeName1, string sAttributeValue1, string sAttributeType1,
                                          string sAttributeName2, string sAttributeValue2, string sAttributeType2,
                                          string sAttributeName3, string sAttributeValue3, string sAttributeType3,
                                          string sAttributeName4, string sAttributeValue4, string sAttributeType4,
                                          string sAttributeName5, string sAttributeValue5, string sAttributeType5,
                                          string sCheckinComments, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                string[] sAttributeNames = new string[1];
                string[] sAttributeValues = new string[1];
                string[] sAttributeTypes = new string[1];

                sAttributeNames[0] = sAttributeName1;
                sAttributeValues[0] = sAttributeValue1;
                sAttributeTypes[0] = sAttributeType1;

                if (sAttributeName2 != null)
                {
                    Array.Resize<string>(ref sAttributeNames, 2);
                    Array.Resize<string>(ref sAttributeValues, 2);
                    Array.Resize<string>(ref sAttributeTypes, 2);

                    sAttributeNames[1] = sAttributeName2;
                    sAttributeValues[1] = sAttributeValue2;
                    sAttributeTypes[1] = sAttributeType2;

                    if (sAttributeName3 != null)
                    {
                        Array.Resize<string>(ref sAttributeNames, 3);
                        Array.Resize<string>(ref sAttributeValues, 3);
                        Array.Resize<string>(ref sAttributeTypes, 3);

                        sAttributeNames[2] = sAttributeName3;
                        sAttributeValues[2] = sAttributeValue3;
                        sAttributeTypes[2] = sAttributeType3;

                        if (sAttributeName4 != null)
                        {
                            Array.Resize<string>(ref sAttributeNames, 4);
                            Array.Resize<string>(ref sAttributeValues, 4);
                            Array.Resize<string>(ref sAttributeTypes, 4);

                            sAttributeNames[3] = sAttributeName4;
                            sAttributeValues[3] = sAttributeValue4;
                            sAttributeTypes[3] = sAttributeType4;

                            if (sAttributeName5 != null)
                            {
                                Array.Resize<string>(ref sAttributeNames, 5);
                                Array.Resize<string>(ref sAttributeValues, 5);
                                Array.Resize<string>(ref sAttributeTypes, 5);

                                sAttributeNames[4] = sAttributeName5;
                                sAttributeValues[4] = sAttributeValue5;
                                sAttributeTypes[4] = sAttributeType5;

                            }
                        }
                    }
                }

                string sReturn = client2.setdocattributes(sDocNumber, sDocName, sAttributeNames, sAttributeValues, sAttributeTypes, sCheckinComments, Convert.ToInt16(sWebAppId));

                return sReturn;
            }
        }

        // sAttributeType can have values
        //		Boolean - bool or boolean
        //		Date & Time - date or datetime
        //		Integer Number - int or integer
        //		Real Number - real or doub or double or float
        //		String - string or the default
        public string UpdatePartAttributes(string sSessionId, string sUserId, string sPartNumber, string sPartName,
                                          string sAttributeName1, string sAttributeValue1, string sAttributeType1,
                                          string sAttributeName2, string sAttributeValue2, string sAttributeType2,
                                          string sAttributeName3, string sAttributeValue3, string sAttributeType3,
                                          string sAttributeName4, string sAttributeValue4, string sAttributeType4,
                                          string sAttributeName5, string sAttributeValue5, string sAttributeType5,
                                          string sCheckinComments, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ArrayList arrUser = GetUserDetails(sUserId);

                string sFullName = arrUser[2].ToString();

                ExampleService.MyJavaService3Client client2 = GetWCService();
                string[] sAttributeNames = new string[1];
                string[] sAttributeValues = new string[1];
                string[] sAttributeTypes = new string[1];

                sAttributeNames[0] = sAttributeName1;
                sAttributeValues[0] = sAttributeValue1;
                sAttributeTypes[0] = sAttributeType1;

                if (sAttributeName2 != null)
                {
                    Array.Resize<string>(ref sAttributeNames, 2);
                    Array.Resize<string>(ref sAttributeValues, 2);
                    Array.Resize<string>(ref sAttributeTypes, 2);

                    sAttributeNames[1] = sAttributeName2;
                    sAttributeValues[1] = sAttributeValue2;
                    sAttributeTypes[1] = sAttributeType2;

                    if (sAttributeName3 != null)
                    {
                        Array.Resize<string>(ref sAttributeNames, 3);
                        Array.Resize<string>(ref sAttributeValues, 3);
                        Array.Resize<string>(ref sAttributeTypes, 3);

                        sAttributeNames[2] = sAttributeName3;
                        sAttributeValues[2] = sAttributeValue3;
                        sAttributeTypes[2] = sAttributeType3;

                        if (sAttributeName4 != null)
                        {
                            Array.Resize<string>(ref sAttributeNames, 4);
                            Array.Resize<string>(ref sAttributeValues, 4);
                            Array.Resize<string>(ref sAttributeTypes, 4);

                            sAttributeNames[3] = sAttributeName4;
                            sAttributeValues[3] = sAttributeValue4;
                            sAttributeTypes[3] = sAttributeType4;

                            if (sAttributeName5 != null)
                            {
                                Array.Resize<string>(ref sAttributeNames, 5);
                                Array.Resize<string>(ref sAttributeValues, 5);
                                Array.Resize<string>(ref sAttributeTypes, 5);

                                sAttributeNames[4] = sAttributeName5;
                                sAttributeValues[4] = sAttributeValue5;
                                sAttributeTypes[4] = sAttributeType5;

                            }
                        }
                    }
                }

                string sReturn = client2.setpartattributes(sPartNumber, sPartName, sFullName, sAttributeNames, sAttributeValues, sAttributeTypes, sCheckinComments, Convert.ToInt16(sWebAppId));

                return sReturn;
            }
        }

        // sAttributeType can have values
        //		Boolean - bool or boolean
        //		Date & Time - date or datetime
        //		Integer Number - int or integer
        //		Real Number - real or doub or double or float
        //		String - string or the default
        public string UpdateOperatingHours(string sSessionId, string sUserId, string sPartNumber,
                                          string sOriginatorName, string sOperatingHours,
                                          string sCheckinComments, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);

                ArrayList arrUser = GetUserDetails(sUserId);

                string sFullName = arrUser[2].ToString();

                ExampleService.MyJavaService3Client client2 = GetWCService();
                string[] sAttributeNames = new string[2];
                string[] sAttributeValues = new string[2];
                string[] sAttributeTypes = new string[2];

                sAttributeNames[0] = "Originator";
                sAttributeValues[0] = sOriginatorName;
                sAttributeTypes[0] = "string";

                sAttributeNames[1] = "MonitorMeasurement";
                sAttributeValues[1] = sOperatingHours;
                sAttributeTypes[1] = "double";


                string sReturn = client2.setpartattributes(sPartNumber, "", sFullName, sAttributeNames, sAttributeValues, sAttributeTypes, sCheckinComments, Convert.ToInt16(sWebAppId));

                return sReturn;
            }
        }

        public string CreateMBAPartUsageLink(string sSessionId, string sUserId, string sFullName, string sParentPartNo, string sChildPartNo, string dQty,
                                             string lLineNumber, string sCheckInComments, string sDispatchDocketNo,
                                             string sTransactionDate, string sComments, string sProdOrderNo, string sMoisturePercentage,
                                             string sContainerId, string sInvoiceStatus, string sBatchNo, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                long llLineNumber = Convert.ToInt64(lLineNumber);
                double ddQty = Convert.ToDouble(dQty);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                string[] sAttributeNames = new string[8];
                string[] sAttributeValues = new string[8];
                string[] sAttributeTypes = new string[8];

                sAttributeNames[0] = "DispatchDocketNo";
                sAttributeNames[1] = "DispatchDocketDate"; //For some reason this has to be the underlying global attribute name
                sAttributeNames[2] = "UsageComments";
                sAttributeNames[3] = "ProdOrderNo";
                sAttributeNames[4] = "MoistureContent";
                sAttributeNames[5] = "ContainerId";
                sAttributeNames[6] = "InvoiceStatus";
                sAttributeNames[7] = "BatchNo";


                sAttributeValues[0] = sDispatchDocketNo;
                sAttributeValues[1] = sTransactionDate;
                sAttributeValues[2] = sComments;
                sAttributeValues[3] = sProdOrderNo;
                sAttributeValues[4] = sMoisturePercentage;
                sAttributeValues[5] = sContainerId;
                sAttributeValues[6] = sInvoiceStatus;
                sAttributeValues[7] = sBatchNo;

                sAttributeTypes[0] = "string";
                sAttributeTypes[1] = "date";
                sAttributeTypes[2] = "string";
                sAttributeTypes[3] = "string";
                sAttributeTypes[4] = "double";
                sAttributeTypes[5] = "string";
                sAttributeTypes[6] = "long";
                sAttributeTypes[7] = "string";


                return client2.setpartpartlinkwithattributes(sFullName, sParentPartNo, sChildPartNo, ddQty, sCheckInComments, "local.rs.vsrs05.Regain.MBAUsageLink", "tonne", llLineNumber, sAttributeNames, sAttributeValues, sAttributeTypes, Convert.ToInt16(sWebAppId));
            }
        }

        public string UpdateMBAPartUsageLinkFromDD(string sSessionId, string sUserId, string sFullName, string sParentPartNo, string sChildPartNo,
                                                   string dQty, string lOldLineNumber, string lNewLineNumber, string sCheckInComments, string sDispatchDocketNo,
                                                   string sTransactionDate, string sComments, string sProdOrderNo, string sContainerId, string sMoisturePercentage,
                                                   string sInvoiceStatus, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                long llOldLineNumber = Convert.ToInt64(lOldLineNumber);
                long llNewLineNumber = Convert.ToInt64(lNewLineNumber);
                double ddQty = Convert.ToDouble(dQty);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                string[] sAttributeNames = new string[7];
                string[] sAttributeValues = new string[7];
                string[] sAttributeTypes = new string[7];

                sAttributeNames[0] = "DispatchDocketNo";
                sAttributeNames[1] = "DispatchDocketDate"; //For some reason this has to be the underlying global attribute name
                sAttributeNames[2] = "UsageComments";
                sAttributeNames[3] = "ProdOrderNo";
                sAttributeNames[4] = "ContainerId";
                sAttributeNames[5] = "MoistureContent";
                sAttributeNames[6] = "InvoiceStatus";

                sAttributeValues[0] = sDispatchDocketNo;
                sAttributeValues[1] = sTransactionDate;
                sAttributeValues[2] = sComments;
                sAttributeValues[3] = sProdOrderNo;
                sAttributeValues[4] = sContainerId;
                sAttributeValues[5] = sMoisturePercentage;
                sAttributeValues[6] = sInvoiceStatus;

                sAttributeTypes[0] = "string";
                sAttributeTypes[1] = "date";
                sAttributeTypes[2] = "string";
                sAttributeTypes[3] = "string";
                sAttributeTypes[4] = "string";
                sAttributeTypes[5] = "double";
                sAttributeTypes[6] = "long";

                return client2.updatedispatchdocketpartpartlinkwithattributes(sFullName, sParentPartNo, sChildPartNo, ddQty, sDispatchDocketNo, sCheckInComments, "local.rs.vsrs05.Regain.MBAUsageLink", "tonne",
                                                                              llOldLineNumber, llNewLineNumber, sAttributeNames, sAttributeValues, sAttributeTypes, Convert.ToInt16(sWebAppId));
            }
        }

        public string UpdateMBAPartUsageLinkFromPO(string sSessionId, string sUserId, string sFullName, string sParentPartNo, string sChildPartNo,
                                                   string dQty, string lOldLineNumber, string lNewLineNumber, string sCheckInComments, string sDispatchDocketNo,
                                                   string sTransactionDate, string sComments, string sProdOrderNo, string sMoisturePercentage, string sInvoiceStatus,
                                                   string sBatchNo, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                long llOldLineNumber = Convert.ToInt64(lOldLineNumber);
                long llNewLineNumber = Convert.ToInt64(lNewLineNumber);
                double ddQty = Convert.ToDouble(dQty);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                string[] sAttributeNames = new string[7];
                string[] sAttributeValues = new string[7];
                string[] sAttributeTypes = new string[7];

                sAttributeNames[0] = "DispatchDocketNo";
                sAttributeNames[1] = "DispatchDocketDate"; //For some reason this has to be the underlying global attribute name
                sAttributeNames[2] = "UsageComments";
                sAttributeNames[3] = "ProdOrderNo";
                sAttributeNames[4] = "MoistureContent";
                sAttributeNames[5] = "InvoiceStatus";
                sAttributeNames[6] = "BatchNo";


                sAttributeValues[0] = sDispatchDocketNo;
                sAttributeValues[1] = sTransactionDate;
                sAttributeValues[2] = sComments;
                sAttributeValues[3] = sProdOrderNo;
                sAttributeValues[4] = sMoisturePercentage;
                sAttributeValues[5] = sInvoiceStatus;
                sAttributeValues[6] = sBatchNo;

                sAttributeTypes[0] = "string";
                sAttributeTypes[1] = "date";
                sAttributeTypes[2] = "string";
                sAttributeTypes[3] = "string";
                sAttributeTypes[4] = "double";
                sAttributeTypes[5] = "long";
                sAttributeTypes[6] = "string";


                return client2.updateprodorderpartpartlinkwithattributes(sFullName, sParentPartNo, sChildPartNo, ddQty, sProdOrderNo, sCheckInComments, "local.rs.vsrs05.Regain.MBAUsageLink", "tonne",
                                                                              llOldLineNumber, llNewLineNumber, sAttributeNames, sAttributeValues, sAttributeTypes, Convert.ToInt16(sWebAppId));
            }
        }

        public string DeletePartToPartLinkByDispatchDocket(string sSessionId, string sUserId, string sFullName, string sDispatchDocketNo, string lLineNumber,
                                                           string sParentPartNo, string sChildPartNumber, string sCheckInComments, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                long llLineNumber = Convert.ToInt64(lLineNumber);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                return client2.deletepartpartlinkbydispatchdocket(sFullName, sDispatchDocketNo, llLineNumber, sParentPartNo, sChildPartNumber, sCheckInComments, Convert.ToInt16(sWebAppId));
            }
        }

        public string DeletePartToPartLinkByProductionOrder(string sSessionId, string sUserId, string sFullName, string sProductionOrderNo, string lLineNumber, string sParentPartNo,
                                                            string sChildPartNumber, string sCheckInComments, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                long llLineNumber = Convert.ToInt64(lLineNumber);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                return client2.deletepartpartlinkbyproductionorder(sFullName, sProductionOrderNo, llLineNumber, sParentPartNo, sChildPartNumber, sCheckInComments, Convert.ToInt16(sWebAppId));
            }
        }

        public string DeletePartToPartLinkByLineNumber(string sSessionId, string sUserId, string sFullName, string lLineNumber, string sParentPartNo,
                                                            string sChildPartNumber, string sCheckInComments, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                if (lLineNumber.Equals("0") || lLineNumber.Equals("-1"))
                {
                    return client2.deletepartpartlink(sFullName, sParentPartNo, sChildPartNumber, sCheckInComments, Convert.ToInt16(sWebAppId));

                }
                else
                {
                    long llLineNumber = Convert.ToInt64(lLineNumber);
                    return client2.deletepartpartlinkbylinenumber(sFullName, llLineNumber, sParentPartNo, sChildPartNumber, sCheckInComments, Convert.ToInt16(sWebAppId));
                }
            }
        }

        public string CreateProdOrder(string sSessionId, string sUserId, string sDocNo, string sDocName, string sProductName, string sDocType, string sFolderNameAndPath,
                                      string sBatchNo, string sTargetQty, string sProdNoDate, string sOriginator, string sJobCode, string sComments, string sRevision,
                                      string sCheckInComments, string iProdOrLibrary, string sWebAppId)
        {

            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                string sRtn = "";
                Update_User_Time(sUserId, sSessionId);
                int iiProdOrLibrary = Convert.ToInt16(iProdOrLibrary);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                string[] sAttributeNames = new string[2];
                string[] sAttributeValues = new string[2];
                string[] sAttributeTypes = new string[2];

                sAttributeNames[0] = "JobCode";
                sAttributeNames[1] = "Originator";

                sAttributeValues[0] = sJobCode;
                sAttributeValues[1] = sOriginator;

                string sRtn1 = client2.doccreate2(sDocNo, sDocName, sProductName, sDocType, sFolderNameAndPath, sRevision, sAttributeNames, sAttributeValues, sCheckInComments, iiProdOrLibrary, Convert.ToInt16(sWebAppId));

                if (sRtn1.StartsWith("Success"))
                {

                    //Get the new document number
                    string[] sSuccess = Extract_Values(sRtn1);

                    sDocNo = sSuccess[1];

                    if (sBatchNo != "")
                    {
                        if (sComments != "")
                        {
                            if (sTargetQty != "")
                            {

                                string[] sAttributeNames1 = new string[4];
                                string[] sAttributeValues1 = new string[4];
                                string[] sAttributeTypes1 = new string[4];
                                sAttributeNames1[0] = "DispatchDocketDate";
                                sAttributeNames1[1] = "BatchNo";
                                sAttributeNames1[2] = "Comments";
                                sAttributeNames1[3] = "TargetQty";

                                sAttributeValues1[0] = sProdNoDate;
                                sAttributeValues1[1] = sBatchNo;
                                sAttributeValues1[2] = sComments;
                                sAttributeValues1[3] = sTargetQty;

                                sAttributeTypes1[0] = "date";
                                sAttributeTypes1[1] = "string";
                                sAttributeTypes1[2] = "string";
                                sAttributeTypes1[3] = "double";

                                sRtn = client2.setdocattributes(sDocNo, sDocName, sAttributeNames1, sAttributeValues1, sAttributeTypes1, sCheckInComments, Convert.ToInt16(sWebAppId));

                            }
                            else
                            {

                                string[] sAttributeNames1 = new string[3];
                                string[] sAttributeValues1 = new string[3];
                                string[] sAttributeTypes1 = new string[3];
                                sAttributeNames1[0] = "DispatchDocketDate";
                                sAttributeNames1[1] = "BatchNo";
                                sAttributeNames1[2] = "Comments";

                                sAttributeValues1[0] = sProdNoDate;
                                sAttributeValues1[1] = sBatchNo;
                                sAttributeValues1[2] = sComments;

                                sAttributeTypes1[0] = "date";
                                sAttributeTypes1[1] = "string";
                                sAttributeTypes1[2] = "string";

                                sRtn = client2.setdocattributes(sDocNo, sDocName, sAttributeNames1, sAttributeValues1, sAttributeTypes1, sCheckInComments, Convert.ToInt16(sWebAppId));
                            }


                        }
                        else
                        {
                            if (sTargetQty != "")
                            {
                                string[] sAttributeNames1 = new string[3];
                                string[] sAttributeValues1 = new string[3];
                                string[] sAttributeTypes1 = new string[3];
                                sAttributeNames1[0] = "DispatchDocketDate";
                                sAttributeNames1[1] = "BatchNo";
                                sAttributeNames1[2] = "TargetQty";

                                sAttributeValues1[0] = sProdNoDate;
                                sAttributeValues1[1] = sBatchNo;
                                sAttributeValues1[2] = sTargetQty;

                                sAttributeTypes1[0] = "date";
                                sAttributeTypes1[1] = "string";
                                sAttributeTypes1[2] = "double";

                                sRtn = client2.setdocattributes(sDocNo, sDocName, sAttributeNames1, sAttributeValues1, sAttributeTypes1, sCheckInComments, Convert.ToInt16(sWebAppId));
                            }
                            else
                            {
                                sAttributeNames[0] = "DispatchDocketDate";
                                sAttributeNames[1] = "BatchNo";

                                sAttributeValues[0] = sProdNoDate;
                                sAttributeValues[1] = sBatchNo;

                                sAttributeTypes[0] = "date";
                                sAttributeTypes[1] = "string";

                                sRtn = client2.setdocattributes(sDocNo, sDocName, sAttributeNames, sAttributeValues, sAttributeTypes, sCheckInComments, Convert.ToInt16(sWebAppId));

                            }

                        }
                    }
                    else
                    {
                        if (sComments != "")
                        {
                            if (sTargetQty != "")
                            {
                                string[] sAttributeNames1 = new string[3];
                                string[] sAttributeValues1 = new string[3];
                                string[] sAttributeTypes1 = new string[3];
                                sAttributeNames1[0] = "DispatchDocketDate";
                                sAttributeNames1[1] = "Comments";
                                sAttributeNames1[2] = "TargetQty";

                                sAttributeValues1[0] = sProdNoDate;
                                sAttributeValues1[1] = sComments;
                                sAttributeValues1[2] = sTargetQty;

                                sAttributeTypes1[0] = "date";
                                sAttributeTypes1[1] = "string";
                                sAttributeTypes1[2] = "double";

                                sRtn = client2.setdocattributes(sDocNo, sDocName, sAttributeNames1, sAttributeValues1, sAttributeTypes1, sCheckInComments, Convert.ToInt16(sWebAppId));
                            }
                            else
                            {
                                string[] sAttributeNames1 = new string[2];
                                string[] sAttributeValues1 = new string[2];
                                string[] sAttributeTypes1 = new string[2];
                                sAttributeNames1[0] = "DispatchDocketDate";
                                sAttributeNames1[1] = "Comments";

                                sAttributeValues1[0] = sProdNoDate;
                                sAttributeValues1[1] = sComments;

                                sAttributeTypes1[0] = "date";
                                sAttributeTypes1[1] = "string";

                                sRtn = client2.setdocattributes(sDocNo, sDocName, sAttributeNames1, sAttributeValues1, sAttributeTypes1, sCheckInComments, Convert.ToInt16(sWebAppId));
                            }


                        }
                        else
                        {
                            if (sTargetQty != "")
                            {
                                string[] sAttributeNames1 = new string[2];
                                string[] sAttributeValues1 = new string[2];
                                string[] sAttributeTypes1 = new string[2];
                                sAttributeNames1[0] = "DispatchDocketDate";
                                sAttributeNames1[1] = "TargetQty";

                                sAttributeValues1[0] = sProdNoDate;
                                sAttributeValues1[1] = sTargetQty;

                                sAttributeTypes1[0] = "date";
                                sAttributeTypes1[1] = "double";

                                sRtn = client2.setdocattributes(sDocNo, sDocName, sAttributeNames1, sAttributeValues1, sAttributeTypes1, sCheckInComments, Convert.ToInt16(sWebAppId));
                            }
                            else
                            {
                                //You must have a date. The javascript ensures this.
                                string[] sAttributeNames1 = new string[1];
                                string[] sAttributeValues1 = new string[1];
                                string[] sAttributeTypes1 = new string[1];
                                sAttributeNames1[0] = "DispatchDocketDate";

                                sAttributeValues1[0] = sProdNoDate;

                                sAttributeTypes1[0] = "date";

                                sRtn = client2.setdocattributes(sDocNo, sDocName, sAttributeNames1, sAttributeValues1, sAttributeTypes1, sCheckInComments, Convert.ToInt16(sWebAppId));
                            }
                        }
                    }
                }
                return sRtn1;
            }
        }

        public string CreateCableSchedule(string sSessionId, string sUserId, string sDocNo, string sDocName, string sProductName, string sDocType, string sFolderNameAndPath,
                                          string sOriginator, string sJobCode, string sComments, string sRevision, string sCheckInComments, string iProdOrLibrary, string sWebAppId)
        {

            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                string sRtn = "";
                Update_User_Time(sUserId, sSessionId);
                int iiProdOrLibrary = Convert.ToInt16(iProdOrLibrary);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                string[] sAttributeNames = new string[2];
                string[] sAttributeValues = new string[2];
                string[] sAttributeTypes = new string[2];

                sAttributeNames[0] = "JobCode";
                sAttributeNames[1] = "Originator";

                sAttributeValues[0] = sJobCode;
                sAttributeValues[1] = sOriginator;

                string sRtn1 = client2.doccreate2(sDocNo, sDocName, sProductName, sDocType, sFolderNameAndPath, sRevision, sAttributeNames, sAttributeValues, sCheckInComments, iiProdOrLibrary, Convert.ToInt16(sWebAppId));

                if (sRtn1.StartsWith("Success"))
                {

                    //Get the new document number
                    string[] sSuccess = Extract_Values(sRtn1);

                    sDocNo = sSuccess[1];

                    if (sComments != "")
                    {

                        string[] sAttributeNames1 = new string[1];
                        string[] sAttributeValues1 = new string[1];
                        string[] sAttributeTypes1 = new string[1];
                        sAttributeNames1[0] = "Comments";

                        sAttributeValues1[0] = sComments;

                        sAttributeTypes1[0] = "string";

                        sRtn = client2.setdocattributes(sDocNo, sDocName, sAttributeNames1, sAttributeValues1, sAttributeTypes1, sCheckInComments, Convert.ToInt16(sWebAppId));
                    }

                }
                return sRtn1;
            }
        }

        public string CreateCableScheduleItem(string sSessionId, string sUserId, string sCSNo, string sProductName, string sFolderNameAndPath,
                                              string sCableNo, string sCableName, string sFromFL, string sToFL, string dLength, string sFromLineNumber,
                                              string sToLineNumber, string sMaterialCableCode,
                                              string sOriginator, string sCableComments, string sCableCheckInComments, string iProdOrLibrary, string sWebAppId)
        {

            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ArrayList arrUser = GetUserDetails(sUserId);
                string sFullName = arrUser[2].ToString();

                int iiProdOrLibrary = Convert.ToInt16(iProdOrLibrary);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                string[] sAttributeNames = new string[2];
                string[] sAttributeValues = new string[2];
                string[] sAttributeTypes = new string[2];
                string[] sAttributeNames2 = new string[1];
                string[] sAttributeValues2 = new string[1];
                string[] sAttributeTypes2 = new string[1];

                sAttributeNames[0] = "Originator";
                sAttributeNames[1] = "Comments";

                sAttributeValues[0] = sOriginator;
                sAttributeValues[1] = sCableComments;

                sAttributeTypes[0] = "string";
                sAttributeTypes[1] = "string";

                sAttributeNames2[0] = "ToOrFrom";
                sAttributeValues2[0] = "0"; //0 = from, 1 = to
                sAttributeTypes2[0] = "string";

                string sRtn1 = client2.createpart(sCableNo, sCableName, sProductName, "local.rs.vsrs05.Regain.CablePart", sFolderNameAndPath, sFullName,
                                                  sAttributeNames, sAttributeValues, sAttributeTypes, sCableCheckInComments, iiProdOrLibrary, Convert.ToInt16(sWebAppId));

                if (sRtn1.StartsWith("Success"))
                {
                    //Now reference it to the cable schedule
                    string sRtn2 = client2.setdoctopartref(sUserId, sCSNo, sCableNo, "Creating reference to cable schedule " + sCSNo, "wt.part.WTPartReferenceLink", Convert.ToInt16(sWebAppId));


                    if (sRtn2.StartsWith("Success"))
                    {
                        //Now create a parent child link to the cable material item
                        string sRtn3 = client2.setpartpartlink(sUserId, sCableNo, sMaterialCableCode, Convert.ToDouble(dLength), "Creating link to cable material item " + sMaterialCableCode,
                                                               "wt.part.WTPartUsageLink", "m", Convert.ToInt16(sWebAppId));

                        if (sRtn3.StartsWith("Success"))
                        {
                            //Now create a parent child link between the from functional location and the cable item
                            string sRtn4 = client2.setpartpartlinkwithattributes(sUserId, sFromFL, sCableNo, 1, "Creating link from the From functional location to the cable " + sCableNo,
                                                                   "local.rs.vsrs05.Regain.CableUsage", "ea", Convert.ToInt32(sFromLineNumber), sAttributeNames2, sAttributeValues2, sAttributeTypes2, Convert.ToInt16(sWebAppId));

                            if (sRtn4.StartsWith("Success"))
                            {
                                //Now create a parent child link between the to functional location and the cable item
                                sAttributeValues2[0] = "1"; //0 = from, 1 = to
                                string sRtn5 = client2.setpartpartlinkwithattributes(sUserId, sToFL, sCableNo, 1, "Creating link from the To functional location to the cable " + sCableNo,
                                                                       "local.rs.vsrs05.Regain.CableUsage", "ea", Convert.ToInt32(sToLineNumber), sAttributeNames2, sAttributeValues2, sAttributeTypes2, Convert.ToInt16(sWebAppId));

                                sRtn1 = sRtn5; //This happens regardless of success or not
                            }
                            else
                                sRtn1 = sRtn4;
                        }
                        else
                            sRtn1 = sRtn3;
                    }
                    else
                        sRtn1 = sRtn2;
                }

                return sRtn1;
            }
        }

        public string CreateCableItem(string sSessionId, string sUserId, string sProductName, string sFolderNameAndPath,
                                      string sCableNo, string sCableName, string sFromFL, string sToFL, string dLength, string sFromLineNumber,
                                      string sToLineNumber, string sMaterialCableCode,
                                      string sOriginator, string sCableComments, string sCableCheckInComments, string iProdOrLibrary, string sWebAppId)
        {

            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ArrayList arrUser = GetUserDetails(sUserId);
                string sFullName = arrUser[2].ToString();

                int iiProdOrLibrary = Convert.ToInt16(iProdOrLibrary);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                string[] sAttributeNames = new string[2];
                string[] sAttributeValues = new string[2];
                string[] sAttributeTypes = new string[2];
                string[] sAttributeNames2 = new string[1];
                string[] sAttributeValues2 = new string[1];
                string[] sAttributeTypes2 = new string[1];
                string[] sAttributeNames3 = new string[1];
                string[] sAttributeValues3 = new string[1];
                string[] sAttributeTypes3 = new string[1];

                sAttributeNames[0] = "Originator";
                sAttributeNames[1] = "Comments";

                sAttributeValues[0] = sOriginator;
                sAttributeValues[1] = sCableComments;

                sAttributeTypes[0] = "string";
                sAttributeTypes[1] = "string";

                sAttributeNames2[0] = "ToOrFrom";
                sAttributeValues2[0] = "0"; //0 = from, 1 = to
                sAttributeTypes2[0] = "string";

                sAttributeNames3[0] = "Originator";
                sAttributeValues3[0] = sOriginator;
                sAttributeTypes3[0] = "string";

                string sRtn = "Success";
                string sRtn1 = "Success";

                bool bCableExists = PartExists(sCableNo, Convert.ToInt32(sWebAppId));
                if (!bCableExists)
                {
                    sRtn1 = client2.createpart(sCableNo, sCableName, sProductName, "local.rs.vsrs05.Regain.CablePart", sFolderNameAndPath, sFullName,
                                                        sAttributeNames, sAttributeValues, sAttributeTypes, sCableCheckInComments, iiProdOrLibrary, Convert.ToInt16(sWebAppId));
                    if (!sRtn1.StartsWith("Success"))
                    {
                        return sRtn1;
                    }
                }
                else
                {
                    sRtn1 = SetPartState(sSessionId, sUserId, sCableNo, "InWork", sWebAppId);
                }

                if (!sMaterialCableCode.Equals(""))
                {
                    //Now create a parent child link to the cable material item
                    int iCabMatLineNumber = GetNewLineNumber(sCableNo, Convert.ToInt32(sWebAppId));
                    string sRtn3 = client2.setpartpartlinkwithattributes(sUserId, sCableNo, sMaterialCableCode, Convert.ToDouble(dLength), "Creating link to cable material item " + sMaterialCableCode,
                                                               "wt.part.WTPartUsageLink", "m", iCabMatLineNumber, sAttributeNames3, sAttributeValues3, sAttributeTypes3, Convert.ToInt16(sWebAppId));
                    if (!sRtn3.StartsWith("Success"))
                    {
                        return sRtn3;
                    }
                }

                if (!sFromFL.Equals(""))
                {
                    //Now create a parent child link between the from functional location and the cable item
                    string sRtn4 = client2.setpartpartlinkwithattributes(sUserId, sFromFL, sCableNo, 1, "Creating link from the From functional location " + sFromFL + " to the cable " + sCableNo,
                                                            "local.rs.vsrs05.Regain.CableUsage", "ea", Convert.ToInt32(sFromLineNumber), sAttributeNames2, sAttributeValues2, sAttributeTypes2, Convert.ToInt16(sWebAppId));
                    if (!sRtn4.StartsWith("Success"))
                    {
                        return sRtn4;
                    }
                }

                if (!sToFL.Equals(""))
                {
                    //Now create a parent child link between the to functional location and the cable item
                    sAttributeValues2[0] = "1"; //0 = from, 1 = to
                    string sRtn5 = client2.setpartpartlinkwithattributes(sUserId, sToFL, sCableNo, 1, "Creating link from the To functional location " + sToFL + " to the cable " + sCableNo,
                                                            "local.rs.vsrs05.Regain.CableUsage", "ea", Convert.ToInt32(sToLineNumber), sAttributeNames2, sAttributeValues2, sAttributeTypes2, Convert.ToInt16(sWebAppId));
                    if (!sRtn5.StartsWith("Success"))
                    {
                        return sRtn5;
                    }
                }


                return sRtn;
            }
        }

        public string UpdateCableItem(string sSessionId, string sUserId, string sCableNo, string sCableName,
                                      string sOriginator, string sCableComments, string sCableCheckInComments, string sWebAppId)
        {

            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ArrayList arrUser = GetUserDetails(sUserId);

                string sFullName = arrUser[2].ToString();

                ExampleService.MyJavaService3Client client2 = GetWCService();
                string[] sAttributeNames = new string[2];
                string[] sAttributeValues = new string[2];
                string[] sAttributeTypes = new string[2];

                sAttributeNames[0] = "Originator";
                sAttributeNames[1] = "Comments";

                sAttributeValues[0] = sOriginator;
                sAttributeValues[1] = sCableComments;

                sAttributeTypes[0] = "string";
                sAttributeTypes[1] = "string";

                string sRtn1 = client2.setpartattributes(sCableNo, sCableName, sFullName, sAttributeNames, sAttributeValues, sAttributeTypes, sCableCheckInComments, Convert.ToInt16(sWebAppId));


                return sRtn1;
            }
        }

        public string UpdateCableMaterial(string sSessionId, string sUserId, string sFullName, string sCableNo,
                                          string sLength, string sMaterialCode, string sOldMaterialCode,
                                          string sCableCheckInComments, string sWebAppId)
        {

            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                int iWebAppId = Convert.ToInt32(sWebAppId);

                string[] sAttributeNames3 = new string[1];
                string[] sAttributeValues3 = new string[1];
                string[] sAttributeTypes3 = new string[1];

                sAttributeNames3[0] = "Originator";
                sAttributeValues3[0] = sFullName;
                sAttributeTypes3[0] = "string";

                string sRtn1 = "Success";
                if (!sOldMaterialCode.Equals(""))
                {
                    sRtn1 = DeletePartToPartLink(sSessionId, sUserId, sFullName, sCableNo, sOldMaterialCode, sCableCheckInComments, sWebAppId);
                }

                if (sRtn1.StartsWith("Success"))
                {
                    //Now create a parent child link to the cable material item
                    int iCabMatLineNumber = GetNewLineNumber(sCableNo, Convert.ToInt32(sWebAppId));
                    string sRtn3 = client2.setpartpartlinkwithattributes(sUserId, sCableNo, sMaterialCode, Convert.ToDouble(sLength), "Creating link to cable material item " + sMaterialCode,
                                                           "wt.part.WTPartUsageLink", "m", iCabMatLineNumber, sAttributeNames3, sAttributeValues3, sAttributeTypes3, Convert.ToInt16(sWebAppId));
                    sRtn1 = sRtn3; //This happens regardless of success or not                
                }

                //Now if the number of core has descreased we need to remove the excess cores from the cable
                rtnInt rntCoresOld = GetPartIntAttribute(sOldMaterialCode, "NoOfCores", iWebAppId);
                rtnInt rntCoresNew = GetPartIntAttribute(sMaterialCode, "NoOfCores", iWebAppId);

                if(rntCoresOld.iReturnValue > rntCoresNew.iReturnValue)
                {
                    for(int i = rntCoresOld.iReturnValue; i> rntCoresNew.iReturnValue; i--)
                    {
                        string sCore = sCableNo + "-" + i.ToString().PadLeft(2, '0');
                        sCableCheckInComments = "Removing link between cable " + sCableNo + " and core " + sCore;
                        DeletePartToPartLink(sSessionId, sUserId, sFullName, sCableNo, sCore, sCableCheckInComments, sWebAppId);
                    }
                }

                return sRtn1;
            }
        }


        public string CreateCablePartLink(string sSessionId, string sUserId, string sCableNo, string sFuncLoc, string sLineNumber, string sToOrFrom, string sCheckInComments, string sWebAppId)
        {

            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                string[] sAttributeNames2 = new string[1];
                string[] sAttributeValues2 = new string[1];
                string[] sAttributeTypes2 = new string[1];

                sAttributeNames2[0] = "ToOrFrom";
                sAttributeValues2[0] = sToOrFrom; //0 = from, 1 = to
                sAttributeTypes2[0] = "string";

                //Now create a parent child link between the from functional location and the cable item
                string sRtn4 = client2.setpartpartlinkwithattributes(sUserId, sFuncLoc, sCableNo, 1, "Creating link from the From functional location to the cable " + sCableNo,
                                                        "local.rs.vsrs05.Regain.CableUsage", "ea", Convert.ToInt32(sLineNumber), sAttributeNames2, sAttributeValues2, sAttributeTypes2, Convert.ToInt16(sWebAppId));


                return sRtn4;
            }
        }

        public string UpdateCableFromDetails(string sSessionId, string sUserId, string sCableNo, string sNewFuncLoc,  string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                int iWebAppId = Convert.ToInt16(sWebAppId);
                rtnTerms[] rtnbTerminations = GetTerminations(sCableNo, iWebAppId);
                string sRtnTerm;
                int i;
                bool bFailure = false;
                int iTermFromLineNumber = 0;

                if (!rtnbTerminations[0].bReturnValue)
                {
                    return "Failure";
                }
                else
                {
                    for (i = 0; i < rtnbTerminations.Length; i++)
                    {
                        if(!rtnbTerminations[i].sFromTermination.Equals(""))
                        {
                            iTermFromLineNumber = GetNewLineNumber(sNewFuncLoc, iWebAppId);
                            sRtnTerm = UpdateCableTerminationLink2(sSessionId, sUserId, sCableNo, sNewFuncLoc, iTermFromLineNumber.ToString(), "0",
                                                                   rtnbTerminations[i].sFromTermination, rtnbTerminations[i].sWireNo, rtnbTerminations[i].iCoreNo.ToString(),
                                                                   rtnbTerminations[i].sCoreLabel, sWebAppId);
                            if (!sRtnTerm.Equals("Success"))
                            {
                                bFailure = true;
                            }
                        }
                    }

                    if (bFailure)
                        return "Failure";
                    else
                        return "Success";
                }
            }

        }

        public string CreateCableTerminationLink(string sSessionId, string sUserId, string sCableNo, string sFuncLoc, string sLineNumber, string sToOrFrom, string sTermination, string sWireNo, string sCoreNo, string sWebAppId)
        {

            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                ArrayList arrUser = GetUserDetails(sUserId);
                string sFullName = arrUser[2].ToString();

                string[] sAttributeNames = new string[3];
                string[] sAttributeValues = new string[3];
                string[] sAttributeTypes = new string[3];
                string sCheckInComments = "";

                sAttributeNames[0] = "ToOrFrom";
                sAttributeNames[1] = "Termination";
                sAttributeNames[2] = "CoreNo";
                sAttributeValues[0] = sToOrFrom; //0 = from, 1 = to
                sAttributeValues[1] = sTermination;
                sAttributeValues[2] = sCoreNo;
                sAttributeTypes[0] = "string";
                sAttributeTypes[1] = "string";
                sAttributeTypes[2] = "string";


                if (sWireNo != "")
                {
                    Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                    Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                    Array.Resize<string>(ref sAttributeTypes, sAttributeTypes.Length + 1);
                    sAttributeNames[sAttributeNames.Length - 1] = "WireNo";
                    sAttributeValues[sAttributeValues.Length - 1] = sWireNo;
                    sAttributeTypes[sAttributeTypes.Length - 1] = "string";
                }

                //Now create a parent child link between the from functional location and the cable item
                if (sToOrFrom.Equals("0"))
                    sCheckInComments = "Creating termination link from the 'From' functional location to the cable " + sCableNo + " for termination " + sTermination;
                if (sToOrFrom.Equals("1"))
                    sCheckInComments = "Creating termination link from the 'To' functional location to the cable " + sCableNo + " for termination " + sTermination;

                string sRtn4 = client2.setpartpartlinkwithattributes(sFullName, sFuncLoc, sCableNo, 1, sCheckInComments,
                                                        "local.rs.vsrs05.Regain.TerminationUsage", "ea", Convert.ToInt32(sLineNumber), sAttributeNames, sAttributeValues, sAttributeTypes, Convert.ToInt16(sWebAppId));


                return sRtn4;
            }
        }

        public string CreateCableTerminationLink2(string sSessionId, string sUserId, string sCableNo, string sFuncLoc, string sLineNumber, string sToOrFrom, string sTermination, string sWireNo, string sCoreNo, string sCoreLabel, string sWebAppId)
        {

            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                ArrayList arrUser = GetUserDetails(sUserId);
                string sFullName = arrUser[2].ToString();
                int iWebAppId = 2;
                bool bWireNoExists = false;
                bool bCoreNoExists = false;
                bool bTerminationExists = false;
                string[] sAttributeNamesBlank = new string[0];
                string[] sAttributeValuesBlank = new string[0];
                string[] sAttributeTypesBlank = new string[0];

                iWebAppId = Convert.ToInt32(sWebAppId);

                string[] sAttributeNames = new string[1];
                string[] sAttributeValues = new string[1];
                string[] sAttributeTypes = new string[1];
                string sCheckInComments = "";

                sAttributeNames[0] = "ToOrFrom";
                sAttributeValues[0] = sToOrFrom; //0 = from, 1 = to
                sAttributeTypes[0] = "string";


                if (sWireNo != "")
                {
                    bWireNoExists  = PartExists(sWireNo, iWebAppId);
                    if(!bWireNoExists)
                    {
                        string sProductName = "";
                        string sJob = "";
                        int iProdOrLib = 0;
                        string sFolder = "";

                        if (sFuncLoc.StartsWith("M"))
                        {
                            iProdOrLib = 1;
                            sJob = "M";
                            sFolder = "Material Catalogue/Cable Schedule";
                        }
                        else
                        {
                            int iJob = Convert.ToInt32(sCableNo.Substring(1, 3));
                            sJob = iJob.ToString();
                            rtnString rtnStr3 = GetPlantJobFolder(iJob, iWebAppId);
                            if (rtnStr3.bReturnValue)
                                sFolder = rtnStr3.sReturnValue + "/Cable Schedule";
                        }

                        rtnString rtnStr1 = GetProductFromJob(sJob, iProdOrLib, iWebAppId);

                        if (rtnStr1.bReturnValue)
                            sProductName = rtnStr1.sReturnValue;

/*                        string sFolder = "";
                        rtnString rtnStr2 = GetPlantJobFolder(iJob, iWebAppId);
                        if (rtnStr2.bReturnValue)
                            sFolder = rtnStr2.sReturnValue + "/Cable Schedule";
*/                        
                        sCheckInComments = "Creating Wire No part " + sWireNo;
                        string sRtn1 = client2.createpart(sWireNo, "Wire Number " + sWireNo, sProductName,"local.rs.vsrs05.Regain.CableWire", sFolder, sFullName, 
                                                          sAttributeNamesBlank, sAttributeValuesBlank, sAttributeTypesBlank, sCheckInComments, iProdOrLib, iWebAppId);

                        if (!sRtn1.StartsWith("Success"))
                        {
                            return sRtn1;
                        }
                    }
                }

                sCoreNo = sCableNo + "-" + sCoreNo.PadLeft(2, '0');
                bCoreNoExists = PartExists(sCoreNo, iWebAppId);
                if (!bCoreNoExists)
                {
                    string[] sAttributeNames2 = new string[1];
                    string[] sAttributeValues2 = new string[1];
                    string[] sAttributeTypes2 = new string[1];

                    sAttributeNames2[0] = "CoreLabel";
                    sAttributeValues2[0] = sCoreLabel; //0 = from, 1 = to
                    sAttributeTypes2[0] = "string";

                    string sProductName = "";
                    string sJob = "";
                    int iProdOrLib = 0;
                    string sFolder = "";

                    if (sFuncLoc.StartsWith("M"))
                    {
                        iProdOrLib = 1;
                        sJob = "M";
                        sFolder = "Material Catalogue/Cable Schedule";
                    }
                    else
                    {
                        int iJob = Convert.ToInt32(sCableNo.Substring(1, 3));
                        sJob = iJob.ToString();
                        rtnString rtnStr3 = GetPlantJobFolder(iJob, iWebAppId);
                        if (rtnStr3.bReturnValue)
                            sFolder = rtnStr3.sReturnValue + "/Cable Schedule";
                    }

                    rtnString rtnStr1 = GetProductFromJob(sJob, iProdOrLib, iWebAppId);

                    if (rtnStr1.bReturnValue)
                        sProductName = rtnStr1.sReturnValue;


/*                    string sFolder = "";
                    rtnString rtnStr2 = GetPlantJobFolder(iJob, iWebAppId);
                    if (rtnStr2.bReturnValue)
                        sFolder = rtnStr2.sReturnValue + "/Cable Schedule";
*/                    
                    sCheckInComments = "Creating Core No part " + sCoreNo;
                    string sRtn2 = client2.createpart(sCoreNo, "Core Number " + sCoreNo, sProductName, "local.rs.vsrs05.Regain.CableCore", sFolder, sFullName,
                                                        sAttributeNames2, sAttributeValues2, sAttributeTypes2, sCheckInComments, iProdOrLib, iWebAppId);

                    if (!sRtn2.StartsWith("Success"))
                    {
                        return sRtn2;
                    }

                }

                rtnInt rtnCoreLinkExists = PartPartLinkExists(sCableNo, sCoreNo, iWebAppId);
                bool bCoreLinkExist = rtnCoreLinkExists.bReturnValue;

                if (!bCoreLinkExist)
                {
                    sCheckInComments = "Creating link between cable " + sCableNo + " and the core " + sCoreNo;
                    int iCoreLineNumber = GetNewLineNumber(sCableNo, iWebAppId);
                    string sRtn6 = client2.setpartpartlinkwithattributes(sFullName, sCableNo, sCoreNo, 1, sCheckInComments,
                                                                         "wt.part.WTPartUsageLink", "ea", iCoreLineNumber,
                                                                         sAttributeNamesBlank, sAttributeValuesBlank, sAttributeTypesBlank,
                                                                         Convert.ToInt16(sWebAppId));
                    if (!sRtn6.StartsWith("Success"))
                    {
                        return sRtn6;
                    }
                }


                if (sTermination != "")
                {
                    sTermination = sFuncLoc + "-" + sTermination;
                    bTerminationExists = PartExists(sTermination, iWebAppId);
                    if (!bTerminationExists)
                    {
                        string sProductName = "";
                        string sJob = "";
                        int iProdOrLib = 0;
                        string sFolder = "";

                        if (sFuncLoc.StartsWith("M"))
                        {
                            iProdOrLib = 1;
                            sJob = "M";
                            sFolder = "Material Catalogue/Cable Schedule";
                        }
                        else
                        {
                            int iJob = Convert.ToInt32(sCableNo.Substring(1, 3));
                            sJob = iJob.ToString();
                            rtnString rtnStr3 = GetPlantJobFolder(iJob, iWebAppId);
                            if (rtnStr3.bReturnValue)
                                sFolder = rtnStr3.sReturnValue + "/Cable Schedule";
                        }

                        rtnString rtnStr1 = GetProductFromJob(sJob, iProdOrLib, iWebAppId);

                        if (rtnStr1.bReturnValue)
                            sProductName = rtnStr1.sReturnValue;
/*
                        string sFolder = "";
                        rtnString rtnStr2 = GetPlantJobFolder(iJob, iWebAppId);
                        if (rtnStr2.bReturnValue)
                            sFolder = rtnStr2.sReturnValue + "/Cable Schedule";
*/
                        sCheckInComments = "Creating Termination part " + sTermination;
                        string sRtn1 = client2.createpart(sTermination, "Termination " + sTermination, sProductName, "local.rs.vsrs05.Regain.Termination", sFolder, sFullName,
                                                          sAttributeNamesBlank, sAttributeValuesBlank, sAttributeTypesBlank, sCheckInComments, iProdOrLib, iWebAppId);

                        if (!sRtn1.StartsWith("Success"))
                        {
                            return sRtn1;
                        }
                    }

                    //Add a link between the termination and the parent
                    rtnInt rtnTermLinkExists = PartPartLinkExists(sFuncLoc, sTermination, iWebAppId);
                    bool bTermLinkExist = rtnTermLinkExists.bReturnValue;

                    if (!bTermLinkExist)
                    {
                        sCheckInComments = "Creating link between functional location " + sFuncLoc + " and the termination " + sTermination;
                        int iTermLineNumber = GetNewLineNumber(sFuncLoc, iWebAppId);
                        string sRtn3 = client2.setpartpartlinkwithattributes(sFullName, sFuncLoc, sTermination, 1, sCheckInComments,
                                                                             "wt.part.WTPartUsageLink", "ea", iTermLineNumber,
                                                                             sAttributeNamesBlank, sAttributeValuesBlank, sAttributeTypesBlank,
                                                                             Convert.ToInt16(sWebAppId));

                        if (!sRtn3.StartsWith("Success"))
                        {
                            return sRtn3;
                        }
                    }

                    //Now create a parent child link between the termination and the cable core item
                    if (sToOrFrom.Equals("0"))
                        sCheckInComments = "Creating termination link from the 'From' termination functional location " + sTermination + " to the cable " + sCableNo + " core " + sCoreNo;
                    if (sToOrFrom.Equals("1"))
                        sCheckInComments = "Creating termination link from the 'To' termination functional location " + sTermination + " to the cable " + sCableNo + " core " + sCoreNo;

                    string sRtn4 = client2.setpartpartlinkwithattributes(sFullName, sTermination, sCoreNo, 1, sCheckInComments,
                                                                         "local.rs.vsrs05.Regain.CableUsage", "ea", Convert.ToInt32(sLineNumber),
                                                                         sAttributeNames, sAttributeValues, sAttributeTypes, Convert.ToInt16(sWebAppId));

                    if (!sRtn4.StartsWith("Success"))
                    {
                        return sRtn4;
                    }
                }

                //Does the link between the core and the wire no exists
                rtnInt rtnint = PartPartLinkExists(sCoreNo, sWireNo, iWebAppId);

                if(!rtnint.bReturnValue)
                {
                    int iWireLineNumber = GetNewLineNumber(sCoreNo, iWebAppId);
                    string sRtn3 = client2.setpartpartlinkwithattributes(sFullName, sCoreNo, sWireNo, 1, sCheckInComments,
                                                                         "wt.part.WTPartUsageLink", "ea", iWireLineNumber,
                                                                         sAttributeNamesBlank, sAttributeValuesBlank, sAttributeTypesBlank, Convert.ToInt16(sWebAppId));

                    if (!sRtn3.StartsWith("Success"))
                    {
                        return sRtn3;
                    }

                }



                return "Success";
            }
        }

        
        public string UpdateCableTerminationLink(string sSessionId, string sUserId, string sCableNo, string sFuncLoc, string sLineNumber, string sToOrFrom, string sTermination, string sWireNo, string sCoreNo, string sWebAppId)
        {

            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                ArrayList arrUser = GetUserDetails(sUserId);
                string sFullName = arrUser[2].ToString();

                string[] sAttributeNames = new string[3];
                string[] sAttributeValues = new string[3];
                string[] sAttributeTypes = new string[3];
                string sCheckInComments = "";

                sAttributeNames[0] = "ToOrFrom";
                sAttributeNames[1] = "Termination";
                sAttributeNames[2] = "CoreNo";
                sAttributeValues[0] = sToOrFrom; //0 = from, 1 = to
                sAttributeValues[1] = sTermination;
                sAttributeValues[2] = sCoreNo;
                sAttributeTypes[0] = "string";
                sAttributeTypes[1] = "string";
                sAttributeTypes[2] = "string";


                Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                Array.Resize<string>(ref sAttributeTypes, sAttributeTypes.Length + 1);
                sAttributeNames[sAttributeNames.Length - 1] = "WireNo";
                sAttributeValues[sAttributeValues.Length - 1] = sWireNo;
                sAttributeTypes[sAttributeTypes.Length - 1] = "string";

                //Now create a parent child link between the from functional location and the cable item
                if (sToOrFrom.Equals("0"))
                    sCheckInComments = "Updating termination link from the 'From' functional location to the cable " + sCableNo + " for termination " + sTermination;
                if (sToOrFrom.Equals("1"))
                    sCheckInComments = "Updating termination link from the 'To' functional location to the cable " + sCableNo + " for termination " + sTermination;

                string sRtn4 = client2.updatepartpartlinkwithattributes(sFullName, sFuncLoc, sCableNo, 1, Convert.ToInt32(sLineNumber), sCheckInComments,
                                                        "local.rs.vsrs05.Regain.TerminationUsage", "ea", sAttributeNames, sAttributeValues, sAttributeTypes, Convert.ToInt16(sWebAppId));


                return sRtn4;
            }
        }

        public string UpdateCableTerminationLink2(string sSessionId, string sUserId, string sCableNo, string sFuncLoc, string sLineNumber, string sToOrFrom, string sTermination, 
                                                  string sWireNo, string sCoreNo, string sCoreLabel, string sWebAppId)
        {

            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                ArrayList arrUser = GetUserDetails(sUserId);
                string sFullName = arrUser[2].ToString();
                int iWebAppId = Convert.ToInt16(sWebAppId);
                bool bWireNoExists = false;
                bool bTerminationExists = false;
                bool bCoreNoExists = false;

                string[] sAttributeNamesBlank = new string[0];
                string[] sAttributeValuesBlank = new string[0];
                string[] sAttributeTypesBlank = new string[0];


                string[] sAttributeNames = new string[1];
                string[] sAttributeValues = new string[1];
                string[] sAttributeTypes = new string[1];
                string sCheckInComments = "";

                sAttributeNames[0] = "ToOrFrom";
                sAttributeValues[0] = sToOrFrom; //0 = from, 1 = to
                sAttributeTypes[0] = "string";

                sCoreNo = sCableNo + "-" + sCoreNo.PadLeft(2, '0');

                //Check to see f the from side has changed
/*                if(sToOrFrom.Equals("1"))
                {
                    rtnString rtnFromLink = GetChildPartOfType(sCableNo, "local.rs.vsrs05.Regain.PlantPart", "", iWebAppId);
                    if (rtnFromLink.bReturnValue)
                    {
                        if (!rtnFromLink.sReturnValue.Equals(sFuncLoc))
                        {
                            //Remove the link to the wire no
                            sCheckInComments = "Removing link between cable " + sCableNo + " and existing from side " + rtnFromLink.sReturnValue;
                            string sRtn5 = client2.deletepartpartlinkbylinenumber(sFullName, rtnFromLink.iLineNumber, sCableNo, rtnFromLink.sReturnValue, sCheckInComments, iWebAppId);

                            if (!sRtn5.StartsWith("Success"))
                            {
                                return sRtn5;
                            }

                            sCheckInComments = "Creating link from the 'From' functional location " + sFuncLoc + " to the cable " + sCableNo;
                            int iFromLineNumber = GetNewLineNumber(sFuncLoc, iWebAppId);

                            string sRtn6 = client2.setpartpartlinkwithattributes(sFullName, sFuncLoc, sCableNo, 1, sCheckInComments,
                                                                                 "local.rs.vsrs05.Regain.CableUsage", "ea", iFromLineNumber,
                                                                                 sAttributeNames, sAttributeValues, sAttributeTypes, Convert.ToInt16(sWebAppId));
                            if (!sRtn6.StartsWith("Success"))
                            {
                                return sRtn6;
                            }

                        }

                    }


                }
*/
                if (sTermination != "")
                {
                    sTermination = sFuncLoc + "-" + sTermination;

                    bTerminationExists = PartExists(sTermination, iWebAppId);
                    if (!bTerminationExists)
                    {
                        string sProductName = "";
                        string sJob = "";
                        int iProdOrLib = 0;
                        string sFolder = "";

                        if (sFuncLoc.StartsWith("M"))
                        {
                            iProdOrLib = 1;
                            sJob = "M";
                            sFolder = "Material Catalogue/Cable Schedule";
                        }
                        else
                        {
                            int iJob = Convert.ToInt32(sCableNo.Substring(1, 3));
                            sJob = iJob.ToString();
                            rtnString rtnStr2 = GetPlantJobFolder(iJob, iWebAppId);
                            if (rtnStr2.bReturnValue)
                                sFolder = rtnStr2.sReturnValue + "/Cable Schedule";
                        }

                        rtnString rtnStr1 = GetProductFromJob(sJob, iProdOrLib, iWebAppId);

                        if (rtnStr1.bReturnValue)
                            sProductName = rtnStr1.sReturnValue;
/*
                        string sFolder = "";
                        rtnString rtnStr2 = GetPlantJobFolder(iJob, iWebAppId);
                        if (rtnStr2.bReturnValue)
                            sFolder = rtnStr2.sReturnValue + "/Cable Schedule";
*/                        
                        sCheckInComments = "Creating Termination part " + sTermination;
                        string sRtn1 = client2.createpart(sTermination, "Termination " + sTermination, sProductName, "local.rs.vsrs05.Regain.Termination", sFolder, sFullName,
                                                          sAttributeNamesBlank, sAttributeValuesBlank, sAttributeTypesBlank, sCheckInComments, iProdOrLib, iWebAppId);

                        if (!sRtn1.StartsWith("Success"))
                        {
                            return sRtn1;
                        }
                    }

                    //Add a link between the termination and the parent
                    rtnInt rtnTermLinkExists = PartPartLinkExists(sFuncLoc, sTermination, iWebAppId);
                    bool bTermLinkExist = rtnTermLinkExists.bReturnValue;

                    if (!bTermLinkExist)
                    {
                        sCheckInComments = "Creating link between functional location " + sFuncLoc + " and the termination " + sTermination;
                        int iTermLineNumber = GetNewLineNumber(sFuncLoc, iWebAppId);
                        string sRtn3 = client2.setpartpartlinkwithattributes(sFullName, sFuncLoc, sTermination, 1, sCheckInComments,
                                                                             "wt.part.WTPartUsageLink", "ea", iTermLineNumber,
                                                                             sAttributeNamesBlank, sAttributeValuesBlank, sAttributeTypesBlank, 
                                                                             Convert.ToInt16(sWebAppId));

                    }

                    bCoreNoExists = PartExists(sCoreNo, iWebAppId);
                    if (!bCoreNoExists)
                    {
                        string[] sAttributeNames2 = new string[1];
                        string[] sAttributeValues2 = new string[1];
                        string[] sAttributeTypes2 = new string[1];

                        sAttributeNames2[0] = "CoreLabel";
                        sAttributeValues2[0] = sCoreLabel; //0 = from, 1 = to
                        sAttributeTypes2[0] = "string";

                        string sProductName = "";
                        string sJob = "";
                        int iProdOrLib = 0;
                        string sFolder = "";

                        if (sFuncLoc.StartsWith("M"))
                        {
                            iProdOrLib = 1;
                            sJob = "M";
                            sFolder = "Material Catalogue/Cable Schedule";
                        }
                        else
                        {
                            int iJob = Convert.ToInt32(sCableNo.Substring(1, 3));
                            sJob = iJob.ToString();
                            rtnString rtnStr2 = GetPlantJobFolder(iJob, iWebAppId);
                            if (rtnStr2.bReturnValue)
                                sFolder = rtnStr2.sReturnValue + "/Cable Schedule";
                        }

                        rtnString rtnStr1 = GetProductFromJob(sJob, iProdOrLib, iWebAppId);

                        if (rtnStr1.bReturnValue)
                            sProductName = rtnStr1.sReturnValue;
/*
                        string sFolder = "";
                        rtnString rtnStr2 = GetPlantJobFolder(iJob, iWebAppId);
                        if (rtnStr2.bReturnValue)
                            sFolder = rtnStr2.sReturnValue + "/Cable Schedule";

*/                        sCheckInComments = "Creating Core No part " + sWireNo;
                        string sRtn2 = client2.createpart(sCoreNo, "Core Number " + sCoreNo, sProductName, "local.rs.vsrs05.Regain.CableCore", sFolder, sFullName,
                                                            sAttributeNames2, sAttributeValues2, sAttributeTypes2, sCheckInComments, iProdOrLib, iWebAppId);

                        if (!sRtn2.StartsWith("Success"))
                        {
                            return sRtn2;
                        }

                    }

                    //Get the parent termination functional location
                    rtnString rtn1 = CablePartLinkParent(sCoreNo, Convert.ToInt32(sToOrFrom), iWebAppId);
                    if (rtn1.bReturnValue)
                    {
                        //This means the termination for this core and side has changed
                        if (!rtn1.sReturnValue.Equals(sTermination))
                        {
                            if (sToOrFrom.Equals("0"))
                                sCheckInComments = "Deleting termination link from the 'From' termination functional location " + rtn1.sReturnValue + " to the cable " + sCableNo + " core " + sCoreNo;
                            if (sToOrFrom.Equals("1"))
                                sCheckInComments = "Deleting termination link from the 'To' termination functional location " + rtn1.sReturnValue + " to the cable " + sCableNo + " core " + sCoreNo;
                            string sRtn2 = client2.deletepartpartlinkbylinenumber(sFullName, Convert.ToInt64(rtn1.iLineNumber), rtn1.sReturnValue, sCoreNo, sCheckInComments, iWebAppId);
                            if (!sRtn2.StartsWith("Success"))
                            {
                                return sRtn2;
                            }

                            if (sToOrFrom.Equals("0"))
                                sCheckInComments = "Creating termination link from the 'From' termination functional location " + sTermination + " to the cable " + sCableNo + " core " + sCoreNo;
                            if (sToOrFrom.Equals("1"))
                                sCheckInComments = "Creating termination link from the 'To' termination functional location " + sTermination + " to the cable " + sCableNo + " core " + sCoreNo;

                            string sRtn3 = client2.setpartpartlinkwithattributes(sFullName, sTermination, sCoreNo, 1, sCheckInComments,
                                                                                 "local.rs.vsrs05.Regain.CableUsage", "ea", Convert.ToInt32(sLineNumber),
                                                                                 sAttributeNames, sAttributeValues, sAttributeTypes, Convert.ToInt16(sWebAppId));

                            if (!sRtn3.StartsWith("Success"))
                            {
                                return sRtn3;
                            }
                        }

                    }
                    else
                    {
                        //Now create a parent child link between the termination and the cable core item
                        if (sToOrFrom.Equals("0"))
                            sCheckInComments = "Creating termination link from the 'From' termination functional location " + sTermination + " to the cable " + sCableNo + " core " + sCoreNo;
                        if (sToOrFrom.Equals("1"))
                            sCheckInComments = "Creating termination link from the 'To' termination functional location " + sTermination + " to the cable " + sCableNo + " core " + sCoreNo;

                        string sRtn4 = client2.setpartpartlinkwithattributes(sFullName, sTermination, sCoreNo, 1, sCheckInComments,
                                                                             "local.rs.vsrs05.Regain.CableUsage", "ea", Convert.ToInt32(sLineNumber),
                                                                             sAttributeNames, sAttributeValues, sAttributeTypes, Convert.ToInt16(sWebAppId));

                        if (!sRtn4.StartsWith("Success"))
                        {
                            return sRtn4;
                        }

                    }
                }

                //Remove any existing link to the wire number if it has changed
                rtnString rtnWireLink = GetChildPartOfType(sCoreNo, "local.rs.vsrs05.Regain.CableWire", "", iWebAppId);
                if(rtnWireLink.bReturnValue)
                {
                    if(!rtnWireLink.sReturnValue.Equals(sWireNo) && !rtnWireLink.sReturnValue.Equals(""))
                    {
                        //Remove the link to the wire no
                        sCheckInComments = "Removing link between cable core " + sCoreNo + " and existing wire no " + rtnWireLink.sReturnValue;
                        string sRtn4 = client2.deletepartpartlinkbylinenumber(sFullName, rtnWireLink.iLineNumber, sCoreNo, rtnWireLink.sReturnValue, sCheckInComments, iWebAppId);

                        if (!sRtn4.StartsWith("Success"))
                        {
                            return sRtn4;
                        }
                    }

                }


                if (sWireNo != "")
                {
                    bWireNoExists = PartExists(sWireNo, iWebAppId);
                    if (!bWireNoExists)
                    {
                        string sProductName = "";
                        string sJob = "";
                        int iProdOrLib = 0;
                        string sFolder = "";

                        if (sFuncLoc.StartsWith("M"))
                        {
                            iProdOrLib = 1;
                            sJob = "M";
                            sFolder = "Material Catalogue/Cable Schedule";
                        }
                        else
                        {
                            int iJob = Convert.ToInt32(sCableNo.Substring(1, 3));
                            sJob = iJob.ToString();
                            rtnString rtnStr2 = GetPlantJobFolder(iJob, iWebAppId);
                            if (rtnStr2.bReturnValue)
                                sFolder = rtnStr2.sReturnValue + "/Cable Schedule";
                        }

                        rtnString rtnStr1 = GetProductFromJob(sJob, iProdOrLib, iWebAppId);

                        if (rtnStr1.bReturnValue)
                            sProductName = rtnStr1.sReturnValue;
/*
                        string sFolder = "";
                        rtnString rtnStr2 = GetPlantJobFolder(iJob, iWebAppId);
                        if (rtnStr2.bReturnValue)
                            sFolder = rtnStr2.sReturnValue + "/Cable Schedule";
*/                        
                        sCheckInComments = "Creating Wire No part " + sWireNo;
                        string sRtn1 = client2.createpart(sWireNo, "Wire Number " + sWireNo, sProductName, "local.rs.vsrs05.Regain.CableWire", sFolder, sFullName,
                                                          sAttributeNamesBlank, sAttributeValuesBlank, sAttributeTypesBlank, sCheckInComments, iProdOrLib, iWebAppId);

                        if (!sRtn1.StartsWith("Success"))
                        {
                            return sRtn1;
                        }
                    }

                    //Link the core to the wire number
                    rtnInt rtnWireLinkExists = PartPartLinkExists(sCoreNo, sWireNo, iWebAppId);
                    bool bWireLinkExist = rtnWireLinkExists.bReturnValue;

                    if(!bWireLinkExist)
                    {
                        sCheckInComments = "Creating link between core " + sCoreNo + " and Wire No " + sWireNo;
                        int iNewWireLineNumber = GetNewLineNumber(sCoreNo, iWebAppId);
                        string sRtn3 = client2.setpartpartlinkwithattributes(sFullName, sCoreNo, sWireNo, 1, sCheckInComments,
                                                         "wt.part.WTPartUsageLink", "ea", iNewWireLineNumber,
                                                         sAttributeNamesBlank, sAttributeValuesBlank, sAttributeTypesBlank, Convert.ToInt16(sWebAppId));

                        if (!sRtn3.StartsWith("Success"))
                        {
                            return sRtn3;
                        }
                    }

                }


                rtnInt rtnCoreLinkExists = PartPartLinkExists(sCableNo, sCoreNo, iWebAppId);
                bool bCoreLinkExist = rtnCoreLinkExists.bReturnValue;

                if (!bCoreLinkExist)
                {
                    sCheckInComments = "Creating link between cable " + sCableNo + " and the core " + sCoreNo;
                    int iCoreLineNumber = GetNewLineNumber(sCableNo, iWebAppId);
                    string sRtn6 = client2.setpartpartlinkwithattributes(sFullName, sCableNo, sCoreNo, 1, sCheckInComments,
                                                                         "wt.part.WTPartUsageLink", "ea", iCoreLineNumber,
                                                                         sAttributeNamesBlank, sAttributeValuesBlank, sAttributeTypesBlank,
                                                                         Convert.ToInt16(sWebAppId));
                    if (!sRtn6.StartsWith("Success"))
                    {
                        return sRtn6;
                    }
                }

                rtnString rtnStr3 = GetCableCoreLabel(sCoreNo, iWebAppId);

                if (rtnStr3.bReturnValue)
                {
                    if(!rtnStr3.sReturnValue.Equals(sCoreLabel))
                    {
                        //Also update the core with the latest label
                        string[] sAttributeNames2 = new string[1];
                        string[] sAttributeValues2 = new string[1];
                        string[] sAttributeTypes2 = new string[1];

                        sAttributeNames2[0] = "CoreLabel";
                        sAttributeValues2[0] = sCoreLabel; //0 = from, 1 = to
                        sAttributeTypes2[0] = "string";
                        sCheckInComments = "Updating core label on core " + sCoreNo;
                        string sRtn5 = client2.setpartattributes(sCoreNo, "Core Number " + sCoreNo, sFullName, sAttributeNames2, sAttributeValues2, sAttributeTypes2, sCheckInComments, iWebAppId);

                        if (!sRtn5.StartsWith("Success"))
                        {
                            return sRtn5;
                        }
                    }
                }
                else
                {
                    return "Cannot find core " + sCoreNo;
                }

                return "Success";
            }
        }

        public string CreateTestAndTagItem(string sSessionId, string sUserId, string sProductName, string sFolderNameAndPath, string sGroupNo,
                                           string sTestAndTagItemNo, string sTestAndTagName, string sTestAndTagDate, string sTestAndTagResult, 
                                           string sTestAndTagTagNumber, string sTestAndTagMaintenanceActionNo, string sCommonActionNo, 
                                           string sNextTestDate, string sWebAppId)
        {

            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ArrayList arrUser = GetUserDetails(sUserId);
                string sFullName = arrUser[2].ToString();

                ExampleService.MyJavaService3Client client2 = GetWCService();
                string[] sAttributeNames = new string[5];
                string[] sAttributeValues = new string[5];
                string[] sAttributeTypes = new string[5];
                string[] sAttributeNames2 = new string[0];
                string[] sAttributeValues2 = new string[0];
                string[] sAttributeTypes2 = new string[0];

                sAttributeNames[0] = "TestDate";
                sAttributeNames[1] = "TestResult";
                sAttributeNames[2] = "TagNumber";
                sAttributeNames[3] = "NextTestDate";
                sAttributeNames[4] = "Frequency";

                sAttributeValues[0] = sTestAndTagDate;
                sAttributeValues[1] = sTestAndTagResult;
                sAttributeValues[2] = sTestAndTagTagNumber;
                sAttributeValues[3] = sNextTestDate;
                sAttributeValues[4] = "3";

                sAttributeTypes[0] = "date";
                sAttributeTypes[1] = "string";
                sAttributeTypes[2] = "string";
                sAttributeTypes[3] = "date";
                sAttributeTypes[4] = "string";

                string sCheckedInComments = "Creating test and tag item " + sTestAndTagItemNo;

                string sRtn1 = client2.createpart(sTestAndTagItemNo, sTestAndTagName, sProductName, "local.rs.vsrs05.Regain.TestAndTagItem", sFolderNameAndPath, sFullName,
                                                  sAttributeNames, sAttributeValues, sAttributeTypes, sCheckedInComments, 0, Convert.ToInt16(sWebAppId));

                if (sRtn1.StartsWith("Success"))
                {
                    //Now create a parent child link from the group to the new test and tag item
                    string sRtn2 = client2.setpartpartlink(sUserId, sGroupNo, sTestAndTagItemNo, 1.0, "Creating link to test and tag item " + sTestAndTagItemNo,
                                                               "wt.part.WTPartUsageLink", "ea", Convert.ToInt16(sWebAppId));


                    if (sRtn2.StartsWith("Success"))
                    {
                        //Now create a part for the planned action
                        sCheckedInComments = "Creating test and tag planned maintenance action " + sTestAndTagMaintenanceActionNo;
                        string sRtn3 = client2.createpart(sTestAndTagMaintenanceActionNo, "Test and Tag - " + sTestAndTagName, sProductName, "local.rs.vsrs05.Regain.PreventiveAction", sFolderNameAndPath, sFullName,
                                                          sAttributeNames2, sAttributeValues2, sAttributeTypes2, sCheckedInComments, 0, Convert.ToInt16(sWebAppId));

                        if (sRtn3.StartsWith("Success"))
                        {
                            //Now create a parent child link between the test and tag and the planned action
                            string sRtn4 = client2.setpartpartlink(sUserId, sTestAndTagItemNo, sTestAndTagMaintenanceActionNo, 1, 
                                                                   "Creating link from the test and tag item " + sTestAndTagItemNo + " and the planned maintenance action " + sTestAndTagMaintenanceActionNo,
                                                                   "wt.part.WTPartUsageLink", "ea",Convert.ToInt16(sWebAppId));

                            if (sRtn4.StartsWith("Success"))
                            {
                                //Now create a parent child link between the planned action and the common action 
                                string sRtn5 = client2.setpartpartlink(sUserId, sTestAndTagMaintenanceActionNo, sCommonActionNo, 1,
                                                                       "Creating link from the  planned maintenance action " + sTestAndTagMaintenanceActionNo + " and the common maintenance action " + sCommonActionNo,
                                                                       "wt.part.WTPartUsageLink", "ea", Convert.ToInt16(sWebAppId));


                                if (sRtn5.StartsWith("Success"))
                                {
                                    //Now create a parent child link between the planned action group and the planned maintenance action 
                                    rtnString rtnStr = GetChildPartOfType(sGroupNo, "local.rs.vsrs05.Regain.PlannedActionGroup", "Test and Tag", Convert.ToInt16(sWebAppId));
                                    string sPlannedActionGroup = sGroupNo.Replace("P", "S") + "_01";
                                    if (rtnStr.bReturnValue)
                                        sPlannedActionGroup = rtnStr.sReturnValue;

                                    string sRtn6 = client2.setpartpartlink(sUserId, sPlannedActionGroup, sTestAndTagMaintenanceActionNo, 1,
                                                                           "Creating link from the planned action group " + sPlannedActionGroup + " and the planned maintenance action " + sTestAndTagMaintenanceActionNo,
                                                                           "wt.part.WTPartUsageLink", "ea", Convert.ToInt16(sWebAppId));


                                    sRtn1 = sRtn6; //This happens regardless of success or not
                                }
                                else
                                    sRtn1 = sRtn5;
                            }
                            else
                                sRtn1 = sRtn4;
                        }
                        else
                            sRtn1 = sRtn3;
                    }
                    else
                        sRtn1 = sRtn2;
                }

                return sRtn1;
            }
        }

        public string UpdateTestAndTagItem(string sSessionId, string sUserId, 
                                           string sTestAndTagItemNo, string sTestAndTagName, string sTestAndTagDate, string sTestAndTagResult,
                                           string sTestAndTagTagNumber, 
                                           string sNextTestDate, string sWebAppId)
        {

            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ArrayList arrUser = GetUserDetails(sUserId);
                string sFullName = arrUser[2].ToString();

                ExampleService.MyJavaService3Client client2 = GetWCService();
                string[] sAttributeNames = new string[5];
                string[] sAttributeValues = new string[5];
                string[] sAttributeTypes = new string[5];

                sAttributeNames[0] = "TestDate";
                sAttributeNames[1] = "TestResult";
                sAttributeNames[2] = "TagNumber";
                sAttributeNames[3] = "NextTestDate";
                sAttributeNames[4] = "Frequency";

                sAttributeValues[0] = sTestAndTagDate;
                sAttributeValues[1] = sTestAndTagResult;
                sAttributeValues[2] = sTestAndTagTagNumber;
                sAttributeValues[3] = sNextTestDate;
                sAttributeValues[4] = "3";

                sAttributeTypes[0] = "date";
                sAttributeTypes[1] = "string";
                sAttributeTypes[2] = "string";
                sAttributeTypes[3] = "date";
                sAttributeTypes[4] = "string";

                string sCheckedInComments = "Updating test and tag item " + sTestAndTagItemNo;

                string sRtn1 = client2.setpartattributes(sTestAndTagItemNo, sTestAndTagName, sFullName,
                                                  sAttributeNames, sAttributeValues, sAttributeTypes, sCheckedInComments, Convert.ToInt16(sWebAppId));

                return sRtn1;
            }
        }

        public string DeleteTestAndTagItem(string sSessionId, string sUserId, string sFullName, string sGroupNo, string sTestAndTagItemNo, string sCheckInComments, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                string sRtn1 = client2.deletepartpartlink(sFullName, sGroupNo, sTestAndTagItemNo, sCheckInComments, Convert.ToInt16(sWebAppId));

                //                string sPlannedActionGroup = sGroupNo.Replace("P", "S") + "_01";

                if (sRtn1.StartsWith("Success"))
                {
                    rtnString rtnStr2 = GetChildPartOfType(sTestAndTagItemNo, "local.rs.vsrs05.Regain.PreventiveAction", "Test and Tag", Convert.ToInt16(sWebAppId));

                    if (rtnStr2.bReturnValue)
                    {
                        rtnString rtnStr = GetParentPartOfType(rtnStr2.sReturnValue, "local.rs.vsrs05.Regain.PlannedActionGroup", Convert.ToInt16(sWebAppId));

                        if (rtnStr.bReturnValue)
                        {
                            sRtn1 = client2.deletepartpartlink(sFullName, rtnStr.sReturnValue, rtnStr2.sReturnValue, sCheckInComments, Convert.ToInt16(sWebAppId));
                        }
                    }
                    sRtn1 = SetPartState(sSessionId, sUserId, sTestAndTagItemNo, "Obsolete", sWebAppId);
                }

                return sRtn1;
            }
        }

        public string CreateMaterialCatalogItem(string sSessionId, string sUserId, string sFullName, string sMatCatNo, string sMatCatType, string sName, string sDesc, string sLongDesc,
                                                string sDrivekW, string sFullLoadCurrent, string sUnitWeight, string sLeadTime, string sRepairable, string sCheckInComments, string sWebAppId)
        {
            string sReturn = "";
            string sReturn3 = "";
            string sFolder = "Material Catalogue/";

            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                string[] sAttributeNames = new string[3];
                string[] sAttributeValues = new string[3];
                string[] sAttributeTypes = new string[3];

                sAttributeNames[0] = "Originator";
                sAttributeNames[1] = "PartDesc";
                sAttributeNames[2] = "LongDescription";

                sAttributeValues[0] = sFullName;
                sAttributeValues[1] = sDesc;
                sAttributeValues[2] = sLongDesc;

                sAttributeTypes[0] = "string";
                sAttributeTypes[1] = "string";
                sAttributeTypes[2] = "string";

                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();

                if (sDrivekW != "")
                {
                    Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                    Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                    Array.Resize<string>(ref sAttributeTypes, sAttributeTypes.Length + 1);
                    sAttributeNames[sAttributeNames.Length - 1] = "DrivekW";
                    sAttributeValues[sAttributeValues.Length - 1] = sDrivekW;
                    sAttributeTypes[sAttributeTypes.Length - 1] = "double";
                }

                if (sFullLoadCurrent != "")
                {
                    Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                    Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                    Array.Resize<string>(ref sAttributeTypes, sAttributeTypes.Length + 1);
                    sAttributeNames[sAttributeNames.Length - 1] = "FullLoadCurrent";
                    sAttributeValues[sAttributeValues.Length - 1] = sFullLoadCurrent;
                    sAttributeTypes[sAttributeTypes.Length - 1] = "double";
                }

                if (sUnitWeight != "")
                {
                    Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                    Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                    Array.Resize<string>(ref sAttributeTypes, sAttributeTypes.Length + 1);
                    sAttributeNames[sAttributeNames.Length - 1] = "UnitWeight";
                    sAttributeValues[sAttributeValues.Length - 1] = sUnitWeight;
                    sAttributeTypes[sAttributeTypes.Length - 1] = "double";
                }

                if (sLeadTime != "")
                {
                    Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                    Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                    Array.Resize<string>(ref sAttributeTypes, sAttributeTypes.Length + 1);
                    sAttributeNames[sAttributeNames.Length - 1] = "LeadTime";
                    sAttributeValues[sAttributeValues.Length - 1] = sLeadTime;
                    sAttributeTypes[sAttributeTypes.Length - 1] = "double";
                }

                if (sRepairable != "")
                {
                    Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                    Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                    Array.Resize<string>(ref sAttributeTypes, sAttributeTypes.Length + 1);
                    sAttributeNames[sAttributeNames.Length - 1] = "Repairable";
                    sAttributeValues[sAttributeValues.Length - 1] = sRepairable;
                    sAttributeTypes[sAttributeTypes.Length - 1] = "boolean";
                }

                sReturn = client2.createpart("", sName, "Regain Material Catalogue", "local.rs.vsrs05.Regain.AutoNumberedPart", sFolder, sFullName, sAttributeNames, sAttributeValues, sAttributeTypes, sCheckInComments, 1, Convert.ToInt16(sWebAppId));
                if (sReturn.StartsWith("Success"))
                {
                    sMatCatNo = sReturn.Substring(sReturn.IndexOf("^") + 1, (sReturn.Length - sReturn.IndexOf("^") - 2));
                    sReturn3 = client2.setpartpartlink(sFullName, sMatCatType, sMatCatNo, 1, sCheckInComments, "wt.part.WTPartUsageLink", "", Convert.ToInt16(sWebAppId));
                    if (sReturn3 != "Success")
                        sReturn = sReturn3;
                }

                return sReturn;
            }
        }

        public string UpdateMaterialCatalogItem(string sSessionId, string sUserId, string sFullName, string sMatCatNo, string sMatCatNewType, string sMatCatOldType,
                                                string sName, string sDesc, string sLongDesc, string sDrivekW, string sFullLoadCurrent,
                                                string sUnitWeight, string sLeadTime, string sRepairable, string sCheckInComments, string sWebAppId, string sNewLink)
        {
            string sReturn = "";
            string sReturn2 = "";
            string sReturn3 = "";

            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                string[] sAttributeNames = new string[3];
                string[] sAttributeValues = new string[3];
                string[] sAttributeTypes = new string[3];

                sAttributeNames[0] = "Originator";
                sAttributeNames[1] = "PartDesc";
                sAttributeNames[2] = "LongDescription";

                sAttributeValues[0] = sFullName;
                sAttributeValues[1] = sDesc;
                sAttributeValues[2] = sLongDesc;

                sAttributeTypes[0] = "string";
                sAttributeTypes[1] = "string";
                sAttributeTypes[2] = "string";

                Update_User_Time(sUserId, sSessionId);
                bool bNewLink = Convert.ToBoolean(sNewLink);
                ExampleService.MyJavaService3Client client2 = GetWCService();

                Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                Array.Resize<string>(ref sAttributeTypes, sAttributeTypes.Length + 1);
                sAttributeNames[sAttributeNames.Length - 1] = "DrivekW";
                sAttributeValues[sAttributeValues.Length - 1] = sDrivekW;
                sAttributeTypes[sAttributeTypes.Length - 1] = "double";

                Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                Array.Resize<string>(ref sAttributeTypes, sAttributeTypes.Length + 1);
                sAttributeNames[sAttributeNames.Length - 1] = "FullLoadCurrent";
                sAttributeValues[sAttributeValues.Length - 1] = sFullLoadCurrent;
                sAttributeTypes[sAttributeTypes.Length - 1] = "double";

                Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                Array.Resize<string>(ref sAttributeTypes, sAttributeTypes.Length + 1);
                sAttributeNames[sAttributeNames.Length - 1] = "UnitWeight";
                sAttributeValues[sAttributeValues.Length - 1] = sUnitWeight;
                sAttributeTypes[sAttributeTypes.Length - 1] = "double";

                Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                Array.Resize<string>(ref sAttributeTypes, sAttributeTypes.Length + 1);
                sAttributeNames[sAttributeNames.Length - 1] = "LeadTime";
                sAttributeValues[sAttributeValues.Length - 1] = sLeadTime;
                sAttributeTypes[sAttributeTypes.Length - 1] = "double";

                Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                Array.Resize<string>(ref sAttributeTypes, sAttributeTypes.Length + 1);
                sAttributeNames[sAttributeNames.Length - 1] = "Repairable";
                sAttributeValues[sAttributeValues.Length - 1] = sRepairable;
                sAttributeTypes[sAttributeTypes.Length - 1] = "boolean";

                sReturn = client2.setpartattributes(sMatCatNo, sName, sFullName, sAttributeNames, sAttributeValues, sAttributeTypes, sCheckInComments, Convert.ToInt16(sWebAppId));
                if (sReturn == "Success")
                {
                    if (bNewLink)
                    {

                        sReturn2 = client2.deletepartpartlink(sFullName, sMatCatOldType, sMatCatNo, "Removing link between " + sMatCatOldType + " and " + sMatCatNo, Convert.ToInt16(sWebAppId));
                        if (sReturn2 != "Success")
                            sReturn = sReturn2;
                        else
                        {
                            sReturn3 = client2.setpartpartlink(sFullName, sMatCatNewType, sMatCatNo, 1, "Setting link between " + sMatCatNewType + " and " + sMatCatNo, "wt.part.WTPartUsageLink", "", Convert.ToInt16(sWebAppId));
                            if (sReturn3 != "Success")
                                sReturn = sReturn3;
                        }
                    }
                }

                return sReturn;
            }
        }

        public string SetPartState(string sSessionId, string sUserId, string sPartNo, string sLifecycleState, string sWebAppId)
        {

            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();

                //Now create a parent child link between the from functional location and the cable item
                string sRtn4 = client2.setpartstate(sPartNo, sLifecycleState, Convert.ToInt16(sWebAppId));

                return sRtn4;
            }
        }

        public string CreatePlantEquipItem(string sSessionId, string sUserId, string sFullName, string sPlantEquipNo,
                                           string sPlantEquipType, string sName, string sDesc, string sLongDesc, string sContSysType, string sDriveRating,
                                           string sEquipRegFlag, string sIPRegFlag, string sIPAddress, string sComments, string sOpZone,
                                           string sProduct, string sFolder,
                                           string sPowerCable, string sControlCable, string sInstrumentationCable, 
                                           string sDataCable, string sEarthCable,
                                           string sInstRegFlag, string sFullLoadCurrent, string sConstructionDate, string sFLGrouping,
                                           string sCheckInComments, string sWebAppId)
        {
            string sReturn = "";

            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                string[] sAttributeNames = new string[10];
                string[] sAttributeValues = new string[10];
                string[] sAttributeTypes = new string[10];

                sAttributeNames[0] = "Originator";
                sAttributeNames[1] = "PartDesc";
                sAttributeNames[2] = "LongDescription";
                sAttributeNames[3] = "ContSysFuncType";
                sAttributeNames[4] = "DriveKW";
                sAttributeNames[5] = "EquipRegFlag";
                sAttributeNames[6] = "IPRegFlag";
                sAttributeNames[7] = "IPAddress";
                sAttributeNames[8] = "Comments";
                sAttributeNames[9] = "OpZone";

                sAttributeValues[0] = sFullName;
                sAttributeValues[1] = sDesc;
                sAttributeValues[2] = sLongDesc;
                sAttributeValues[3] = sContSysType;
                sAttributeValues[4] = sDriveRating;
                sAttributeValues[5] = sEquipRegFlag;
                sAttributeValues[6] = sIPRegFlag;
                sAttributeValues[7] = sIPAddress;
                sAttributeValues[8] = sComments;
                sAttributeValues[9] = sOpZone;

                sAttributeTypes[0] = "string";
                sAttributeTypes[1] = "string";
                sAttributeTypes[2] = "string";
                sAttributeTypes[3] = "string";
                sAttributeTypes[4] = "double";
                sAttributeTypes[5] = "boolean";
                sAttributeTypes[6] = "boolean";
                sAttributeTypes[7] = "string";
                sAttributeTypes[8] = "string";
                sAttributeTypes[9] = "string";

                if (sPowerCable != "")
                {
                    Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                    Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                    Array.Resize<string>(ref sAttributeTypes, sAttributeTypes.Length + 1);
                    sAttributeNames[sAttributeNames.Length - 1] = "PowerCable";
                    sAttributeValues[sAttributeValues.Length - 1] = sPowerCable;
                    sAttributeTypes[sAttributeTypes.Length - 1] = "long";
                }

                if (sControlCable != "")
                {
                    Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                    Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                    Array.Resize<string>(ref sAttributeTypes, sAttributeTypes.Length + 1);
                    sAttributeNames[sAttributeNames.Length - 1] = "ControlCable";
                    sAttributeValues[sAttributeValues.Length - 1] = sControlCable;
                    sAttributeTypes[sAttributeTypes.Length - 1] = "long";
                }

                if (sInstrumentationCable != "")
                {
                    Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                    Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                    Array.Resize<string>(ref sAttributeTypes, sAttributeTypes.Length + 1);
                    sAttributeNames[sAttributeNames.Length - 1] = "InstrumentationCable";
                    sAttributeValues[sAttributeValues.Length - 1] = sInstrumentationCable;
                    sAttributeTypes[sAttributeTypes.Length - 1] = "long";
                }

                if (sDataCable != "")
                {
                    Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                    Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                    Array.Resize<string>(ref sAttributeTypes, sAttributeTypes.Length + 1);
                    sAttributeNames[sAttributeNames.Length - 1] = "DataCable";
                    sAttributeValues[sAttributeValues.Length - 1] = sDataCable;
                    sAttributeTypes[sAttributeTypes.Length - 1] = "long";
                }

                if (sEarthCable != "")
                {
                    Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                    Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                    Array.Resize<string>(ref sAttributeTypes, sAttributeTypes.Length + 1);
                    sAttributeNames[sAttributeNames.Length - 1] = "EarthCable";
                    sAttributeValues[sAttributeValues.Length - 1] = sEarthCable;
                    sAttributeTypes[sAttributeTypes.Length - 1] = "long";
                }

                if (sInstRegFlag != "")
                {
                    Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                    Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                    Array.Resize<string>(ref sAttributeTypes, sAttributeTypes.Length + 1);
                    sAttributeNames[sAttributeNames.Length - 1] = "InstRegFlag";
                    sAttributeValues[sAttributeValues.Length - 1] = sInstRegFlag;
                    sAttributeTypes[sAttributeTypes.Length - 1] = "boolean";
                }

                if (sFullLoadCurrent != "")
                {
                    Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                    Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                    Array.Resize<string>(ref sAttributeTypes, sAttributeTypes.Length + 1);
                    sAttributeNames[sAttributeNames.Length - 1] = "FullLoadCurrent";
                    sAttributeValues[sAttributeValues.Length - 1] = sFullLoadCurrent;
                    sAttributeTypes[sAttributeTypes.Length - 1] = "double";
                }

                if (sConstructionDate != "")
                {
                    Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                    Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                    Array.Resize<string>(ref sAttributeTypes, sAttributeTypes.Length + 1);
                    sAttributeNames[sAttributeNames.Length - 1] = "ConstructionDate";
                    sAttributeValues[sAttributeValues.Length - 1] = sConstructionDate;
                    sAttributeTypes[sAttributeTypes.Length - 1] = "string";
                }

                if (sFLGrouping != "")
                {
                    Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                    Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                    Array.Resize<string>(ref sAttributeTypes, sAttributeTypes.Length + 1);
                    sAttributeNames[sAttributeNames.Length - 1] = "FLGroupingFlag";
                    sAttributeValues[sAttributeValues.Length - 1] = sFLGrouping;
                    sAttributeTypes[sAttributeTypes.Length - 1] = "long";
                }

                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();

                sReturn = client2.createpart(sPlantEquipNo, sName, sProduct, sPlantEquipType, sFolder, sFullName, sAttributeNames, sAttributeValues, sAttributeTypes, sCheckInComments, 0, Convert.ToInt16(sWebAppId));
                if (sReturn.StartsWith("Success"))
                {
                    sReturn = "Success";
                }

                return sReturn;
            }
        }

        public string UpdatePlantEquipItem(string sSessionId, string sUserId, string sFullName, string sPlantEquipNo,
                                           string sName, string sDesc, string sLongDesc,
                                           string sContSysType, string sDriveRating, string sEquipRegFlag,
                                           string sIPRegFlag, string sIPAddress, string sComments, string sOpZone,
                                           string sPowerCable, string sControlCable, string sInstrumentationCable, 
                                           string sDataCable, string sEarthCable,
                                           string sInstRegFlag, string sFullLoadCurrent, string sConstructionDate, string sFLGrouping,
                                           string sCheckInComments, string sWebAppId)
        {
            string sReturn = "";

            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                string[] sAttributeNames = new string[10];
                string[] sAttributeValues = new string[10];
                string[] sAttributeTypes = new string[10];

                sAttributeNames[0] = "Originator";
                sAttributeNames[1] = "PartDesc";
                sAttributeNames[2] = "LongDescription";
                sAttributeNames[3] = "ContSysFuncType";
                sAttributeNames[4] = "DriveKW";
                sAttributeNames[5] = "EquipRegFlag";
                sAttributeNames[6] = "IPRegFlag";
                sAttributeNames[7] = "IPAddress";
                sAttributeNames[8] = "Comments";
                sAttributeNames[9] = "OpZone";

                sAttributeValues[0] = sFullName;
                sAttributeValues[1] = sDesc;
                sAttributeValues[2] = sLongDesc;
                sAttributeValues[3] = sContSysType;
                sAttributeValues[4] = sDriveRating;
                sAttributeValues[5] = sEquipRegFlag;
                sAttributeValues[6] = sIPRegFlag;
                sAttributeValues[7] = sIPAddress;
                sAttributeValues[8] = sComments;
                sAttributeValues[9] = sOpZone;

                sAttributeTypes[0] = "string";
                sAttributeTypes[1] = "string";
                sAttributeTypes[2] = "string";
                sAttributeTypes[3] = "string";
                sAttributeTypes[4] = "double";
                sAttributeTypes[5] = "boolean";
                sAttributeTypes[6] = "boolean";
                sAttributeTypes[7] = "string";
                sAttributeTypes[8] = "string";
                sAttributeTypes[9] = "string";

                if (sPowerCable != "")
                {
                    Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                    Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                    Array.Resize<string>(ref sAttributeTypes, sAttributeTypes.Length + 1);
                    sAttributeNames[sAttributeNames.Length - 1] = "PowerCable";
                    sAttributeValues[sAttributeValues.Length - 1] = sPowerCable;
                    sAttributeTypes[sAttributeTypes.Length - 1] = "long";
                }

                if (sControlCable != "")
                {
                    Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                    Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                    Array.Resize<string>(ref sAttributeTypes, sAttributeTypes.Length + 1);
                    sAttributeNames[sAttributeNames.Length - 1] = "ControlCable";
                    sAttributeValues[sAttributeValues.Length - 1] = sControlCable;
                    sAttributeTypes[sAttributeTypes.Length - 1] = "long";
                }

                if (sInstrumentationCable != "")
                {
                    Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                    Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                    Array.Resize<string>(ref sAttributeTypes, sAttributeTypes.Length + 1);
                    sAttributeNames[sAttributeNames.Length - 1] = "InstrumentationCable";
                    sAttributeValues[sAttributeValues.Length - 1] = sInstrumentationCable;
                    sAttributeTypes[sAttributeTypes.Length - 1] = "long";
                }

                if (sDataCable != "")
                {
                    Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                    Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                    Array.Resize<string>(ref sAttributeTypes, sAttributeTypes.Length + 1);
                    sAttributeNames[sAttributeNames.Length - 1] = "DataCable";
                    sAttributeValues[sAttributeValues.Length - 1] = sDataCable;
                    sAttributeTypes[sAttributeTypes.Length - 1] = "long";
                }

                if (sEarthCable != "")
                {
                    Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                    Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                    Array.Resize<string>(ref sAttributeTypes, sAttributeTypes.Length + 1);
                    sAttributeNames[sAttributeNames.Length - 1] = "EarthCable";
                    sAttributeValues[sAttributeValues.Length - 1] = sEarthCable;
                    sAttributeTypes[sAttributeTypes.Length - 1] = "long";
                }


                if (sInstRegFlag != "")
                {
                    Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                    Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                    Array.Resize<string>(ref sAttributeTypes, sAttributeTypes.Length + 1);
                    sAttributeNames[sAttributeNames.Length - 1] = "InstRegFlag";
                    sAttributeValues[sAttributeValues.Length - 1] = sInstRegFlag;
                    sAttributeTypes[sAttributeTypes.Length - 1] = "boolean";
                }

                Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                Array.Resize<string>(ref sAttributeTypes, sAttributeTypes.Length + 1);
                sAttributeNames[sAttributeNames.Length - 1] = "FullLoadCurrent";
                sAttributeValues[sAttributeValues.Length - 1] = sFullLoadCurrent;
                sAttributeTypes[sAttributeTypes.Length - 1] = "double";

                Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                Array.Resize<string>(ref sAttributeTypes, sAttributeTypes.Length + 1);
                sAttributeNames[sAttributeNames.Length - 1] = "ConstructionDate";
                sAttributeValues[sAttributeValues.Length - 1] = sConstructionDate;
                sAttributeTypes[sAttributeTypes.Length - 1] = "string";

                Array.Resize<string>(ref sAttributeNames, sAttributeNames.Length + 1);
                Array.Resize<string>(ref sAttributeValues, sAttributeValues.Length + 1);
                Array.Resize<string>(ref sAttributeTypes, sAttributeTypes.Length + 1);
                sAttributeNames[sAttributeNames.Length - 1] = "FLGroupingFlag";
                sAttributeValues[sAttributeValues.Length - 1] = sFLGrouping;
                sAttributeTypes[sAttributeTypes.Length - 1] = "long";

                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();

                sReturn = client2.setpartattributes(sPlantEquipNo, sName, sFullName, sAttributeNames, sAttributeValues, sAttributeTypes, sCheckInComments, Convert.ToInt16(sWebAppId));

                return sReturn;
            }
        }

        public string SetMaterialIOPartToPartLink(string sSessionId, string sUserId, string sFullName, string sParentPartNo, string sChildPartNo, string sLineNumber, string sIOType, string sIOTag, string sCheckinComments, string sWebAppId)
        {
            string[] sAttributeNames = new string[2];
            string[] sAttributeValues = new string[2];
            string[] sAttributeTypes = new string[2];


            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                long lLineNumber = Convert.ToInt32(sLineNumber);

                sAttributeNames[0] = "IOTag";
                sAttributeNames[1] = "IOType";

                sAttributeValues[0] = sIOTag;
                sAttributeValues[1] = sIOType;

                sAttributeTypes[0] = "string";
                sAttributeTypes[1] = "string";

                return client2.setpartpartlinkwithattributes(sFullName, sParentPartNo, sChildPartNo, 1, sCheckinComments, "wt.part.WTPartUsageLink", "ea", lLineNumber, sAttributeNames, sAttributeValues, sAttributeTypes, Convert.ToInt16(sWebAppId));
            }
        }

        public string UpdateIOPartToPartLink(string sSessionId, string sUserId, string sFullName, string sParentPartNo, string sChildPartNo, string sLineNumber, string sIOType, string sIOTag, string sCheckinComments, string sWebAppId)
        {
            string[] sAttributeNames = new string[2];
            string[] sAttributeValues = new string[2];
            string[] sAttributeTypes = new string[2];


            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                long lLineNumber = Convert.ToInt32(sLineNumber);

                sAttributeNames[0] = "IOTag";
                sAttributeNames[1] = "IOType";

                sAttributeValues[0] = sIOTag;
                sAttributeValues[1] = sIOType;

                sAttributeTypes[0] = "string";
                sAttributeTypes[1] = "string";

                return client2.updatepartpartlinkwithattributes(sFullName, sParentPartNo, sChildPartNo, 1, lLineNumber, sCheckinComments, "wt.part.WTPartUsageLink", "ea", sAttributeNames, sAttributeValues, sAttributeTypes, Convert.ToInt16(sWebAppId));
            }
        }

        public string SetMaintenanceTemplates(string sSessionId, string sUserId, string sWONo, string sWOName, string sTemplateIndex, string sWebAppId)
        {
            object nullobject = Type.Missing;
            string sReturn = "Success";

            try
            {
                if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
                {
                    return "User " + sUserId + " is not logged in";
                }
                else
                {
                    string sBaseFolder = GetConstantValue("TemplateFolder", 2);
                    string sOutFolder = GetConstantValue("GeneratedDocsFolder", 2);
                    string sTemplateName = sBaseFolder;
                    string sWONameFile = sWOName;
                    sWONameFile = RemoveInvalidCharacters(sWONameFile);
                    string sFileOutName = sOutFolder + @"\" + sWONo + " " + sWONameFile + ".docm";
                    string sFileOutNamePdf = sOutFolder + @"\" + sWONo + " " + sWONameFile + ".pdf";
                    int iTemplateIndex;

                    iTemplateIndex = Convert.ToInt32(sTemplateIndex);

                    sTemplateName = sTemplateName + @"\" + GetTemplateName(iTemplateIndex, 2);
                    word.Application ap = new word.Application();
                    word.Document doc = ap.Documents.Open(sTemplateName);

                    word.Cell cell = doc.Tables[1].Cell(1, 2);

                    cell.Range.Text = sWONo;

                    ap.Run("btnProcess_Click");

                    if (FileExists(sFileOutName))
                        File.Delete(sFileOutName);

                    doc.SaveAs2(sFileOutName);

                    doc.Shapes[1].Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
                    doc.Shapes[2].Visible = Microsoft.Office.Core.MsoTriState.msoFalse;

                    if (FileExists(sFileOutNamePdf))
                        File.Delete(sFileOutNamePdf);

                    doc.ExportAsFixedFormat(sFileOutNamePdf, word.WdExportFormat.wdExportFormatPDF, false);
                    doc.Shapes[1].Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
                    doc.Shapes[2].Visible = Microsoft.Office.Core.MsoTriState.msoTrue;

                    ((Microsoft.Office.Interop.Word._Document)doc).Close(ref nullobject, ref nullobject, ref nullobject);
                    ((Microsoft.Office.Interop.Word._Application)ap).Quit(ref nullobject, ref nullobject, ref nullobject);

                    return sReturn;
                }
            }
            catch (Exception ex)
            {
                return "Failure^" + ex.Message + "^";
            }
        }
        public string SetMaintenanceTemplatesOpenXML(string sSessionId, string sUserId, string sWONo, string sWOName)
        {
            object nullobject = Type.Missing;
            string sReturn = "Success";
            string sBaseFolder = @"C:\temp\";
            string sTemplateName = sBaseFolder + "WorkOrderSummaryTemplate_v6.docm";
            string sFileOutName = sBaseFolder + sWONo + ".docm";

            word.Application ap = new word.Application();
            word.Document doc = ap.Documents.Open(sTemplateName);

            WordprocessingDocument doc2 = WordprocessingDocument.Open(sTemplateName, true);

            word.Cell cell = doc.Tables[1].Cell(1, 2);

            cell.Range.Text = sWONo;

            if (FileExists(sFileOutName))
                File.Delete(sFileOutName);

            doc.SaveAs2(sFileOutName);
            //doc.Close();
            //ap.Quit();

            ((Microsoft.Office.Interop.Word._Document)doc).Close(ref nullobject, ref nullobject, ref nullobject);
            ((Microsoft.Office.Interop.Word._Application)ap).Quit(ref nullobject, ref nullobject, ref nullobject);

            return sReturn;
        }

        public string ProcessIOSpreadsheet(string sSessionId, string sUserId, string sFile, string sWebAppId)
        {

            Excel.Application xlApp = null;
            Excel.Workbooks xlWbks = null;
            try
            {
                int iWebAppId = Convert.ToInt32(sWebAppId);

                if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
                {
                    return "User " + sUserId + " is not logged in";
                }
                else
                {
                    Update_User_Time(sUserId, sSessionId);
                    ArrayList arrUser = GetUserDetails(sUserId);
                    string sFullName = arrUser[2].ToString();
                    string sRecipeints = arrUser[3].ToString();
                    string sCheckinComments = "";

                    xlApp = new Excel.Application();
                    xlWbks = xlApp.Workbooks;

                    Excel.Workbook xlWorkbook = xlWbks.Open(@"C:\Webroot\Regain\Uploads\" + sFile);
                    Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                    Excel.Range xlRange = xlWorksheet.UsedRange;

                    int rowCount = xlRange.Rows.Count;
                    int colCount = xlRange.Columns.Count;
                    int i = 0;
                    string sBody = "";

                    for (i = 7; i <= rowCount; i++)
                    {
                        string sParentPartNo = "";
                        if (xlRange.Cells[i, 1].Value2 != null)
                            sParentPartNo = xlRange.Cells[i, 1].Value2.ToString();

                        string sExisitngChildPartNo = "";
                        if (xlRange.Cells[i, 9].Value2 != null)
                            sExisitngChildPartNo = xlRange.Cells[i, 9].Value2.ToString();

                        string sChildPartNoNew = "";
                        if (xlRange.Cells[i, 10].Value2 != null)
                            sChildPartNoNew = xlRange.Cells[i, 10].Value2.ToString();

                        string sIOType = "";
                        if (xlRange.Cells[i, 7].Value2 != null)
                            sIOType = xlRange.Cells[i, 7].Value2.ToString();

                        string sIOTag = "";
                        if (xlRange.Cells[i, 8].Value2 != null)
                            sIOTag = xlRange.Cells[i, 8].Value2.ToString();

                        string sAction = "";
                        if (xlRange.Cells[i, 11].Value2 != null)
                            sAction = xlRange.Cells[i, 11].Value2.ToString();

                        sAction = sAction.ToUpper();

                        if (!sAction.Equals("") && !sParentPartNo.Equals(""))
                        {
                            string sRowMsg = "";
                            bool bParentExists = PartExists(sParentPartNo, iWebAppId);
                            bool bChildExists = PartExists(sExisitngChildPartNo, iWebAppId);
                            bool bChildExistsNew = PartExists(sChildPartNoNew, iWebAppId);
                            rtnInt clsLinkExists = PartIOLinkExists(sParentPartNo, sExisitngChildPartNo, sIOType, sIOTag, iWebAppId);
                            rtnInt clsLinkExistsNew = PartIOLinkExists(sParentPartNo, sChildPartNoNew, sIOType, sIOTag, iWebAppId);
                            switch (sAction)
                            {
                                case "ADD":
                                    if(clsLinkExistsNew.bReturnValue)
                                    {
                                        sRowMsg = "The I/O link for row " + i + " of file " + sFile + " already exists. You cannot add 2 links of the same I/O type and tag.\r\n";
                                    }
                                    else
                                    {
                                        if (!bParentExists)
                                        {
                                            if(!bChildExistsNew)
                                                sRowMsg = "Plant equipment item " + sParentPartNo + " and PLC Address " + sChildPartNoNew +  " on row " + i + " of file " + sFile + " do not exist. Plant equipment and PLC I/O Addresses must exist before you can add a PLC I/O to them.\r\n";
                                            else
                                                sRowMsg = "Plant equipment item " + sParentPartNo + " on row " + i + " of file " + sFile + " does not exist. Plant equipment must exist before you can add a PLC I/O to it.\r\n";
                                        }
                                        else
                                        {

                                            if (!bChildExistsNew)
                                                sRowMsg = "PLC Address " + sChildPartNoNew + " on row " + i + " of file " + sFile + " does not exist. PLC I/O Addresses must exist before you can add a PLC I/O to it.\r\n";
                                            else
                                            {
                                                int iNewLineNumber = GetNewLineNumber(sParentPartNo, iWebAppId);
                                                sCheckinComments = "Adding I/O link between " + sParentPartNo + " and " + sChildPartNoNew + " of type " + sIOType + " with tag " + sIOTag + " via spreadsheet processing";
                                                string sRtn = SetMaterialIOPartToPartLink(sSessionId, sUserId, sFullName, sParentPartNo, sChildPartNoNew,
                                                                                          iNewLineNumber.ToString(), sIOType, sIOTag, sCheckinComments, sWebAppId);
                                                if(!sRtn.Equals("Success"))
                                                {
                                                    sBody += sRtn;
                                                }
                                            }
                                        }

                                    }
                                    break;
                                case "UPDATE":
                                    if (!clsLinkExists.bReturnValue)
                                    {
                                        sRowMsg = "The I/O link for row " + i + " of file " + sFile + " must exist for you to update it. Plant equipment " + sParentPartNo + " and PLC Address " + sExisitngChildPartNo + " with I/O type " + sIOType + " and tag " + sIOTag + " must exist for this update to take place.\r\n";
                                    }
                                    else
                                    {
                                        if (!bParentExists)
                                        {
                                            if (!bChildExistsNew)
                                                sRowMsg = "Plant equipment item " + sParentPartNo + " and PLC Address " + sChildPartNoNew + " on row " + i + " of file " + sFile + " do not exist. Plant equipment and PLC I/O Addresses must exist before you can update a PLC I/O between them.\r\n";
                                            else
                                                sRowMsg = "Plant equipment item " + sParentPartNo + " on row " + i + " of file " + sFile + " does not exist. Plant equipment must exist before you can update a PLC I/O from it.\r\n";
                                        }
                                        else
                                        {

                                            if (!bChildExistsNew)
                                                sRowMsg = "PLC Address " + sChildPartNoNew + " on row " + i + " of file " + sFile + " does not exist. PLC I/O Addresses must exist before you can update a PLC I/O from it.\r\n";
                                            else
                                            {
                                                //THis means we are changing the child part of the link
                                                //Firstly delete the existing link and then create a new one
                                                int iLineNumber = clsLinkExists.iReturnValue;
                                                sCheckinComments = "Deleting I/O link between " + sParentPartNo + " and " + sExisitngChildPartNo + " of type " + sIOType + " with tag " + sIOTag + " via spreadsheet processing";
                                                string sRtnDelete = DeletePartToPartLinkByLineNumber(sSessionId, sUserId, sFullName, iLineNumber.ToString(), sParentPartNo, sExisitngChildPartNo, sCheckinComments, sWebAppId);
                                                if (!sRtnDelete.Equals("Success"))
                                                {
                                                    sBody += sRtnDelete;
                                                }
                                                else
                                                {
                                                    int iNewLineNumber = GetNewLineNumber(sParentPartNo, iWebAppId);
                                                    sCheckinComments = "Creating I/O link between " + sParentPartNo + " and " + sChildPartNoNew + " of type " + sIOType + " with tag " + sIOTag + " via spreadsheet processing";
                                                    string sRtn = SetMaterialIOPartToPartLink(sSessionId, sUserId, sFullName, sParentPartNo, sChildPartNoNew,
                                                                                              iNewLineNumber.ToString(), sIOType, sIOTag, sCheckinComments, sWebAppId);
                                                    if (!sRtn.Equals("Success"))
                                                    {
                                                    }
                                                }
                                            }
                                        }

                                    }
                                    break;
                                case "DELETE":
                                    if (!clsLinkExists.bReturnValue)
                                    {
                                        sRowMsg = "The I/O link for row " + i + " of file " + sFile + " must exist for you to delete it. Plant equipment " + sParentPartNo + " and PLC Address " + sExisitngChildPartNo + " with I/O type " + sIOType + " and tag " + sIOTag + " must exist for this deletion to take place.\r\n";
                                    }
                                    else
                                    {
                                        int iLineNumber = clsLinkExists.iReturnValue;
                                        sCheckinComments = "Deleting I/O link between " + sParentPartNo + " and " + sExisitngChildPartNo + " of type " + sIOType + " with tag " + sIOTag + " via spreadsheet processing";
                                        string sRtnDelete = DeletePartToPartLinkByLineNumber(sSessionId, sUserId, sFullName, iLineNumber.ToString(), sParentPartNo, sExisitngChildPartNo, sCheckinComments, sWebAppId);
                                        if (!sRtnDelete.Equals("Success"))
                                        {
                                            sBody += sRtnDelete;
                                        }

                                    }
                                    break;
                                default:
                                    sRowMsg = "The action must be one of ADD, UPDATE OR DELETE. Row " + i + " of file " + sFile + " cannot be processed.\r\n";
                                    break;
                            }

                            sBody += sRowMsg;

                        }
                    }

                    xlWorkbook.Close(true);
                    xlWbks.Close();
                    xlApp.Quit();

                    while (System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp) != 0) ;
                    while (System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbks) != 0) ;
                    while (System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkbook) != 0) ;
                    while (System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorksheet) != 0) ;
                    while (System.Runtime.InteropServices.Marshal.ReleaseComObject(xlRange) != 0) ;
                    xlApp = null;
                    xlWbks = null;
                    xlWorkbook = null;
                    xlWorksheet = null;
                    xlRange = null;

                    //Now email the user
                    string sSubject = "Processing of File " + sFile;
                    if (sBody.Length == 0)
                        sBody = "No issues.";
                    sBody = "File " + sFile + " was processed with the following issues.\r\n" + sBody;
//                    emailmessage(sSessionId, sUserId, sSubject, sBody, " ", sRecipeints, "", "", sWebAppId);

                    return "Success^"+ sBody;
                }
            }
            catch(Exception ex)
            {
                return "Failure:" + ex.Message + "^";
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();

                System.Diagnostics.Process[] excelProcs = System.Diagnostics.Process.GetProcessesByName("EXCEL");
                foreach (System.Diagnostics.Process proc in System.Diagnostics.Process.GetProcessesByName("EXCEL"))
                {
                    proc.Kill();
                }
            }
        }

        public string ProcessIOPreallocatedSpreadsheet(string sSessionId, string sUserId, string sFile, string sWebAppId)
        {

            Excel.Application xlApp = null;
            Excel.Workbooks xlWbks = null;
            int iUpdateCount = 0;
            try
            {
                int iWebAppId = Convert.ToInt32(sWebAppId);

                if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
                {
                    return "User " + sUserId + " is not logged in";
                }
                else
                {
                    Update_User_Time(sUserId, sSessionId);
                    ArrayList arrUser = GetUserDetails(sUserId);
                    string sFullName = arrUser[2].ToString();
                    string sRecipeints = arrUser[3].ToString();
                    string sCheckinComments = "";

                    xlApp = new Excel.Application();
                    xlWbks = xlApp.Workbooks;

                    Excel.Workbook xlWorkbook = xlWbks.Open(@"C:\Webroot\Regain\Uploads\" + sFile);
                    Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                    Excel.Range xlRange = xlWorksheet.UsedRange;

                    int rowCount = xlRange.Rows.Count;
                    int colCount = xlRange.Columns.Count;
                    int i = 0;
                    string sBody = "";
                    bool bRtn;

                    for (i = 7; i <= rowCount; i++)
                    {
                        string sPartNo = "";
                        if (xlRange.Cells[i, 1].Value2 != null)
                            sPartNo = xlRange.Cells[i, 1].Value2.ToString();

                        string sIOType = "";
                        if (xlRange.Cells[i, 2].Value2 != null)
                            sIOType = xlRange.Cells[i, 2].Value2.ToString();

                        string sIOTag = "";
                        if (xlRange.Cells[i, 3].Value2 != null)
                            sIOTag = xlRange.Cells[i, 3].Value2.ToString();

                        string sChassis = "";
                        if (xlRange.Cells[i, 10].Value2 != null)
                            sChassis = xlRange.Cells[i, 10].Value2.ToString();

                        string sSlot = "";
                        if (xlRange.Cells[i, 11].Value2 != null)
                            sSlot = xlRange.Cells[i, 11].Value2.ToString();

                        string sChannel = "";
                        if (xlRange.Cells[i, 12].Value2 != null)
                            sChannel = xlRange.Cells[i, 12].Value2.ToString();

                        string sLockStatus = "";
                        if (xlRange.Cells[i, 19].Value2 != null)
                            sLockStatus = xlRange.Cells[i, 19].Value2.ToString();

                        string sAction = "";
                        if (xlRange.Cells[i, 20].Value2 != null)
                            sAction = xlRange.Cells[i, 20].Value2.ToString();

                        sAction = sAction.ToUpper();

                        if (!sAction.Equals(""))
                        {
                            string sRowMsg = "";
                            bool bPartExists = PartExists(sPartNo, iWebAppId);
                            rtnInt clsLinkExists = PartIOLinkExistsNoChildRequired(sPartNo, sIOType, sIOTag, iWebAppId);
                            switch (sAction)
                            {
                                case "UPDATE":
                                    if (!clsLinkExists.bReturnValue)
                                    {
                                        sRowMsg = "The I/O tag for row " + i + " of file " + sFile + " must exist for you to update it. Plant equipment " + sPartNo + " with I/O type " + sIOType + " and tag " + sIOTag + " must exist for this update to take place.\r\n";
                                    }
                                    else
                                    {
                                        if (!bPartExists)
                                        {
                                                sRowMsg = "Plant equipment item " + sPartNo + " on row " + i + " of file " + sFile + " does not exist. Plant equipment must exist before you can lock in a PLC chassis/slot/channel.\r\n";
                                        }
                                        else
                                        {
                                                //This means we are removing a hard lock (going from existing hard to soft)
                                                if(sLockStatus.ToUpper().Equals("SOFT"))
                                                {
                                                    bRtn = SetAlbaPLC_LockedInfo(sChassis, Convert.ToInt16(sSlot), Convert.ToInt16(sChannel), "", "", "", iWebAppId);
                                                    if (!bRtn)
                                                        sRowMsg = "Could not update equipment " + sPartNo + " with I/O type " + sIOType + " and I/O Tag " + sIOTag + " to soft. Attempt to remove the hard lock was not successful.";

                                                }
                                                else
                                                {
                                                    bRtn = SetAlbaPLC_LockedInfo(sChassis, Convert.ToInt16(sSlot), Convert.ToInt16(sChannel), sPartNo, sIOType, sIOTag, iWebAppId);
                                                    if (!bRtn)
                                                        sRowMsg = "Could not update equipment " + sPartNo + " with I/O type " + sIOType + " and I/O Tag " + sIOTag + " to soft. Attempt to remove the hard lock was not successful.";

                                                }

                                        }

                                    }
                                    iUpdateCount++;
                                    break;
                                case "DELETE":
                                    bRtn = SetAlbaPLC_LockedInfo(sChassis, Convert.ToInt16(sSlot), Convert.ToInt16(sChannel), "", "", "", iWebAppId);
                                    if (!bRtn)
                                        sRowMsg = "Could not update equipment " + sPartNo + " with I/O type " + sIOType + " and I/O Tag " + sIOTag + " to soft. Attempt to remove the hard lock was not successful.";
                                    iUpdateCount++;
                                    break;
                                default:
                                    sRowMsg = "The action must be one of UPDATE OR DELETE. Row " + i + " of file " + sFile + " cannot be processed.\r\n";
                                    iUpdateCount++;
                                    break;
                            }

                            sBody += sRowMsg;

                        }
                    }

                    xlWorkbook.Close(true);
                    xlWbks.Close();
                    xlApp.Quit();

                    while (System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp) != 0) ;
                    while (System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbks) != 0) ;
                    while (System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkbook) != 0) ;
                    while (System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorksheet) != 0) ;
                    while (System.Runtime.InteropServices.Marshal.ReleaseComObject(xlRange) != 0) ;
                    xlApp = null;
                    xlWbks = null;
                    xlWorkbook = null;
                    xlWorksheet = null;
                    xlRange = null;

                    //Now email the user
                    string sSubject = "PLC Preallocation Processing of File " + sFile;
                    if (sBody.Length == 0)
                        sBody = "No issues.";

                    if (iUpdateCount == 0)
                        sBody += "\r\nNo items were marked to be modified.";

                    sBody = "File " + sFile + " was processed with the following issues.\r\n" + sBody;
                    //                    emailmessage(sSessionId, sUserId, sSubject, sBody, " ", sRecipeints, "", "", sWebAppId);

                    return "Success^" + sBody;
                }
            }
            catch (Exception ex)
            {
                return "Failure:" + ex.Message + "^";
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();

                System.Diagnostics.Process[] excelProcs = System.Diagnostics.Process.GetProcessesByName("EXCEL");
                foreach (System.Diagnostics.Process proc in System.Diagnostics.Process.GetProcessesByName("EXCEL"))
                {
                    proc.Kill();
                }
            }
        }

        public string ProcessCableSpreadsheet(string sSessionId, string sUserId, string sFile, string sWebAppId)
        {

            Excel.Application xlApp = null;
            Excel.Workbooks xlWbks = null;
            try
            {
                int iWebAppId = Convert.ToInt32(sWebAppId);

                if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
                {
                    return "User " + sUserId + " is not logged in";
                }
                else
                {
                    Update_User_Time(sUserId, sSessionId);
                    ArrayList arrUser = GetUserDetails(sUserId);
                    string sFullName = arrUser[2].ToString();
                    string sRecipeints = arrUser[3].ToString();
                    string sCheckinComments = "";
                    string sExistingComments = "";

                    xlApp = new Excel.Application();
                    xlWbks = xlApp.Workbooks;

                    Excel.Workbook xlWorkbook = xlWbks.Open(@"C:\Webroot\Regain\Uploads\" + sFile);
                    Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                    Excel.Range xlRange = xlWorksheet.UsedRange;

                    int rowCount = xlRange.Rows.Count;
                    int colCount = xlRange.Columns.Count;
                    int i = 0;
                    string sBody = "";

                    for (i = 7; i <= rowCount; i++)
                    {
                        string sCableNo = "";
                        if (xlRange.Cells[i, 1].Value2 != null)
                            sCableNo = xlRange.Cells[i, 1].Value2.ToString();

                        string sFromEquipNo = "";
                        if (xlRange.Cells[i, 2].Value2 != null)
                            sFromEquipNo = xlRange.Cells[i, 2].Value2.ToString();

                        string sToEquipNo = "";
                        if (xlRange.Cells[i, 4].Value2 != null)
                            sToEquipNo = xlRange.Cells[i, 4].Value2.ToString();

                        string sMaterialCode = "";
                        if (xlRange.Cells[i, 6].Value2 != null)
                            sMaterialCode = xlRange.Cells[i, 6].Value2.ToString();

                        double dCableLength = 0.0;
                        if (xlRange.Cells[i, 8].Value2 != null)
                            dCableLength = Convert.ToDouble(xlRange.Cells[i, 8].Value2.ToString());

                        string sComments = "";
                        if (xlRange.Cells[i, 9].Value2 != null)
                            sComments = xlRange.Cells[i, 9].Value2.ToString();

                        string sAction = "";
                        if (xlRange.Cells[i, 10].Value2 != null)
                            sAction = xlRange.Cells[i, 10].Value2.ToString();

                        sAction = sAction.ToUpper();

                        if (!sAction.Equals(""))
                        {
                            string sRowMsg = "";
                            int iExistingFromLineNumber = -1;
                            int iExistingToLineNumber = -1;
                            bool bFromPartExists = PartExists(sFromEquipNo, iWebAppId);
                            bool bToPartExists = PartExists(sToEquipNo, iWebAppId);
                            bool bCableExists = PartExists(sCableNo, iWebAppId);
                            bool bCableMaterialExists = PartExists(sMaterialCode, iWebAppId);
                            rtnInt clsFromLinkExists = CablePartLinkExists(sFromEquipNo, sCableNo, 0,  iWebAppId);
                            rtnInt clsToLinkExists = CablePartLinkExists(sToEquipNo, sCableNo, 1, iWebAppId);
                            rtnInt clsCableMaterialExists = CableMaterialLinkExists(sCableNo, sMaterialCode, 2);
                            rtnString clsCableComments = GetPartStringAttribute(sCableNo, "Comments", iWebAppId);

                            if(clsCableComments.bReturnValue)
                            {
                                sExistingComments = clsCableComments.sReturnValue;
                            }
                            //Get any existing From equipment should it exist
                            string sExistingFromEquip = "";
                            if (!clsFromLinkExists.bReturnValue)
                            {
                                rtnString clsFromLinkParent = CablePartLinkParent(sCableNo, 0, iWebAppId);
                                sExistingFromEquip = clsFromLinkParent.sReturnValue;
                                iExistingFromLineNumber = clsFromLinkParent.iLineNumber;
                            }
                            else
                            {
                                sExistingFromEquip = sFromEquipNo;
                                iExistingFromLineNumber = clsFromLinkExists.iReturnValue;
                            }

                            //Get any existing To equipment should it exist
                            string sExistingToEquip = "";
                            if (!clsToLinkExists.bReturnValue)
                            {
                                rtnString clsToLinkParent = CablePartLinkParent(sCableNo, 1, iWebAppId);
                                sExistingToEquip = clsToLinkParent.sReturnValue;
                                iExistingToLineNumber = clsToLinkParent.iLineNumber;
                            }
                            else
                            {
                                sExistingToEquip = sToEquipNo;
                                iExistingToLineNumber = clsToLinkExists.iReturnValue;
                            }

                            string sExistingMaterialCode = "";
                            long lExistingMaterialLineNumber = 0;
                            if (!clsCableMaterialExists.bReturnValue)  //This means the material to be set does not already exist linked to this cable. So find the one that does in case we need to delete it.
                            {
                                rtnString clsCableMaterial = CableMaterialChild(sCableNo, iWebAppId);
                                sExistingMaterialCode = clsCableMaterial.sReturnValue;
                                lExistingMaterialLineNumber = (long)clsCableMaterial.iLineNumber;
                            }
                            else
                            {
                                sExistingMaterialCode = sMaterialCode;
                                rtnString clsCableMaterial = CableMaterialChild(sCableNo, iWebAppId);
                                lExistingMaterialLineNumber = (long)clsCableMaterial.iLineNumber;
                            }

                            switch (sAction)
                            {
                                case "ADD":
                                    if (clsToLinkExists.bReturnValue)
                                    {
                                        sRowMsg = "The cable for row " + i + " of file " + sFile + " already exists. You cannot add 2 cables of the same type to the same equipment.\r\n";
                                    }
                                    else
                                    {
                                        if (!bToPartExists)
                                        {
                                            sRowMsg = "The 'To' Equipment " + sToEquipNo + " on row " + i + " of file " + sFile + " does not exist. Plant equipment must exist before you can connect a cable.\r\n";
                                        }
                                        else
                                        {

                                            int iIsNumber = 0;
                                            bool bCableCounter = int.TryParse(sCableNo.Substring(sCableNo.Length - 1, 1), out iIsNumber);
                                            if (!sCableNo.Substring(0, sCableNo.IndexOf("-")).Equals(sToEquipNo) || !bCableCounter)
                                            {
                                                sRowMsg = "The cable no " + sCableNo + " on row " + i + " of file " + sFile + " does not match the 'To' Equipment " + sToEquipNo +". The cable no has the same prefix as the 'To' Equipment followed by a '-' and then P,C,I,D or E and a number.\r\n";
                                            }
                                            else
                                            {
                                                if (!bFromPartExists && !sFromEquipNo.Equals(""))
                                                    sRowMsg = "The 'From' Equipment " + sFromEquipNo + " on row " + i + " of file " + sFile + " does not exist. Plant equipment must exist before you can connect a cable.\r\n";
                                                else
                                                {
                                                    if (bFromPartExists && !sExistingFromEquip.Equals(""))
                                                        sRowMsg = "The 'From' side is already connected to " + sExistingFromEquip + " on row " + i + " of file " + sFile + ". You need to modify this row, not add.\r\n";
                                                    else
                                                    {
                                                        if (bToPartExists && !sExistingToEquip.Equals(""))
                                                            sRowMsg = "The 'To' side is already connected to " + sExistingToEquip + " on row " + i + " of file " + sFile + ". You cannot add or modify this row, because the cable number is related to the 'To' equipment. Please delete this cable if you wish to change the 'To' end.\r\n";
                                                        else
                                                        {
                                                            if (!bCableMaterialExists)
                                                                sRowMsg = "The cable material code " + sMaterialCode + " on row " + i + " of file " + sFile + " does not exist. You cannot add cable mateial unless it exists.\r\n";
                                                            else
                                                            {
                                                                if (bCableMaterialExists && !sExistingMaterialCode.Equals(""))
                                                                    sRowMsg = "There is already cable material of " + sExistingMaterialCode + " connected for cable " + sCableNo + " on row " + i + " of file " + sFile + ". You cannot add a cable when it already exists. Please use a modify action if you wish to change the material of this cable.\r\n";
                                                                else
                                                                {
                                                                    string sJob = "";
                                                                    int iProdOrLib = 0;
                                                                    string sFolder = "";

                                                                    if (sToEquipNo.StartsWith("M"))
                                                                    {
                                                                        iProdOrLib = 1;
                                                                        sJob = "M";
                                                                        sFolder = "Material Catalogue/Cable Schedule";
                                                                    }
                                                                    else
                                                                    {
                                                                        int iJob = Convert.ToInt32(sCableNo.Substring(1, 3));
                                                                        sJob = iJob.ToString();
                                                                        rtnString rtnStr2 = GetPlantJobFolder(iJob, iWebAppId);
                                                                        if (rtnStr2.bReturnValue)
                                                                            sFolder = rtnStr2.sReturnValue + "/Cable Schedule";
                                                                    }

                                                                    string sProductName = "";
                                                                    rtnString rtnStr1 = GetProductFromJob(sJob, iProdOrLib,iWebAppId);
                                                                    if (rtnStr1.bReturnValue)
                                                                        sProductName = rtnStr1.sReturnValue;


                                                                    int iNewFromLineNumber = GetNewLineNumber(sFromEquipNo, iWebAppId);
                                                                    int iNewToLineNumber = GetNewLineNumber(sToEquipNo, iWebAppId);
                                                                    sCheckinComments = "Adding cable " + sCableNo + " via spreadsheet processing";
                                                                    string sRtn = CreateCableItem(sSessionId, sUserId, sProductName, sFolder,
                                                                                                  sCableNo, "Cable " + sCableNo, sFromEquipNo, sToEquipNo, dCableLength.ToString(),
                                                                                                  iNewFromLineNumber.ToString(), iNewToLineNumber.ToString(), sMaterialCode,
                                                                                                  sFullName, sComments, sCheckinComments, iProdOrLib.ToString(), sWebAppId);
                                                                    if (!sRtn.Equals("Success"))
                                                                    {
                                                                        sBody += sRtn;
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }

                                    }
                                    break;
                                case "UPDATE":
                                    if (!bCableExists)
                                    {
                                        sRowMsg = "The cable " + sCableNo + " on row " + i + " of file " + sFile + " does not exist. Cables must exist before you can perform an update. Maybe you wanted to have the action as an 'Add' instead.\r\n";

                                    }
                                    else
                                    {
                                        if (!bToPartExists)
                                        {
                                            sRowMsg = "The 'To' Equipment " + sToEquipNo + " on row " + i + " of file " + sFile + " does not exist. Plant equipment must exist before you can connect a cable.\r\n";
                                        }
                                        else
                                        {

                                            if (!bFromPartExists && !sFromEquipNo.Equals(""))
                                                sRowMsg = "The 'From' Equipment " + sFromEquipNo + " on row " + i + " of file " + sFile + " does not exist. Plant equipment must exist before you can connect a cable.\r\n";
                                            else
                                            {
                                                if (bFromPartExists && !sFromEquipNo.Equals(sExistingFromEquip))
                                                {
                                                    string sRtn = "";
                                                    string sCheckInComments = "";
                                                    if (!sExistingFromEquip.Equals(""))
                                                    {
                                                        sRtn = UpdateCableFromDetails(sSessionId, sUserId, sCableNo, sFromEquipNo, sWebAppId);
                                                        if (!sRtn.StartsWith("Success"))
                                                        {
                                                            sRowMsg = "Could not process row " + i + " of file " + sFile + ". Could not swap terminations from functional location " + sExistingFromEquip +
                                                                      " to functional location " + sFromEquipNo + ".\r\n";
                                                        }

                                                        sCheckInComments = "Deleting link to cable " + sCableNo;
                                                        sRtn = DeletePartToPartLinkByLineNumber(sSessionId, sUserId, sFullName, iExistingFromLineNumber.ToString(),
                                                                                                      sExistingFromEquip, sCableNo, sCheckInComments, "2");
                                                        if (!sRtn.StartsWith("Success"))
                                                        {
                                                            sRowMsg = "Could not process row " + i + " of file " + sFile + ". Could not remove exisitng link from functional location " + sExistingFromEquip + ".\r\n";
                                                        }
                                                    }
                                                    else
                                                        sRtn = "Success";

                                                    if (sRtn.StartsWith("Success"))
                                                    {
                                                        sCheckInComments = "Adding from functional location " + sFromEquipNo + " link to cable " + sCableNo;
                                                        int iNewLineNumber = GetNewLineNumber(sFromEquipNo, iWebAppId);
                                                        sRtn = CreateCablePartLink(sSessionId, sUserId, sCableNo, sFromEquipNo, iNewLineNumber.ToString(), "0", sCheckinComments, sWebAppId);
                                                        if (!sRtn.StartsWith("Success"))
                                                        {
                                                            sRowMsg = "Could not process row " + i + " of file " + sFile + ". Could not add link to funcitonal location " + sFromEquipNo + ".\r\n";
                                                        }
                                                    }
                                                }

                                                if (!bCableMaterialExists)
                                                    sRowMsg = "The cable material code " + sMaterialCode + " on row " + i + " of file " + sFile + " does not exist. You cannot add cable material unless it exists.\r\n";
                                                else
                                                {
                                                    string sRtn = "";
                                                    if (!sExistingComments.Equals(sComments))
                                                    {
                                                        sCheckinComments = "Updating cable " + sCableNo + " via spreadsheet processing";
                                                        sRtn = UpdateCableItem(sSessionId, sUserId, sCableNo, "Cable " + sCableNo, sFullName, sComments, sCheckinComments, sWebAppId);
                                                    }
                                                    else
                                                        sRtn = "Success";

                                                    if (!sRtn.Equals("Success"))
                                                    {
                                                        sRowMsg += "The cable material code " + sMaterialCode + " on row " + i + " of file " + sFile + " could not be updated.\r\n";
                                                    }
                                                    else
                                                    {
                                                        if (sExistingMaterialCode.Equals(sMaterialCode))
                                                        {
                                                            double dExistingQty = GetUsageLinkExistingQty(sCableNo, sExistingMaterialCode, lExistingMaterialLineNumber, iWebAppId);
                                                            if (dCableLength != dExistingQty)
                                                                sRtn = SetPartUsageLinkQty(sSessionId, sUserId, sCableNo, sExistingMaterialCode, dCableLength.ToString(), sWebAppId);
                                                            else
                                                                sRtn = "Success";
                                                        }
                                                        else
                                                        {
                                                            sCheckinComments = "Creating link betwwwn cable no " + sCableNo + " and material with code " + sMaterialCode;
                                                            sRtn = UpdateCableMaterial(sSessionId, sUserId, sFullName, sCableNo, dCableLength.ToString(), sMaterialCode, sExistingMaterialCode, sCheckinComments, sWebAppId);
                                                            if (!sRtn.Equals("Success"))
                                                            {
                                                                sRowMsg += "The link between cable material code " + sMaterialCode + " and cable no " + sCableNo + " on row " + i + " of file " + sFile + " could not be updated.\r\n";
                                                            }

                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }

                                    break;
                                case "DELETE":
                                    if (!bCableExists)
                                    {
                                        sRowMsg = "The cable " + sCableNo + " on row " + i + " of file " + sFile + " does not exist. Cables must exist before you can perform a delete. Maybe you wanted to have the action as an 'Add' instead.\r\n";

                                    }
                                    else
                                    {
                                        if (bFromPartExists && clsFromLinkExists.bReturnValue)
                                        {
                                            string sCheckInComments = "Deleting link to cable " + sCableNo;
                                            string sRtn2 = DeletePartToPartLinkByLineNumber(sSessionId, sUserId, sFullName, iExistingFromLineNumber.ToString(),
                                                                                            sExistingFromEquip, sCableNo, sCheckInComments, "2");

                                            if (!sRtn2.StartsWith("Success"))
                                            {
                                                sRowMsg = "Could not process row " + i + " of file " + sFile + ". Could not remove exisitng link from functional location " + sExistingFromEquip + ".\r\n";
                                            }
                                        }

                                        if (bToPartExists && clsToLinkExists.bReturnValue)
                                        {
                                            string sCheckInComments = "Deleting link to cable " + sCableNo;
                                            string sRtn3 = DeletePartToPartLinkByLineNumber(sSessionId, sUserId, sFullName, iExistingToLineNumber.ToString(),
                                                                                          sExistingToEquip, sCableNo, sCheckInComments, "2");

                                            if (!sRtn3.StartsWith("Success"))
                                            {
                                                sRowMsg = "Could not process row " + i + " of file " + sFile + ". Could not remove exisitng link to functional location " + sExistingToEquip + ".\r\n";
                                            }
                                        }

                                        if (bCableMaterialExists && clsCableMaterialExists.bReturnValue)
                                        {
                                            string sCheckInComments = "Deleting link from cable " + sCableNo + " to cable material " + sExistingMaterialCode;
                                            string sRtn4 = DeletePartToPartLinkByLineNumber(sSessionId, sUserId, sFullName, iExistingToLineNumber.ToString(),
                                                                                          sCableNo, sExistingMaterialCode, sCheckInComments, "2");

                                            if (!sRtn4.StartsWith("Success"))
                                            {
                                                sRowMsg = "Could not process row " + i + " of file " + sFile + ". Could not remove exisitng link to functional location " + sExistingToEquip + ".\r\n";
                                            }
                                        }

                                        string sRtn5 = SetPartState(sSessionId, sUserId, sCableNo, "Obsolete", sWebAppId);
                                        if (!sRtn5.StartsWith("Success"))
                                        {
                                            sRowMsg = "Could not process row " + i + " of file " + sFile + ". Could not set the state of cable no " + sCableNo + " to obsolete." + ".\r\n";
                                        }
                                    }

                                    break;
                                default:
                                    sRowMsg = "The action must be one of ADD, UPDATE OR DELETE. Row " + i + " of file " + sFile + " cannot be processed.\r\n";
                                    break;
                            }

                            sBody += sRowMsg;

                        }
                    }

                    xlWorkbook.Close(true);
                    xlWbks.Close();
                    xlApp.Quit();

                    while (System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp) != 0) ;
                    while (System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbks) != 0) ;
                    while (System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkbook) != 0) ;
                    while (System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorksheet) != 0) ;
                    while (System.Runtime.InteropServices.Marshal.ReleaseComObject(xlRange) != 0) ;
                    xlApp = null;
                    xlWbks = null;
                    xlWorkbook = null;
                    xlWorksheet = null;
                    xlRange = null;

                    //Now email the user
                    string sSubject = "Processing of File " + sFile;
                    if (sBody.Length == 0)
                        sBody = "No issues.";
                    sBody = "File " + sFile + " was processed with the following issues.\r\n" + sBody;
                    //                    emailmessage(sSessionId, sUserId, sSubject, sBody, " ", sRecipeints, "", "", sWebAppId);

                    return "Success^" + sBody;
                }
            }
            catch (Exception ex)
            {
                return "Failure:" + ex.Message + "^";
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();

                System.Diagnostics.Process[] excelProcs = System.Diagnostics.Process.GetProcessesByName("EXCEL");
                foreach (System.Diagnostics.Process proc in System.Diagnostics.Process.GetProcessesByName("EXCEL"))
                {
                    proc.Kill();
                }
            }
        }


        public string ProcessTerminationSpreadsheet(string sSessionId, string sUserId, string sFile, string sWebAppId, string sFLOrMat)
        {

            Excel.Application xlApp = null;
            Excel.Workbooks xlWbks = null;
            try
            {
                int iWebAppId = Convert.ToInt32(sWebAppId);
                int iFLOrMat = Convert.ToInt32(sFLOrMat);

                if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
                {
                    return "User " + sUserId + " is not logged in";
                }
                else
                {
                    Update_User_Time(sUserId, sSessionId);
                    ArrayList arrUser = GetUserDetails(sUserId);
                    string sFullName = arrUser[2].ToString();
                    string sRecipeints = arrUser[3].ToString();

                    xlApp = new Excel.Application();
                    xlWbks = xlApp.Workbooks;

                    Excel.Workbook xlWorkbook = xlWbks.Open(@"C:\Webroot\Regain\Uploads\" + sFile);
                    Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                    Excel.Range xlRange = xlWorksheet.UsedRange;

                    int rowCount = xlRange.Rows.Count;
                    int colCount = xlRange.Columns.Count;
                    int i = 0;
                    string sBody = "";

                    for (i = 7; i <= rowCount; i++)
                    {
                        string sCableNo = "";
                        if (xlRange.Cells[i, 1].Value2 != null)
                            sCableNo = xlRange.Cells[i, 1].Value2.ToString();

                        string sFromEquipNo = "";
                        string sFromTermination = "";
                        string sAction = "";
                        int iIsNumber = -1;
                        int iCoreNo = -1;
                        bool bCoreIsNumeric;
                        string sCoreNo = "";
                        string sCoreLabel = "";
                        string sWireNo = "";
                        string sToTermination = "";
                        string sToEquipNo = "";

                        if (iFLOrMat == 0)
                        {
                            if (xlRange.Cells[i, 2].Value2 != null)
                                sFromEquipNo = xlRange.Cells[i, 2].Value2.ToString();

                            if (xlRange.Cells[i, 4].Value2 != null)
                                sFromTermination = xlRange.Cells[i, 4].Value2.ToString();

                            //if (!sFromTermination.StartsWith("X") && !sFromTermination.Contains("-"))
                            //    sFromTermination = "X1-" + sFromTermination;



                            if (xlRange.Cells[i, 5].Value2 != null)
                            {
                                string sCoreRaw = xlRange.Cells[i, 5].Value2.ToString();
                                iIsNumber = -1;
                                bCoreIsNumeric = int.TryParse(sCoreRaw, out iIsNumber);

                                if (bCoreIsNumeric)
                                    iCoreNo = Convert.ToInt32(xlRange.Cells[i, 5].Value2.ToString());
                            }

                            sCoreNo = iCoreNo.ToString();

                            if (xlRange.Cells[i, 6].Value2 != null)
                                sCoreLabel = xlRange.Cells[i, 6].Value2.ToString();

                            if (xlRange.Cells[i, 7].Value2 != null)
                                sWireNo = xlRange.Cells[i, 7].Value2.ToString();

                            if (xlRange.Cells[i, 8].Value2 != null)
                                sToTermination = xlRange.Cells[i, 8].Value2.ToString();

                            //if (!sToTermination.StartsWith("X") && !sToTermination.Contains("-"))
                            //    sToTermination = "X1-" + sToTermination;

                            if (xlRange.Cells[i, 9].Value2 != null)
                                sToEquipNo = xlRange.Cells[i, 9].Value2.ToString();

                            if (xlRange.Cells[i, 11].Value2 != null)
                                sAction = xlRange.Cells[i, 11].Value2.ToString();
                        }
                        else
                        {

                            if (xlRange.Cells[i, 2].Value2 != null)
                            {
                                string sCoreRaw = xlRange.Cells[i, 2].Value2.ToString();
                                iIsNumber = -1;
                                bCoreIsNumeric = int.TryParse(sCoreRaw, out iIsNumber);

                                if (bCoreIsNumeric)
                                    iCoreNo = Convert.ToInt32(xlRange.Cells[i, 2].Value2.ToString());
                            }

                            sCoreNo = iCoreNo.ToString();

                            if (xlRange.Cells[i, 3].Value2 != null)
                                sCoreLabel = xlRange.Cells[i, 3].Value2.ToString();

                            if (xlRange.Cells[i, 4].Value2 != null)
                                sWireNo = xlRange.Cells[i, 4].Value2.ToString();

                            if (xlRange.Cells[i, 5].Value2 != null)
                                sToTermination = xlRange.Cells[i, 5].Value2.ToString();

                            //if (!sToTermination.StartsWith("X") && !sToTermination.Contains("-"))
                            //    sToTermination = "X1-" + sToTermination;

                            if (xlRange.Cells[i, 6].Value2 != null)
                                sToEquipNo = xlRange.Cells[i, 6].Value2.ToString();

                            if (xlRange.Cells[i, 8].Value2 != null)
                                sAction = xlRange.Cells[i, 8].Value2.ToString();

                        }

                        sAction = sAction.ToUpper();

                        if (!sAction.Equals(""))
                        {
                            string sRowMsg = "";
                            bool bFromPartExists = PartExists(sFromEquipNo, iWebAppId);
                            bool bToPartExists = PartExists(sToEquipNo, iWebAppId);
                            bool bCableExists = PartExists(sCableNo, iWebAppId);

                            //To check the core we must prefix with the cable. But we do not send this
                            //to the create and update functins because it is done within them also.
                            //To those functions we just send the simple core integer number
                            string sThisCore = sCableNo + "-" + sCoreNo.PadLeft(2, '0');
                            bool bCoreExists = PartExists(sThisCore, iWebAppId);

                            //Get the material for this cable and the number of cores
                            rtnString clsCableMaterial = CableMaterialChild(sCableNo, iWebAppId);
                            string sExistingMaterialCode = clsCableMaterial.sReturnValue;
                            string sExistingMaterialName = clsCableMaterial.sReturnValueExtra1;
                            int iNoOfCores = 0;

                            if(clsCableMaterial.bReturnValue)
                            {
                                rtnInt rntCores = GetPartIntAttribute(sExistingMaterialCode, "NoOfCores", iWebAppId);
                                iNoOfCores = rntCores.iReturnValue;
                            }


                            switch (sAction)
                            {
                                case "ADD":
                                    if (!bCableExists)
                                    {
                                        sRowMsg = "The cable on row " + i + " of file " + sFile + " does not exist. You cannot terminate unless the cable exists.\r\n";
                                    }
                                    else
                                    {
                                        if (!bToPartExists)
                                        {
                                            if(iFLOrMat == 0)
                                                sRowMsg = "The 'To' Equipment " + sToEquipNo + " on row " + i + " of file " + sFile + " does not exist. Plant equipment must exist before you can create terminations on that equipment.\r\n";
                                            else
                                                sRowMsg = "The material code " + sToEquipNo + " on row " + i + " of file " + sFile + " does not exist. Material must exist before you can create terminations on that equipment.\r\n";
                                        }
                                        else
                                        {
                                            iIsNumber = 0;
                                            bool bCableCounter = int.TryParse(sCableNo.Substring(sCableNo.Length - 1, 1), out iIsNumber);
                                            if (!sCableNo.Substring(0, sCableNo.IndexOf("-")).Equals(sToEquipNo) || !bCableCounter)
                                            {
                                                sRowMsg = "The cable no " + sCableNo + " on row " + i + " of file " + sFile + " does not match the 'To' Equipment " + sToEquipNo + ". The cable no has the same prefix as the 'To' Equipment followed by a '-' and then P,C,I,D or E and a number.\r\n";
                                            }
                                            else
                                            {
                                                if (!bFromPartExists && !sFromEquipNo.Equals(""))
                                                    if (iFLOrMat == 0)
                                                        sRowMsg = "The 'From' Equipment " + sFromEquipNo + " on row " + i + " of file " + sFile + " does not exist. Plant equipment must exist before you can reate terminations on that equipment.\r\n";
                                                    else
                                                        sRowMsg = "The 'From' Equipment " + sFromEquipNo + " on row " + i + " of file " + sFile + " does not exist. Plant equipment must exist before you can reate terminations on that equipment.\r\n";
                                                else
                                                {
                                                    if(iCoreNo < 0)
                                                        sRowMsg = "The core " + iCoreNo + " for cable " + sCableNo + " on row " + i + " of file " + sFile + " is not allocated. You cannot perfrom any action unless the core number is allocated.\r\n";
                                                    else
                                                    {
                                                        if (bCoreExists)
                                                            sRowMsg = "The core " + iCoreNo + " for cable " + sCableNo + " already exists on row " + i + " of file " + sFile + ". You cannot add this row. Please change the action to 'Update' or 'Delete'.\r\n";
                                                        else
                                                        {
                                                            if (iCoreNo > iNoOfCores)
                                                            {
                                                                if (sExistingMaterialCode.Equals(""))
                                                                    sRowMsg = "The cable " + sCableNo + " on row " + i + " of file " + sFile + " does not have any allocated material. Therefore the number of cores is unknown and you cannot add any terminations until this is known.You cannot add cable mateial unless it exists.\r\n";
                                                                else
                                                                    sRowMsg = "The cable " + sCableNo + " on row " + i + " of file " + sFile + " has the material " + sExistingMaterialCode + " - " + sExistingMaterialName + " which has " + iNoOfCores + " cores. The core number on this row is " + iCoreNo + " which exceeds the number of cores in the cable.\r\n";
                                                            }
                                                            else
                                                            {
                                                                //If we have go to here then all items check out properly and we can add the termination row
                                                                int iTermFromLineNumber = 0;
                                                                iTermFromLineNumber = GetNewLineNumber(sFromEquipNo + "-" + sFromTermination, iWebAppId);
                                                                string sRtn = "";
                                                                if (!sFromEquipNo.Equals("") && !sFromTermination.Equals(""))
                                                                    sRtn = CreateCableTerminationLink2(sSessionId, sUserId, sCableNo, sFromEquipNo, iTermFromLineNumber.ToString(), "0",
                                                                                                          sFromTermination, sWireNo, sCoreNo, sCoreLabel, sWebAppId);
                                                                else
                                                                    sRtn = "Success";

                                                                if (!sRtn.Equals("Success"))
                                                                {
                                                                    sBody += sRtn;
                                                                }
                                                                else
                                                                {
                                                                    int iTermToLineNumber = 0;
                                                                    iTermToLineNumber = GetNewLineNumber(sToEquipNo + "-" + sToTermination, iWebAppId);
                                                                    string sRtn2 = CreateCableTerminationLink2(sSessionId, sUserId, sCableNo, sToEquipNo, iTermToLineNumber.ToString(), "1",
                                                                                                              sToTermination, sWireNo, sCoreNo, sCoreLabel, sWebAppId);

                                                                    if (!sRtn2.Equals("Success"))
                                                                    {
                                                                        sBody += sRtn2;
                                                                    }

                                                                }
                                                            }
                                                        }

                                                    }

                                                }
                                            }
                                        }

                                    }
                                    break;
                                case "UPDATE":
                                    if (!bCableExists)
                                    {
                                        sRowMsg = "The cable on row " + i + " of file " + sFile + " does not exist. You cannot terminate unless the cable exists.\r\n";
                                    }
                                    else
                                    {
                                        if (!bToPartExists)
                                        {
                                            if (iFLOrMat == 0)
                                                sRowMsg = "The 'To' Equipment " + sToEquipNo + " on row " + i + " of file " + sFile + " does not exist. Plant equipment must exist before you can create terminations on that equipment.\r\n";
                                            else
                                                sRowMsg = "The material code " + sToEquipNo + " on row " + i + " of file " + sFile + " does not exist. Material must exist before you can create terminations on that equipment.\r\n";
                                        }
                                        else
                                        {
                                            iIsNumber = 0;
                                            bool bCableCounter = int.TryParse(sCableNo.Substring(sCableNo.Length - 1, 1), out iIsNumber);
                                            if (!sCableNo.Substring(0, sCableNo.IndexOf("-")).Equals(sToEquipNo) || !bCableCounter)
                                            {
                                                sRowMsg = "The cable no " + sCableNo + " on row " + i + " of file " + sFile + " does not match the 'To' Equipment " + sToEquipNo + ". The cable no has the same prefix as the 'To' Equipment followed by a '-' and then P,C,I,D or E and a number.\r\n";
                                            }
                                            else
                                            {
                                                if (!bFromPartExists && !sFromEquipNo.Equals(""))
                                                    sRowMsg = "The 'From' Equipment " + sFromEquipNo + " on row " + i + " of file " + sFile + " does not exist. Plant equipment must exist before you can reate terminations on that equipment.\r\n";
                                                else
                                                {
                                                    if (iCoreNo < 0)
                                                        sRowMsg = "The core " + iCoreNo + " for cable " + sCableNo + " on row " + i + " of file " + sFile + " is not allocated. You cannot perfrom any action unless the core number is allocated.\r\n";
                                                    else
                                                    {
                                                        if (!bCoreExists)
                                                            sRowMsg = "The core " + iCoreNo + " for cable " + sCableNo + " on row " + i + " of file " + sFile + " does not exist. You cannot modify this row. Please change the action to 'Add'.\r\n";
                                                        else
                                                        {
                                                            if (iCoreNo > iNoOfCores)
                                                            {
                                                                if (sExistingMaterialCode.Equals(""))
                                                                    sRowMsg = "The cable " + sCableNo + " on row " + i + " of file " + sFile + " does not have any allocated material. Therefore the number of cores is unknown and you cannot add any terminations until this is known.You cannot add cable mateial unless it exists.\r\n";
                                                                else
                                                                    sRowMsg = "The cable " + sCableNo + " on row " + i + " of file " + sFile + " has the material " + sExistingMaterialCode + " - " + sExistingMaterialName + " which has " + iNoOfCores + " cores. The core number on this row is " + iCoreNo + " which exceeds the number of cores in the cable.\r\n";
                                                            }
                                                            else
                                                            {
                                                                //If we have go to here then all items check out properly and we can add the termination row
                                                                int iTermFromLineNumber = 0;
                                                                iTermFromLineNumber = GetNewLineNumber(sFromEquipNo, iWebAppId);
                                                                string sRtn = "";
                                                                if (!sFromEquipNo.Equals("") && !sFromTermination.Equals(""))
                                                                    sRtn = UpdateCableTerminationLink2(sSessionId, sUserId, sCableNo, sFromEquipNo, iTermFromLineNumber.ToString(), "0",
                                                                                                              sFromTermination, sWireNo, sCoreNo, sCoreLabel, sWebAppId);
                                                                else
                                                                    sRtn = "Success";

                                                                if (!sRtn.Equals("Success"))
                                                                {
                                                                    sBody += sRtn;
                                                                }
                                                                else
                                                                {
                                                                    int iTermToLineNumber = 0;
                                                                    iTermToLineNumber = GetNewLineNumber(sFromEquipNo, iWebAppId);
                                                                    string sRtn2 = UpdateCableTerminationLink2(sSessionId, sUserId, sCableNo, sToEquipNo, iTermToLineNumber.ToString(), "1",
                                                                                                              sToTermination, sWireNo, sCoreNo, sCoreLabel, sWebAppId);

                                                                    if (!sRtn2.Equals("Success"))
                                                                    {
                                                                        sBody += sRtn2;
                                                                    }

                                                                }
                                                            }
                                                        }

                                                    }

                                                }
                                            }
                                        }

                                    }
                                    break;
                                case "DELETE":
                                    if (!bCableExists)
                                    {
                                        sRowMsg = "The cable on row " + i + " of file " + sFile + " does not exist. You cannot delete a termination unless the cable exists.\r\n";
                                    }
                                    else
                                    {
                                        if (!bCoreExists)
                                            sRowMsg = "The core " + iCoreNo + " for cable " + sCableNo + " on row " + i + " of file " + sFile + " does not exist. You cannot delete this teermination.\r\n";
                                        else
                                        {

                                            string sCableCheckInComments = "Removing link between the cable and the core " + sCoreNo;
                                            //Note that wqe have to send across the core with the cable prefix and the apdded core number which is sThisCore
                                            rtnInt rtnCoreLinkExists = PartPartLinkExists(sCableNo, sThisCore, iWebAppId);
                                            bool bCoreinkExist = rtnCoreLinkExists.bReturnValue;
                                            if(bCoreinkExist)
                                            {
                                                string sDeleteLineNumber = rtnCoreLinkExists.iReturnValue.ToString();

                                                string sRtn2 = DeletePartToPartLinkByLineNumber(sSessionId, sUserId, sFullName, sDeleteLineNumber, sCableNo, sCoreNo, sCableCheckInComments, sWebAppId);
                                                if (!sRtn2.Equals("Success"))
                                                {
                                                    sBody += sRtn2;
                                                }

                                            }
                                        }
                                    }
                                    break;
                                default:
                                    sRowMsg = "The action must be one of ADD, UPDATE OR DELETE. Row " + i + " of file " + sFile + " cannot be processed.\r\n";
                                    break;
                            }

                            sBody += sRowMsg;

                        }
                    }

                    xlWorkbook.Close(true);
                    xlWbks.Close();
                    xlApp.Quit();

                    while (System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp) != 0) ;
                    while (System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbks) != 0) ;
                    while (System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkbook) != 0) ;
                    while (System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorksheet) != 0) ;
                    while (System.Runtime.InteropServices.Marshal.ReleaseComObject(xlRange) != 0) ;
                    xlApp = null;
                    xlWbks = null;
                    xlWorkbook = null;
                    xlWorksheet = null;
                    xlRange = null;

                    //Now email the user
                    string sSubject = "Processing of File " + sFile;
                    if (sBody.Length == 0)
                        sBody = "No issues.";
                    sBody = "File " + sFile + " was processed with the following issues.\r\n" + sBody;
                    //                    emailmessage(sSessionId, sUserId, sSubject, sBody, " ", sRecipeints, "", "", sWebAppId);

                    return "Success^" + sBody;
                }
            }
            catch (Exception ex)
            {
                return "Failure:" + ex.Message + "^";
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();

                System.Diagnostics.Process[] excelProcs = System.Diagnostics.Process.GetProcessesByName("EXCEL");
                foreach (System.Diagnostics.Process proc in System.Diagnostics.Process.GetProcessesByName("EXCEL"))
                {
                    proc.Kill();
                }
            }
        }

        public string RemoveInvalidCharacters(string sInputString)
        {
            string sOutputString = "";

            sOutputString = sInputString.Replace("<", "_");
            sOutputString = sOutputString.Replace(">", "_");
            sOutputString = sOutputString.Replace(":", "_");
            sOutputString = sOutputString.Replace("\"", "_");
            sOutputString = sOutputString.Replace("/", "_");
            sOutputString = sOutputString.Replace("\\", "_");
            sOutputString = sOutputString.Replace("|", "_");
            sOutputString = sOutputString.Replace("?", "_");
            sOutputString = sOutputString.Replace("*", "_");
            sOutputString = sOutputString.Replace("\r\n", "");
            sOutputString = sOutputString.Replace("\r", "");
            sOutputString = sOutputString.Replace("\n", "");

            return sOutputString;
        }

        private ExampleService.MyJavaService3Client GetWCService()
        {
            Environment env = new Environment();
            string sCertVal = env.Get_Environment_String_Value("CertificateValue");

            ExampleService.MyJavaService3Client client2 = new ExampleService.MyJavaService3Client();

            client2.ClientCredentials.UserName.UserName = "benmess";
            client2.ClientCredentials.UserName.Password = "mo9anaapr!";
            client2.ClientCredentials.ServiceCertificate.SetDefaultCertificate(StoreLocation.CurrentUser,
                                                                              StoreName.TrustedPeople, X509FindType.FindBySubjectName,
                                                                              sCertVal); //Make this read from an environment file so we can change between dev and production
            return client2;

        }

        public string[] Extract_Values(string sValues)
		{
			string[] sLocalArray = sValues.Split('^');
			return sLocalArray;
		}

        public string emailmessage(string sSessionId, string sUserId, string sSubject, string sBody, string sAttachments, string sRecipients, string sCCRecipients, string sBCCRecipients, string sWebAppId)
        {
            if (!IsExternalUserValid(sSessionId, sUserId, Convert.ToInt16(sWebAppId)))
            {
                return "User " + sUserId + " is not logged in";
            }
            else
            {
                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                char[] charSeparators = new char[] { '^' };
                string[] sAttachArr = null;
                sCCRecipients = sCCRecipients.Trim();
                sBCCRecipients = sBCCRecipients.Trim();
                if (sAttachments != " ")
                    sAttachArr = sAttachments.Split(charSeparators, StringSplitOptions.RemoveEmptyEntries);

//                String sSubject, String sBody, String[] sAttachments, String sRecipients, String sCCRecipients, String sBCCRecipients, string sWebAppId
                client2.emailmessage(sSubject, sBody, sAttachArr, sRecipients, sCCRecipients, sBCCRecipients, Convert.ToInt16(sWebAppId));
                return "Success";
            }
        }

        public bool FileExists(string sFileNameAndPath)
        {
            if (File.Exists(sFileNameAndPath))
                return true;
            else
                return false;
        }

        public bool PartExists(String sPartNo, int iWebAppId)
        {
            RecordSet rst = new RecordSet();
            rst.SetWebApp(iWebAppId);
            string sSQL = "select * from vwWindchillLatestPart where WTPartNumber = '" + sPartNo + "' COLLATE SQL_Latin1_General_CP1_CI_AS";
            DataSet ds = rst.OpenRecordset(sSQL, rst.SqlConnectionStr());
            bool bRtn = false;

            if (rst.m_RecordCount > 0)
            {
                bRtn = true;
            }


            ds.Dispose();

            return bRtn;
        }

        public rtnInt PartIOLinkExists(String sParentPartNo, String sChildPartNo, string sIOType, string sIOTag, int iWebAppId)
        {
            RecordSet rst = new RecordSet();
            int? iRtnValueDB = -1;
            int iRtnValue = -1;
            rtnInt rtnCls = new rtnInt();
            rst.SetWebApp(iWebAppId);
            string sSQL = "select VIO1.* " +
                          "from vwWindchillPartUsageStringAttributes VIO1, vwWindchillPartUsageStringAttributes VIO2 " + 
                          "where VIO1.PMAPartNumber = '" + sParentPartNo + "' " +
                          "and VIO1.PMBPartNumber = '" + sChildPartNo + "' " +
                          "and VIO1.name = 'IOType' " +
                          "and VIO1.value = '" + sIOType + "' " +
                          "and VIO1.PULId = VIO2.PULId " +
                          "and VIO2.name = 'IOTag'           " +
                          "and VIO2.value = '" + sIOTag + "' ";

            //select * from vwWindchillLatestPart where WTPartNumber = '" + sPartNo + "'";
            DataSet ds = rst.OpenRecordset(sSQL, rst.SqlConnectionStr());
            bool bRtn = false;

            if (rst.m_RecordCount > 0)
            {
                iRtnValueDB = rst.Get_Int(ds, "LineNumber", 0);
                if (iRtnValueDB == null)
                    iRtnValue = -1;
                else
                    iRtnValue = (int)iRtnValueDB;

                rtnCls.bReturnValue = true;
                rtnCls.iReturnValue = iRtnValue;
                ds.Dispose();
            }
            else
            {
                rtnCls.bReturnValue = false;
                rtnCls.iReturnValue = iRtnValue;
            }

            return rtnCls;
        }

        public rtnInt PartIOLinkExistsNoChildRequired(String sParentPartNo, string sIOType, string sIOTag, int iWebAppId)
        {
            RecordSet rst = new RecordSet();
            int? iRtnValueDB = -1;
            int iRtnValue = -1;
            rtnInt rtnCls = new rtnInt();
            rst.SetWebApp(iWebAppId);
            string sSQL = "select VIO1.* " +
                          "from vwWindchillPartUsageStringAttributes VIO1, vwWindchillPartUsageStringAttributes VIO2 " +
                          "where VIO1.PMAPartNumber = '" + sParentPartNo + "' " +
                          "and VIO1.name = 'IOType' " +
                          "and VIO1.value = '" + sIOType + "' " +
                          "and VIO1.PULId = VIO2.PULId " +
                          "and VIO2.name = 'IOTag'           " +
                          "and VIO2.value = '" + sIOTag + "' collate SQL_Latin1_General_CP1_CI_AS";

            //select * from vwWindchillLatestPart where WTPartNumber = '" + sPartNo + "'";
            DataSet ds = rst.OpenRecordset(sSQL, rst.SqlConnectionStr());
            bool bRtn = false;

            if (rst.m_RecordCount > 0)
            {
                iRtnValueDB = rst.Get_Int(ds, "LineNumber", 0);
                if (iRtnValueDB == null)
                    iRtnValue = -1;
                else
                    iRtnValue = (int)iRtnValueDB;

                rtnCls.bReturnValue = true;
                rtnCls.iReturnValue = iRtnValue;
                ds.Dispose();
            }
            else
            {
                rtnCls.bReturnValue = false;
                rtnCls.iReturnValue = iRtnValue;
            }

            return rtnCls;
        }

        // iToOrFrom = 1 - To End
        //           = 0 - From End
        public rtnInt CablePartLinkExists(String sEquipPartNo, String sCablePartNo, int iToOrFrom, int iWebAppId)
        {
            RecordSet rst = new RecordSet();
            int? iRtnValueDB = -1;
            int iRtnValue = -1;
            rtnInt rtnCls = new rtnInt();
            rst.SetWebApp(iWebAppId);
            string sSQL = "select VIO1.* " +
                          "from vwWindchillPartUsageStringAttributes VIO1 " +
                          "where VIO1.PMAPartNumber = '" + sEquipPartNo + "' " +
                          "and VIO1.PMBPartNumber = '" + sCablePartNo + "' " +
                          "and VIO1.name = 'ToOrFrom' " +
                          "and VIO1.value = '" + iToOrFrom + "'";

            //select * from vwWindchillLatestPart where WTPartNumber = '" + sPartNo + "'";
            DataSet ds = rst.OpenRecordset(sSQL, rst.SqlConnectionStr());
            bool bRtn = false;

            if (rst.m_RecordCount > 0)
            {
                iRtnValueDB = rst.Get_Int(ds, "LineNumber", 0);
                if (iRtnValueDB == null)
                    iRtnValue = -1;
                else
                    iRtnValue = (int)iRtnValueDB;

                rtnCls.bReturnValue = true;
                rtnCls.iReturnValue = iRtnValue;
                ds.Dispose();
            }
            else
            {
                rtnCls.bReturnValue = false;
                rtnCls.iReturnValue = iRtnValue;
            }

            return rtnCls;
        }

        // iToOrFrom = 1 - To End
        //           = 0 - From End
        public rtnString CablePartLinkParent(String sCablePartNo, int iToOrFrom, int iWebAppId)
        {
            RecordSet rst = new RecordSet();
            string sRtnValue = "";
            rtnString rtnCls = new rtnString();
            rst.SetWebApp(iWebAppId);
            string sSQL = "select isnull(VIO1.PMAPartNumber,'') as PMAPartNumber, isnull(LineNumber, -1) as LineNumber " +
                          "from vwWindchillPartUsageStringAttributes VIO1 " +
                          "where VIO1.PMBPartNumber = '" + sCablePartNo + "' " +
                          "and VIO1.name = 'ToOrFrom' " +
                          "and VIO1.value = '" + iToOrFrom + "'";

            //select * from vwWindchillLatestPart where WTPartNumber = '" + sPartNo + "'";
            DataSet ds = rst.OpenRecordset(sSQL, rst.SqlConnectionStr());

            if (rst.m_RecordCount > 0)
            {
                sRtnValue = rst.Get_NVarchar(ds, "PMAPartNumber", 0);
                int iLineNumber = rst.Get_Int(ds, "LineNumber", 0);
                rtnCls.bReturnValue = true;
                rtnCls.sReturnValue = sRtnValue;
                rtnCls.iLineNumber = iLineNumber;
                ds.Dispose();
            }
            else
            {
                rtnCls.bReturnValue = false;
                rtnCls.sReturnValue = sRtnValue;
            }

            return rtnCls;
        }

        public rtnInt CableMaterialLinkExists(String sCablePartNo, String sMaterialPartNo, int iWebAppId)
        {
            RecordSet rst = new RecordSet();
            int iRtnValue = -1;
            rtnInt rtnCls = new rtnInt();
            rst.SetWebApp(iWebAppId);
            string sSQL = "select isnull(VIO1.LineNumber,-1) as LineNumber " +
                          "from vwWindchillPartUsageInfo VIO1 " + 
                          "where VIO1.PMAPartNumber = '" + sCablePartNo + "' " +
                          "and VIO1.PMBPartNumber = '" + sMaterialPartNo + "'";

            //select * from vwWindchillLatestPart where WTPartNumber = '" + sPartNo + "'";
            DataSet ds = rst.OpenRecordset(sSQL, rst.SqlConnectionStr());
            bool bRtn = false;

            if (rst.m_RecordCount > 0)
            {
                iRtnValue = rst.Get_Int(ds, "LineNumber", 0);
                rtnCls.bReturnValue = true;
                rtnCls.iReturnValue = iRtnValue;
                ds.Dispose();
            }
            else
            {
                rtnCls.bReturnValue = false;
                rtnCls.iReturnValue = iRtnValue;
            }

            return rtnCls;
        }

        public rtnInt PartPartLinkExists(String sParentPartNo, String sChiildPartNo, int iWebAppId)
        {
            RecordSet rst = new RecordSet();
            int iRtnValue = -1;
            rtnInt rtnCls = new rtnInt();
            rst.SetWebApp(iWebAppId);
            string sSQL = "select isnull(VIO1.LineNumber,-1) as LineNumber " +
                          "from vwWindchillPartUsageInfo VIO1 " +
                          "where VIO1.PMAPartNumber = '" + sParentPartNo + "' " +
                          "and VIO1.PMBPartNumber = '" + sChiildPartNo + "'";

            //select * from vwWindchillLatestPart where WTPartNumber = '" + sPartNo + "'";
            DataSet ds = rst.OpenRecordset(sSQL, rst.SqlConnectionStr());
            bool bRtn = false;

            if (rst.m_RecordCount > 0)
            {
                iRtnValue = rst.Get_Int(ds, "LineNumber", 0);
                rtnCls.bReturnValue = true;
                rtnCls.iReturnValue = iRtnValue;
                ds.Dispose();
            }
            else
            {
                rtnCls.bReturnValue = false;
                rtnCls.iReturnValue = iRtnValue;
            }

            return rtnCls;
        }

        public rtnInt GetPartIntAttribute(String sPartNo, String sAttributeName, int iWebAppId)
        {
            RecordSet rst = new RecordSet();
            int iRtnValue = -1;
            rtnInt rtnCls = new rtnInt();
            rst.SetWebApp(iWebAppId);
            string sSQL = "select value from vwWindchillPartIntegerAttributes where WTPartNumber = '" + sPartNo + "' and name = '" + sAttributeName + "'";

            //select * from vwWindchillLatestPart where WTPartNumber = '" + sPartNo + "'";
            DataSet ds = rst.OpenRecordset(sSQL, rst.SqlConnectionStr());
            bool bRtn = false;

            if (rst.m_RecordCount > 0)
            {
                iRtnValue = rst.Get_Int(ds, "value", 0);
                rtnCls.bReturnValue = true;
                rtnCls.iReturnValue = iRtnValue;
                ds.Dispose();
            }
            else
            {
                rtnCls.bReturnValue = false;
                rtnCls.iReturnValue = iRtnValue;
            }

            return rtnCls;
        }

        public rtnString GetPartStringAttribute(String sPartNo, String sAttributeName, int iWebAppId)
        {
            RecordSet rst = new RecordSet();
            string sRtnValue = "";
            rtnString rtnCls = new rtnString();
            rst.SetWebApp(iWebAppId);
            string sSQL = "select value from vwWindchillPartStringAttributes where WTPartNumber = '" + sPartNo + "' and name = '" + sAttributeName + "'";

            //select * from vwWindchillLatestPart where WTPartNumber = '" + sPartNo + "'";
            DataSet ds = rst.OpenRecordset(sSQL, rst.SqlConnectionStr());
            bool bRtn = false;

            if (rst.m_RecordCount > 0)
            {
                sRtnValue = rst.Get_NVarchar(ds, "value", 0);
                rtnCls.bReturnValue = true;
                rtnCls.sReturnValue = sRtnValue;
                ds.Dispose();
            }
            else
            {
                rtnCls.bReturnValue = false;
                rtnCls.sReturnValue = sRtnValue;
            }

            return rtnCls;
        }

        public rtnString CableMaterialChild(String sCablePartNo, int iWebAppId)
        {
            RecordSet rst = new RecordSet();
            string sRtnValue = "";
            string sMaterialName = "";
            rtnString rtnCls = new rtnString();
            int iLineNumber = 0;
            rst.SetWebApp(iWebAppId);
            string sSQL = "select isnull(VIO1.PMBPartNumber,'') as PMBPartNumber, isnull(LineNumber, 0) as LineNumber, PMBPartName as MaterialName " +
                          "from vwWindchillPartUsageInfo VIO1 " +
                          "where VIO1.PMAPartNumber = '" + sCablePartNo + "' " +
                          "and VIO1.PBPartType = 'local.rs.vsrs05.Regain.AutoNumberedPart'";

            //select * from vwWindchillLatestPart where WTPartNumber = '" + sPartNo + "'";
            DataSet ds = rst.OpenRecordset(sSQL, rst.SqlConnectionStr());
            bool bRtn = false;

            if (rst.m_RecordCount > 0)
            {
                sRtnValue = rst.Get_NVarchar(ds, "PMBPartNumber", 0);
                iLineNumber = rst.Get_Int(ds, "LineNumber", 0);
                sMaterialName = rst.Get_NVarchar(ds, "MaterialName", 0);
                rtnCls.bReturnValue = true;
                rtnCls.sReturnValue = sRtnValue;
                rtnCls.iLineNumber = iLineNumber;
                rtnCls.sReturnValueExtra1 = sMaterialName;
            }
            else
            {
                rtnCls.bReturnValue = false;
                rtnCls.sReturnValue = sRtnValue;
                rtnCls.iLineNumber = iLineNumber;
                rtnCls.sReturnValueExtra1 = sMaterialName;
            }

            ds.Dispose();
            return rtnCls;
        }

        public rtnInt GetCableNoOfCores(String sCablePartNo, int iWebAppId)
        {
            RecordSet rst = new RecordSet();
            string sRtnValue = "";
            string sMaterialName = "";
            rtnInt rtnCls = new rtnInt();
            int iNoOfCores = 0;
            rst.SetWebApp(iWebAppId);
            string sSQL = "select isnull(VIO1.PMBPartNumber,'') as PMBPartNumber, IA.value as NoOfCores " +
                          "from vwWindchillPartUsageInfo VIO1, vwWindchillPartIntegerAttributes IA " +
                          "where VIO1.PMAPartNumber = '" + sCablePartNo + "' " +
                          "and VIO1.PBPartType = 'local.rs.vsrs05.Regain.AutoNumberedPart' " +
                          "and isnull(VIO1.PMBPartNumber,'') = IA.WTPartNumber " +
                          "and IA.name = 'NoOfCores'";

            //select * from vwWindchillLatestPart where WTPartNumber = '" + sPartNo + "'";
            DataSet ds = rst.OpenRecordset(sSQL, rst.SqlConnectionStr());
            bool bRtn = false;

            if (rst.m_RecordCount > 0)
            {
                sRtnValue = rst.Get_NVarchar(ds, "PMBPartNumber", 0);
                iNoOfCores = rst.Get_Int(ds, "NoOfCores", 0);
                rtnCls.bReturnValue = true;
                rtnCls.iReturnValue = iNoOfCores;
            }
            else
            {
                rtnCls.bReturnValue = false;
                rtnCls.iReturnValue = iNoOfCores;
            }

            ds.Dispose();
            return rtnCls;
        }

        public rtnString GetParentPartOfType(String sChildPart, string sType, int iWebAppId)
        {
            RecordSet rst = new RecordSet();
            string sRtnValue = "";
            rtnString rtnCls = new rtnString();
            rst.SetWebApp(iWebAppId);
            string sSQL = "exec SP_GetWindchillParentParts @pvchChildPartNumber = '" + sChildPart + "', @pvchParentPartType = '" + sType + "'";

            //select * from vwWindchillLatestPart where WTPartNumber = '" + sPartNo + "'";
            DataSet ds = rst.OpenRecordset(sSQL, rst.SqlConnectionStr());
            bool bRtn = false;

            if (rst.m_RecordCount > 0)
            {
                sRtnValue = rst.Get_NVarchar(ds, "WTPartNumber", 0);
                rtnCls.bReturnValue = true;
                rtnCls.sReturnValue = sRtnValue;
                ds.Dispose();
            }
            else
            {
                rtnCls.bReturnValue = false;
                rtnCls.sReturnValue = sRtnValue;
            }

            return rtnCls;
        }

        public rtnString GetChildPartOfType(String sParentPart, string sType, string sNameContains, int iWebAppId)
        {
            RecordSet rst = new RecordSet();
            string sRtnValue = "";
            string sName = "";
            int iLineNumber = -1;
            rtnString rtnCls = new rtnString();
            rst.SetWebApp(iWebAppId);
            int i;
            string sSQL = "exec SP_GetWindchillChildPartsWithLineNumber @pvchParentPartNumber = '" + sParentPart + "', @pvchChildPartType = '" + sType + "'";

            //select * from vwWindchillLatestPart where WTPartNumber = '" + sPartNo + "'";
            DataSet ds = rst.OpenRecordset(sSQL, rst.SqlConnectionStr());
            bool bRtn = false;

            if (rst.m_RecordCount > 0)
            {
                for (i = 0; i < rst.m_RecordCount; i++)
                {
                    sRtnValue = rst.Get_NVarchar(ds, "WTPartNumber", i);
                    sName = rst.Get_NVarchar(ds, "PartName", i);
                    iLineNumber = rst.Get_Int(ds, "LineNumber", i);
                    if (!sNameContains.Equals(""))
                    {
                        if(sName.Contains(sNameContains))
                        {
                            rtnCls.bReturnValue = true;
                            rtnCls.sReturnValue = sRtnValue;
                            rtnCls.iLineNumber = iLineNumber;
                            ds.Dispose();
                            return rtnCls;
                        }
                    }
                    else
                    {
                        rtnCls.bReturnValue = true;
                        rtnCls.sReturnValue = sRtnValue;
                        rtnCls.iLineNumber = iLineNumber;
                        ds.Dispose();
                        return rtnCls;
                    }
                }
                ds.Dispose();
            }
            else
            {
                rtnCls.bReturnValue = false;
                rtnCls.sReturnValue = sRtnValue;
                rtnCls.iLineNumber = -1;
            }

            return rtnCls;
        }

        public rtnString GetPlantJobFolder(int iJob, int iWebAppId)
        {
            RecordSet rst = new RecordSet();
            string sRtnValue = "";
            string sName = "";
            rtnString rtnCls = new rtnString();
            rst.SetWebApp(iWebAppId);
            string sSQL = "exec SP_JobPlantFolder " + iJob;

            //select * from vwWindchillLatestPart where WTPartNumber = '" + sPartNo + "'";
            DataSet ds = rst.OpenRecordset(sSQL, rst.SqlConnectionStr());
            bool bRtn = false;

            if (rst.m_RecordCount > 0)
            {
                sRtnValue = rst.Get_NVarchar(ds, "FolderPath", 0);
                rtnCls.bReturnValue = true;
                rtnCls.sReturnValue = sRtnValue;
                ds.Dispose();
            }
            else
            {
                rtnCls.bReturnValue = false;
                rtnCls.sReturnValue = sRtnValue;
            }

            return rtnCls;
        }

        public rtnString GetProductFromJob(string sJob, int iProdOrLib, int iWebAppId)
        {
            RecordSet rst = new RecordSet();
            string sRtnValue = "";
            string sName = "";
            rtnString rtnCls = new rtnString();
            rst.SetWebApp(iWebAppId);
            string sSQL = "exec SP_GetWindchillDocumentProduct '" + sJob + "', " + iProdOrLib;

            //select * from vwWindchillLatestPart where WTPartNumber = '" + sPartNo + "'";
            DataSet ds = rst.OpenRecordset(sSQL, rst.SqlConnectionStr());
            bool bRtn = false;

            if (rst.m_RecordCount > 0)
            {
                sRtnValue = rst.Get_NVarchar(ds, "ProductName", 0);
                rtnCls.bReturnValue = true;
                rtnCls.sReturnValue = sRtnValue;
                ds.Dispose();
            }
            else
            {
                rtnCls.bReturnValue = false;
                rtnCls.sReturnValue = sRtnValue;
            }

            return rtnCls;
        }


        public int GetNewLineNumber(String sParentPartNo, int iWebAppId)
        {
            RecordSet rst = new RecordSet();
            rst.SetWebApp(iWebAppId);
            int iRtnValue = 10;
            string sSQL = "select TOP 1 cast((round(isnull(UL.valueB7,-1)/10, 0) + 1) * 10 as bigint) as NewLineNumber " +
                    "from wcadmin.wcadmin.WTPartUsageLink UL, vwWindchillLatestPart LP " +
                    "where UL.idA3A5 = LP.PartId " +
                    "and LP.WTPartNumber = '" + sParentPartNo + "' " +
                    "order by isnull(UL.valueB7,-1) desc";

            //select * from vwWindchillLatestPart where WTPartNumber = '" + sPartNo + "'";
            DataSet ds = rst.OpenRecordset(sSQL, rst.SqlConnectionStr());

            if (rst.m_RecordCount > 0)
            {
                iRtnValue = rst.Get_Int(ds, "NewLineNumber", 0);
                ds.Dispose();
            }

            return iRtnValue;
        }

        public double GetUsageLinkExistingQty(String sParentPartNo, String sChildPartNo, long lLineNumber, int iWebAppId)
        {
            RecordSet rst = new RecordSet();
            rst.SetWebApp(iWebAppId);
            double dRtnValue = -1;
            string sSQL = "select Quantity from vwWindchillPartUsageInfo where PMAPartNumber = '" + sParentPartNo + "' " +
                    "and PMBPartNumber = '" + sChildPartNo + "' " +
                    "and isnull(LineNumber, 0) = " + lLineNumber;

            //select * from vwWindchillLatestPart where WTPartNumber = '" + sPartNo + "'";
            DataSet ds = rst.OpenRecordset(sSQL, rst.SqlConnectionStr());

            if (rst.m_RecordCount > 0)
            {
                dRtnValue = rst.Get_Float(ds, "Quantity", 0);
                ds.Dispose();
            }

            return dRtnValue;
        }

        public rtnString GetCableCoreLabel(String sCableCoreNo, int iWebAppId)
        {
            RecordSet rst = new RecordSet();
            string sRtnValue = "";
            rtnString rtnCls = new rtnString();
            rst.SetWebApp(iWebAppId);
            string sSQL = "select value as CoreLabel " +
                          "from vwWindchillPartStringAttributes " +
                          "where WTPartNumber = '" + sCableCoreNo + "' " +
                          "and name = 'CoreLabel'";

            //select * from vwWindchillLatestPart where WTPartNumber = '" + sPartNo + "'";
            DataSet ds = rst.OpenRecordset(sSQL, rst.SqlConnectionStr());
            bool bRtn = false;

            if (rst.m_RecordCount > 0)
            {
                sRtnValue = rst.Get_NVarchar(ds, "CoreLabel", 0);
                rtnCls.bReturnValue = true;
                rtnCls.sReturnValue = sRtnValue;
                ds.Dispose();
            }
            else
            {
                rtnCls.bReturnValue = false;
                rtnCls.sReturnValue = sRtnValue;
            }

            return rtnCls;
        }

        public rtnTerms[] GetTerminations(string sCableNo, int iWebAppId)
        {
            RecordSet rst = new RecordSet();
            rtnString rtnCls = new rtnString();
            rst.SetWebApp(iWebAppId);
            int iNoOfCores = 0;
            rtnInt rtnCores = new rtnInt();

            rtnCores = GetCableNoOfCores(sCableNo, iWebAppId);
            bool bCableExists = false;

            if (rtnCores.bReturnValue)
            {
                iNoOfCores = rtnCores.iReturnValue;
                bCableExists = true;
            }
            else
                iNoOfCores = 1;

            rtnTerms[] rtnTerminations = new rtnTerms[iNoOfCores];

            if (bCableExists)
            {
                string sSQL = "exec SP_GetWindchillTerminations '" + sCableNo + "', " + iNoOfCores;

                //select * from vwWindchillLatestPart where WTPartNumber = '" + sPartNo + "'";
                DataSet ds = rst.OpenRecordset(sSQL, rst.SqlConnectionStr());

                if (rst.m_RecordCount > 0)
                {
                    for (int i = 0; i < rst.m_RecordCount; i++)
                    {
                        rtnTerminations[i] = new rtnTerms();
                        rtnTerminations[i].iCoreNo = rst.Get_Int(ds, "CoreNo", i);
                        rtnTerminations[i].sFromTermination = rst.Get_NVarchar(ds, "FromTermination", i);
                        rtnTerminations[i].iFromLineNumber = rst.Get_Int(ds, "FromLineNumber", i);
                        rtnTerminations[i].sToTermination = rst.Get_NVarchar(ds, "ToTermination", i);
                        rtnTerminations[i].iToLineNumber = rst.Get_Int(ds, "ToLineNumber", i);
                        rtnTerminations[i].sWireNo = rst.Get_NVarchar(ds, "WireNo", i);
                        rtnTerminations[i].sCoreLabel = rst.Get_NVarchar(ds, "CoreLabel", i);
                        rtnTerminations[i].bReturnValue = true;
                    }
                    ds.Dispose();
                }
                else
                {
                    rtnTerminations[0].bReturnValue = false;
                }
            }
            else
            {
                rtnTerminations[0].bReturnValue = false;
            }

            return rtnTerminations;
        }

        public bool SetAlbaPLC_LockedInfo(string sChassis, int iSlot, int iChannel, string sEquipNo, string sIOType, string sIOTag, int iWebAppId)
        {
            string sSQL;
            RecordSet rst = new RecordSet();
            rst.SetWebApp(iWebAppId);

            if(sIOType.Length == 0)
                sSQL = "UPDATE AlbaPLCInfo SET LockedRegainId = '" + sEquipNo + "', LockedIOType = '" + sIOType + "', LockedIOTag = '" + sIOTag + "' " +
                       " WHERE ChassisId = '" + sChassis + "' " +
                       "and Slot = " + iSlot + " " +
                       "and Channel = " + iChannel;
            else
                sSQL = "UPDATE AlbaPLCInfo SET LockedRegainId = '" + sEquipNo + "', LockedIOType = '" + sIOType + "', LockedIOTag = '" + sIOTag + "' " +
                   " WHERE ChassisId = '" + sChassis + "' " +
                   "and Slot = " + iSlot + " " +
                   "and Channel = " + iChannel + " " +
                   "and IOType = 'PLC connection point, " + sIOType + "'";

            bool bRtn = rst.ExecuteSQL(sSQL);
            return bRtn;
        }


    }
}
