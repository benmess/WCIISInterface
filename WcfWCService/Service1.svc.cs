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
using DocumentFormat.OpenXml.Packaging;

namespace WcfWCService
{
    // NOTE: You can use the "Rename" command on the "Refactor" menu to change the class name "Service1" in code, svc and config file together.
    // NOTE: In order to launch WCF Test Client for testing this service, please select Service1.svc or Service1.svc.cs at the Solution Explorer and start debugging.
    [ServiceBehavior(Namespace = "http://regain.com/rest")]
    public class Service1 : IService1
    {
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
                catch(Exception ex)
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

        public string SetDocPartDescribedByLink(string sSessionId, string sUserId,  string sDocNo, string sPartNo, string sWebAppId)
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
                                               string sCheckInComments, string sLineNumber,  string sWebAppId)
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
                if(sReturn == "Success")
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

        public string SetSupplierToPartRef(string sSessionId, string sUserId, string sFullName, string sSupplierNo, string sPartNo, string sSupplierPartNo, string sPartDocRefLinkType, string sCheckinComments, string sWebAppId)
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
                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();

                sAttributeNames[0] = "SupplierPartNo";
                sAttributeValues[0] = sSupplierPartNo;
                sAttributeTypes[0] = "string";

                return client2.setpartreferencedbydoclinkwithattributes(sFullName, sSupplierNo, sPartNo, sAttributeNames, sAttributeValues, sAttributeTypes, sCheckinComments, sPartDocRefLinkType, Convert.ToInt16(sWebAppId));
            }
        }

        //THis uses the special part reference type that takes attributes. Attributes for relationships cannot be updated in native Windchill with the exception of the Part to Part link.
        public string UpdateSupplierToPartRef(string sSessionId, string sUserId, string sFullName, string sSupplierNo, string sPartNo, string sSupplierPartNo, string sCheckinComments, string sWebAppId)
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
                Update_User_Time(sUserId, sSessionId);
                ExampleService.MyJavaService3Client client2 = GetWCService();

                sAttributeNames[0] = "SupplierPartNo";
                sAttributeValues[0] = sSupplierPartNo;
                sAttributeTypes[0] = "string";

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

                    return client2.setpartusageattributesfromlinenumber(sParentPartNo, sChildPartNo, sFullName,lLineNumber, sAttributeNames, sAttributeValues, sAttributeTypes, Convert.ToInt16(sWebAppId));

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
                    sAttributeNames[i*3 + 1] = "InvoiceStatus";
                    sAttributeNames[i*3 + 2] = "QtyInvoiced";

                    sAttributeValues[i * 3] = sFullName;
                    sAttributeValues[i*3 + 1] = sInvoiceStatuses[i];
                    sAttributeValues[i*3 + 2] = sQtysInvoiced[i];

                    sAttributeTypes[i * 3] = "string";
                    sAttributeTypes[i*3 + 1] = "long";
                    sAttributeTypes[i*3 + 2] = "double";
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
                                  string dTargetAl2O3, string dActualAl2O3, string dTargetCaO, string dActualCaO,string dTargetF, string dActualF,
                                  string dTargetFe2O3, string dActualFe2O3, string dTargetK2O, string dActualK2O,string dTargetMgO, string dActualMgO,
                                  string dTargetMnO, string dActualMnO, string dTargetNa2O3, string dActualNa2O3,string dTargetSiO2, string dActualSiO2,
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


                return client2.createpart(sBatchNo, sBatchName, sProductName, sBatchType, sFolder, sFullName,sAttributeNames, sAttributeValues, sAttributeTypes, sCheckInComments, iiProdOrLibrary, Convert.ToInt16(sWebAppId));
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

                return client2.createpart(sBatchNo, sBatchName, sProductName, sBatchType, sFolder, sFullName,sAttributeNames, sAttributeValues, sAttributeTypes, sCheckInComments, iiProdOrLibrary, Convert.ToInt16(sWebAppId));
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
                                  string dTargetAl2O3, string dActualAl2O3, string dTargetCaO, string dActualCaO,string dTargetF, string dActualF,
                                  string dTargetFe2O3, string dActualFe2O3, string dTargetK2O, string dActualK2O,string dTargetMgO, string dActualMgO,
                                  string dTargetMnO, string dActualMnO, string dTargetNa2O3, string dActualNa2O3,string dTargetSiO2, string dActualSiO2,
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
                                           string sPlantCode, string sTechActionDesc, string sComments, string iProdOrLibrary, string sWebAppId)
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

                return client2.createproblemreport2(sTechActionNo, sTechActionName, sTechActionDesc, sProductName, sPRType, sFolderNameAndPath, sAttributeNames, sAttributeValues, sAttributeTypes, iiProdOrLibrary, "", Convert.ToInt16(sWebAppId));
            }
        }

        public string UpdateTechnicalAction(string sSessionId, string sUserId, string sFullName, string sTechActionNo, string sTechActionName, string sTechActionDesc, 
                                            string sComments, string sWebAppId)
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

                return client2.setproblemreportattributes(sTechActionNo, sTechActionName, sAttributeNames, sAttributeValues, sAttributeTypes, "", Convert.ToInt16(sWebAppId));
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

                sVariableNames[0] = "gdtElapsedNextDateLocal"; //THis is the variable in the task. A local variable
                sVariableValues[0] = sNextElapsedDate;
                sVariableTypes[0] = "date";

                return client2.completetask(Convert.ToInt32(sWorkItemId), Convert.ToInt32(sAssignedActivityId), sRoute, sVariableNames, sVariableTypes, sVariableValues, Convert.ToInt16(sWebAppId));
            }
        }

        public string SetTaskOperationalHoursOnCompletion(string sSessionId, string sUserId, string sWorkItemId, string sAssignedActivityId, string sHoursOnCompletion, string sWebAppId)
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

                return client2.completetask(Convert.ToInt32(sWorkItemId), Convert.ToInt32(sAssignedActivityId), "", sVariableNames, sVariableTypes, sVariableValues, Convert.ToInt16(sWebAppId));
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

                for(i=0; i< sAffectedParts2.Length; i++)
                {
                    iAffectedPartTypes[i] = Convert.ToInt16(sAffectedParts2[i]);
                }
                return client2.setpraffectedobjects(sProdLossNo, sAffectedParts, iAffectedPartTypes, Convert.ToInt16(sWebAppId));
            }
        }

        public string SetProbRptState(string sSessionId, string sUserId, string sFullName, string sProbRptNo, string sProbRptName, string sLifecycleState, string sComments,  string sWebAppId)
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
                if(sRtn == "Success")
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

        public string ReviseDocumentAndRemoveAttachments(string sSessionId, string sUserId, string sFullname, string sDocNo, string sDocName,  string sRevision, 
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
                long llLineNumber = Convert.ToInt64(lLineNumber);
                ExampleService.MyJavaService3Client client2 = GetWCService();
                return client2.deletepartpartlinkbylinenumber(sFullName, llLineNumber, sParentPartNo, sChildPartNumber, sCheckInComments, Convert.ToInt16(sWebAppId));
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

                if(sRtn1.StartsWith("Success"))
                {

                    //Get the new document number
                    string[] sSuccess = Extract_Values(sRtn1);

                    sDocNo = sSuccess[1];

                    if (sBatchNo != "")
                    {
                        if(sComments != "")
                        {
                            if(sTargetQty != "")
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

        public string CreateCableScheduleItem(string sSessionId, string sUserId, string sCSNo, string sProductName,  string sFolderNameAndPath,
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

                sAttributeNames[0] = "Originator";
                sAttributeNames[1] = "Comments";

                sAttributeValues[0] = sOriginator;
                sAttributeValues[1] = sCableComments;

                sAttributeTypes[0] = "string";
                sAttributeTypes[1] = "string";

                sAttributeNames2[0] = "ToOrFrom";
                sAttributeValues2[0] = "0"; //0 = from, 1 = to
                sAttributeTypes2[0] = "string";

                string sRtn = "Success";

                string sRtn1 = client2.createpart(sCableNo, sCableName, sProductName, "local.rs.vsrs05.Regain.CablePart", sFolderNameAndPath, sFullName,
                                                    sAttributeNames, sAttributeValues, sAttributeTypes, sCableCheckInComments, iiProdOrLibrary, Convert.ToInt16(sWebAppId));
                if (!sRtn1.StartsWith("Success"))
                {
                    return sRtn1;
                }                    

                if (!sMaterialCableCode.Equals(""))
                {
                    //Now create a parent child link to the cable material item
                    string sRtn3 = client2.setpartpartlink(sUserId, sCableNo, sMaterialCableCode, Convert.ToDouble(dLength), "Creating link to cable material item " + sMaterialCableCode,
                                                               "wt.part.WTPartUsageLink", "m", Convert.ToInt16(sWebAppId));
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

                string sRtn1 = "Success";
                if (!sOldMaterialCode.Equals(""))
                {
                    sRtn1 = DeletePartToPartLink(sSessionId, sUserId, sFullName, sCableNo, sOldMaterialCode, sCableCheckInComments, sWebAppId);
                }

                if (sRtn1.StartsWith("Success"))
                {
                    //Now create a parent child link to the cable material item
                    string sRtn3 = client2.setpartpartlink(sUserId, sCableNo, sMaterialCode, Convert.ToDouble(sLength), "Creating link to cable material item " + sMaterialCode,
                                                           "wt.part.WTPartUsageLink", "m", Convert.ToInt16(sWebAppId));
                    sRtn1 = sRtn3; //This happens regardless of success or not                
                }

                return sRtn1;
            }
        }


        public string CreateCablePartLink(string sSessionId, string sUserId, string sCableNo, string sFuncLoc, string sLineNumber,string sToOrFrom, string sCheckInComments, string sWebAppId)
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

        public string CreateMaterialCatalogItem(string sSessionId, string sUserId, string sFullName, string sMatCatNo, string sMatCatType, string sName, string sDesc, string sLongDesc, 
                                                string sDrivekW, string sFullLoadCurrent, string sCheckInComments, string sWebAppId)
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
                    sAttributeNames[sAttributeNames.Length-1] = "DrivekW";
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


                sReturn = client2.createpart("", sName, "Regain Material Catalogue", "local.rs.vsrs05.Regain.AutoNumberedPart", sFolder, sFullName,  sAttributeNames, sAttributeValues, sAttributeTypes, sCheckInComments, 1, Convert.ToInt16(sWebAppId));
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

        public string UpdateMaterialCatalogItem(string sSessionId, string sUserId, string sFullName, string sMatCatNo, string sMatCatNewType, string sMatCatOldType, string sName, string sDesc, string sLongDesc, string sDrivekW, string sFullLoadCurrent, string sCheckInComments, string sWebAppId, string sNewLink)
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
                                           string sPowerCable, string sControlCable, string sInstrumentationCable, string sDataCable,
                                           string sInstRegFlag, string sFullLoadCurrent,
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
                                           string sPowerCable, string sControlCable, string sInstrumentationCable, string sDataCable,
                                           string sInstRegFlag, string sFullLoadCurrent,
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
            catch(Exception ex)
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
    }
}
