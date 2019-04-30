using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.ServiceModel.Web;
using System.Text;

namespace WcfWCService
{
    // NOTE: You can use the "Rename" command on the "Refactor" menu to change the interface name "IService1" in both code and config file together.
    [ServiceContract(Namespace = "http://regain.com/rest")]
    public interface IService1
    {

        [OperationContract]
        [WebGet(UriTemplate = "add/{sSessionId}/{sUserId}/{a}/{b}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string add(string sSessionId, string sUserId, string a, string b, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "simpleadd/{a}/{b}", ResponseFormat = WebMessageFormat.Xml)]
        string simpleadd(string a, string b);

        [OperationContract]
        [WebGet(UriTemplate = "cookielogin/{sUsername}/{sPassword}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string CookieLogin(string sUsername, string sPassword, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "createwcdoc/{sSessionId}/{sUserId}/{sDocNo}/{sDocName}/{sProductName}/{sDocType}/{sFolderNameAndPath}/{sLongDesc}/{sOriginator}/{sOriginatorDocId}/{sJobCode}/{sRevision}/{sCheckInComments}/{iProdOrLibrary}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string CreateWCDoc(string sSessionId, string sUserId, string sDocNo, string sDocName, string sProductName, string sDocType, string sFolderNameAndPath,
                           string sLongDesc, string sOriginator, string sOriginatorDocId, string sJobCode, string sRevision, string sCheckInComments, string iProdOrLibrary, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "createworkexecutionpackage/{sSessionId}/{sUserId}/{sWorkItemId}/{sAssignedActivityId}/{sRoute}/{sPlannedWorkPackageNo}/{sWEDName}/{sProductName}/{sDocType}/{sFolderNameAndPath}/{sOriginator}/{sJobCode}/{sNew}/{sExistingWEDNo}/{sWebAppId}/{sSkipCompleteTask}", ResponseFormat = WebMessageFormat.Xml)]
        string CreateWorkExecutionPackage(string sSessionId, string sUserId, string sWorkItemId, string sAssignedActivityId, string sRoute, 
                                          string sPlannedWorkPackageNo, string sWEDName, string sProductName, string sDocType, string sFolderNameAndPath,
                                          string sOriginator, string sJobCode, string sNew, string sExistingWEDNo, string sWebAppId, string sSkipCompleteTask);

        [OperationContract]
        [WebGet(UriTemplate = "setdocpartdescribedbylink/{sSessionId}/{sUserId}/{sDocNo}/{sPartNo}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string SetDocPartDescribedByLink(string sSessionId, string sUserId, string sDocNo, string sPartNo, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "createprojectworkitem/{sSessionId}/{sUserId}/{sFullName}/{sParentPartNo}/{sPartNo}/{sPartName}/{sProductName}/{sPartType}/{sPartUsageType}/{sPartUsageUnit}/{sFolderNameAndPath}/{sCheckInComments}/{sLineNumber}/{iProdOrLibrary}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string CreateProjectWorkItem(string sSessionId, string sUserId, string sFullName, string sParentPartNo, string sPartNo, string sPartName,
                                            string sProductName, string sPartType, string sPartUsageType, string sPartUsageUnit, string sFolderNameAndPath,
                                            string sCheckInComments,string sLineNumber, string iProdOrLibrary, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "insertexistingprojectworkitem/{sSessionId}/{sUserId}/{sFullName}/{sParentPartNo}/{sExistingPWIPartNo}/{sPartUsageType}/{sPartUsageUnit}/{sCheckInComments}/{sLineNumber}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string InsertExistingProjectWorkItem(string sSessionId, string sUserId, string sFullName, string sParentPartNo, string sExistingPWIPartNo,
                                                    string sPartUsageType, string sPartUsageUnit, string sCheckInComments, string sLineNumber, string sWebAppId);
        
        [OperationContract]
        [WebGet(UriTemplate = "createfronesisproject/{sSessionId}/{sUserId}/{sProjNo}/{sProjDesc}/{sProductName}/{sDocType}/{sPartType}/{sFolderNameAndPath}/{sClientDesc}/{sOriginator}/{sClientProjNo}/{sRevision}/{sCheckInComments}/{iProdOrLibrary}/{sWebAppId}/{sProjType}", ResponseFormat = WebMessageFormat.Xml)]
        string CreateFronesisProject(string sSessionId, string sUserId, string sProjNo, string sProjDesc, string sProductName, string sDocType, string sPartType, string sFolderNameAndPath,
                                  string sClientDesc, string sOriginator, string sClientProjNo, string sRevision, string sCheckInComments, string iProdOrLibrary, string sWebAppId, string sProjType);

        [OperationContract]
        [WebGet(UriTemplate = "createfronesisprojectchilddoc/{sSessionId}/{sUserId}/{sProjNo}/{sChildDocNo}/{sChildDocName}/{sProductName}/{sDocType}/{sFolderNameAndPath}/{sOriginator}/{sRevision}/{sCheckInComments}/{iProdOrLibrary}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string CreateFronesisProjectChildDoc(string sSessionId, string sUserId, string sProjNo, string sChildDocNo, string sChildDocName, string sProductName, string sDocType, string sFolderNameAndPath,
                                             string sOriginator, string sRevision, string sCheckInComments, string iProdOrLibrary, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "attachwcdoc/{sSessionId}/{sUserId}/{sFullname}/{sDocNo}/{sAttachDesc}/{sAttachPath}/{bSecondary}/{sAttachComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string AttachWCDoc(string sSessionId, string sUserId, string sFullName, string sDocNo, string sAttachDesc, string sAttachPath, string bSecondary, string sAttachComments, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "deletewcdoc/{sSessionId}/{sUserId}/{sFullname}/{sDocNo}/{sAttachFile}/{bSecondary}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string DeleteWCDoc(string sSessionId, string sUserId, string sFullName, string sDocNo, string sAttachFile, string bSecondary, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "setdocattributestring/{sSessionId}/{sUserId}/{sDocNo}/{sDocName}/{sLongDesc}/{sOriginator}/{sOriginatorDocId}/{sJobCode}/{sCheckInComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string SetDocAttributeStrings(string sSessionId, string sUserId, string sDocNo, string sDocName, string sLongDesc, string sOriginator, string sOriginatorDocId, string sJobCode,
                                      string sCheckInComments, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "setdoctodocref/{sSessionId}/{sUserId}/{sFullname}/{sDocNo}/{sReferencedDocNo}/{sCheckinComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string SetDocToDocRef(string sSessionId, string sUserId, string sFullName, string sDocNo, string sReferencedDocNo, string sCheckinComments, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "setdoctodocrefs/{sSessionId}/{sUserId}/{sFullname}/{sDocNo}/{sReferencedDocNos}/{sCheckinComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string SetDocToDocRefs(string sSessionId, string sUserId, string sFullName, string sDocNo, string sReferencedDocNos, string sCheckinComments, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "deletedoctodocref/{sSessionId}/{sUserId}/{sFullname}/{sDocNo}/{sReferencedDocNo}/{sCheckinComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string DeleteDocToDocRef(string sSessionId, string sUserId, string sFullName, string sDocNo, string sReferencedDocNo, string sCheckinComments, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "deletedoctodocrefs/{sSessionId}/{sUserId}/{sFullname}/{sDocNo}/{sReferencedDocNos}/{sCheckinComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string DeleteDocToDocRefs(string sSessionId, string sUserId, string sFullName, string sDocNo, string sReferencedDocNos, string sCheckinComments, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "setdoctopartref/{sSessionId}/{sUserId}/{sFullname}/{sDocNo}/{sPartNo}/{sCheckinComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string SetDocToPartRef(string sSessionId, string sUserId, string sFullName, string sDocNo, string sPartNo, string sCheckinComments, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "setdoctopartrefs/{sSessionId}/{sUserId}/{sFullname}/{sDocNo}/{sPartNos}/{sCheckinComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string SetDocToPartRefs(string sSessionId, string sUserId, string sFullName, string sDocNo, string sPartNos, string sCheckinComments, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "deletedoctopartref/{sSessionId}/{sUserId}/{sFullname}/{sDocNo}/{sPartNo}/{sCheckinComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string DeleteDocToPartRef(string sSessionId, string sUserId, string sFullName, string sDocNo, string sPartNo, string sCheckinComments, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "deletedoctopartrefs/{sSessionId}/{sUserId}/{sFullname}/{sDocNo}/{sPartNos}/{sCheckinComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string DeleteDocToPartRefs(string sSessionId, string sUserId, string sFullName, string sDocNo, string sPartNos, string sCheckinComments, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "deletedoctopartdescribeby/{sSessionId}/{sUserId}/{sFullname}/{sDocNo}/{sPartNo}/{sCheckinComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string DeleteDocToPartDescribeBy(string sSessionId, string sUserId, string sFullName, string sDocNo, string sPartNo, string sCheckinComments, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "deletedoctopartdescribebys/{sSessionId}/{sUserId}/{sFullname}/{sDocNo}/{sPartNos}/{sCheckinComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string DeleteDocToPartDescribeBys(string sSessionId, string sUserId, string sFullName, string sDocNo, string sPartNos, string sCheckinComments, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "updateactionrequest/{sSessionId}/{sUserId}/{sFullname}/{sARCode}/{sARName}/{sARCategory}/{sARCause}/{sARComments}/{sARLongDesc}/{sARDate}/{sRequestActionType}/{sCheckInComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string UpdateActionRequest(string sSessionId, string sUserId, string sFullName, string sARCode, string sARName, string sARCategory, string sARCause, string sARComments,
                                   string sARLongDesc, string sARDate, string sRequestActionType, string sCheckinComments, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "setparttopartlink/{sSessionId}/{sUserId}/{sFullname}/{sParentPartNo}/{sChildPartNumber}/{dQty}/{sCheckInComments}/{sPartUsageType}/{sUnit}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string SetPartToPartLink(string sSessionId, string sUserId, string sFullName, string sParentPartNo, string sChildPartNumber, string dQty,
                                 string sCheckInComments, string sPartUsageType, string sUnit, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "setpartusagelinkqty/{sSessionId}/{sUserId}/{sParentPartNo}/{sChildPartNo}/{dQty}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string SetPartUsageLinkQty(string sSessionId, string sUserId, string sParentPartNo, string sChildPartNo, string dQty, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "deleteparttopartlink/{sSessionId}/{sUserId}/{sFullname}/{sParentPartNo}/{sChildPartNumber}/{sCheckInComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string DeletePartToPartLink(string sSessionId, string sUserId, string sFullName, string sParentPartNo, string sChildPartNumber, string sCheckInComments, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "createactionrequest/{sSessionId}/{sUserId}/{sFullname}/{sProductName}/{sFolder}/{sARName}/{sARCategory}/{sARCause}/{sARComments}/{sARLongDesc}/{sARDate}/{sRequestActionType}/{sCheckInComments}/{iProdOrLibrary}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string CreateActionRequest(string sSessionId, string sUserId, string sFullName, string sProductName, string sFolder, string sARName, string sARCategory, string sARCause, string sARComments,
                                          string sARLongDesc, string sARDate, string sRequestActionType, string sCheckInComments, string iProdOrLibrary, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "createproductionloss/{sSessionId}/{sUserId}/{sFullname}/{sProdLossNo}/{sProdLossName}/{sProductName}/{sPRType}/{sFolderNameAndPath}/{sPlant}/{sRegainCategory}/{sRegainSubCategory}/{sStartDateAndTime}/{sEndDateAndTime}/{dDurationInHours}/{sSuspectedFailureReason}/{sComments}/{iProdOrLibrary}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string CreateProductionLoss(string sSessionId, string sUserId, string sFullName, string sProdLossNo, string sProdLossName, string sProductName, string sPRType, string sFolderNameAndPath,
                                    string sPlant, string sRegainCategory, string sRegainSubCategory, string sStartDateAndTime, string sEndDateAndTime,
                                    string dDurationInHours, string sSuspectedFailureReason, string sComments, string iProdOrLibrary, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "updateproductionloss/{sSessionId}/{sUserId}/{sFullname}/{sProdLossNo}/{sProdLossName}/{sPlant}/{sRegainCategory}/{sRegainSubCategory}/{sStartDateAndTime}/{sEndDateAndTime}/{dDurationInHours}/{sSuspectedFailureReason}/{sComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string UpdateProductionLoss(string sSessionId, string sUserId, string sFullName, string sProdLossNo, string sProdLossName, string sPlant, string sRegainCategory, string sRegainSubCategory,
                                    string sStartDateAndTime, string sEndDateAndTime, string dDurationInHours, string sSuspectedFailureReason, string sComments, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "createissuereport/{sSessionId}/{sUserId}/{sFullname}/{sIssueRptNo}/{sIssueRptName}/{sPlant}/{sProductName}/{sPRType}/{sFolderNameAndPath}/{sComments}/{iProdOrLibrary}/{sNeedDate}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string CreateIssueReport(string sSessionId, string sUserId, string sFullName, string sIssueRptNo, string sIssueRptName, string sPlant, string sProductName,
                                 string sPRType, string sFolderNameAndPath, string sComments, string iProdOrLibrary, string sNeedDate, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "createImprovementreport/{sSessionId}/{sUserId}/{sFullname}/{sImprovementRptNo}/{sImprovementRptName}/{sPlant}/{sProductName}/{sPRType}/{sFolderNameAndPath}/{sComments}/{iProdOrLibrary}/{sNeedDate}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string CreateImprovementReport(string sSessionId, string sUserId, string sFullName, string sImprovementRptNo, string sImprovementRptName, string sPlant,
                                       string sProductName, string sPRType, string sFolderNameAndPath, string sComments, string iProdOrLibrary, string sNeedDate, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "createbatchevent/{sSessionId}/{sUserId}/{sFullname}/{sBatchEventNo}/{sBatchEventName}/{sProductName}/{sPRType}/{sFolderNameAndPath}/{sComments}/{iProdOrLibrary}/{sTransDate}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string CreateBatchEvent(string sSessionId, string sUserId, string sFullName, string sBatchEventNo, string sBatchEventName, string sProductName, string sPRType,
                                string sFolderNameAndPath, string sComments, string iProdOrLibrary, string sTransDate, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "updateissuereport/{sSessionId}/{sUserId}/{sFullname}/{sIssueRptNo}/{sIssueRptName}/{sPlant}/{sComments}/{sTransDate}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string UpdateIssueReport(string sSessionId, string sUserId, string sFullName, string sIssueRptNo, string sIssueRptName, string sPlant, string sComments, string sTransDate, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "updateImprovementreport/{sSessionId}/{sUserId}/{sFullname}/{sImprovementRptNo}/{sImprovementRptName}/{sPlant}/{sComments}/{sTransDate}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string UpdateImprovementReport(string sSessionId, string sUserId, string sFullName, string sImprovementRptNo, string sImprovementRptName, string sPlant, string sComments, string sTransDate, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "updatebatchevent/{sSessionId}/{sUserId}/{sFullname}/{sBatchEventNo}/{sBatchEventName}/{sComments}/{sTransDate}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string UpdateBatchEvent(string sSessionId, string sUserId, string sFullName, string sBatchEventNo, string sBatchEventName, string sComments, string sTransDate, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "settaskoperationalhoursoncompletion/{sSessionId}/{sUserId}/{sWorkItemId}/{sAssignedActivityId}/{sHoursOnCompletion}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string SetTaskOperationalHoursOnCompletion(string sSessionId, string sUserId, string sWorkItemId, string sAssignedActivityId, string sHoursOnCompletion, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "settaskwocompletiondate/{sSessionId}/{sUserId}/{sWorkItemId}/{sAssignedActivityId}/{sRoute}/{sDateOnCompletion}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string SetTaskWOCompletionDate(string sSessionId, string sUserId, string sWorkItemId, string sAssignedActivityId, string sRoute, string sDateOnCompletion, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "progresstask/{sSessionId}/{sUserId}/{sWorkItemId}/{sAssignedActivityId}/{sRoute}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string ProgressTask(string sSessionId, string sUserId, string sWorkItemId, string sAssignedActivityId, string sRoute,  string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "setproductionlossaffectedobjects/{sSessionId}/{sUserId}/{sProdLossNo}/{sAffectdPartsString}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string SetProductionLossAffectedObjects(string sSessionId, string sUserId, string sProdLossNo, string sAffectdPartsString, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "deleteproductionlossaffectedobjects/{sSessionId}/{sUserId}/{sProdLossNo}/{sAffectdPartsString}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string DeleteProductionLossAffectedObjects(string sSessionId, string sUserId, string sProdLossNo, string sAffectdPartsString, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "attachproductionlossdoc/{sSessionId}/{sUserId}/{sProdLossNo}/{sAttachDesc}/{sAttachPath}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string AttachProductionLossDoc(string sSessionId, string sUserId, string sProdLossNo, string sAttachDesc, string sAttachPath, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "deleteproductionlossattachment/{sSessionId}/{sUserId}/{sProdLossNo}/{sAttachFileName}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string DeleteProductionLossAttachment(string sSessionId, string sUserId, string sProdLossNo, string sAttachFileName, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "deleteproblemreport/{sSessionId}/{sUserId}/{sProbReportNo}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string DeleteProblemReport(string sSessionId, string sUserId, string sProbReportNo, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "revisedocument/{sSessionId}/{sUserId}/{sDocNo}/{sRevision}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string ReviseDocument(string sSessionId, string sUserId, string sDocNo, string sRevision, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "updatedocattributes/{sSessionId}/{sUserId}/{sDocNumber}/{sDocName}/{sAttributeName1}/{sAttributeValue1}/{sAttributeType1}/{sAttributeName2}/{sAttributeValue2}/{sAttributeType2}/{sAttributeName3}/{sAttributeValue3}/{sAttributeType3}/{sAttributeName4}/{sAttributeValue4}/{sAttributeType4}/{sAttributeName5}/{sAttributeValue5}/{sAttributeType5}/{sCheckInComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string UpdateDocAttributes(string sSessionId, string sUserId, string sDocNumber, string sDocName,
                                   string sAttributeName1, string sAttributeValue1, string sAttributeType1,
                                   string sAttributeName2, string sAttributeValue2, string sAttributeType2,
                                   string sAttributeName3, string sAttributeValue3, string sAttributeType3,
                                   string sAttributeName4, string sAttributeValue4, string sAttributeType4,
                                   string sAttributeName5, string sAttributeValue5, string sAttributeType5,
                                   string sCheckinComments, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "updatepartattributes/{sSessionId}/{sUserId}/{sPartNumber}/{sPartName}/{sAttributeName1}/{sAttributeValue1}/{sAttributeType1}/{sAttributeName2}/{sAttributeValue2}/{sAttributeType2}/{sAttributeName3}/{sAttributeValue3}/{sAttributeType3}/{sAttributeName4}/{sAttributeValue4}/{sAttributeType4}/{sAttributeName5}/{sAttributeValue5}/{sAttributeType5}/{sCheckInComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string UpdatePartAttributes(string sSessionId, string sUserId, string sPartNumber, string sPartName,
                                   string sAttributeName1, string sAttributeValue1, string sAttributeType1,
                                   string sAttributeName2, string sAttributeValue2, string sAttributeType2,
                                   string sAttributeName3, string sAttributeValue3, string sAttributeType3,
                                   string sAttributeName4, string sAttributeValue4, string sAttributeType4,
                                   string sAttributeName5, string sAttributeValue5, string sAttributeType5,
                                   string sCheckinComments,string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "updateoperatinghours/{sSessionId}/{sUserId}/{sPartNumber}/{sOriginatorName}/{sOperatingHours}/{sCheckinComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string UpdateOperatingHours(string sSessionId, string sUserId, string sPartNumber,
                                   string sOriginatorName, string sOperatingHours,
                                   string sCheckinComments, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "creatembapartusagelink/{sSessionId}/{sUserId}/{sFullname}/{sParentPartNo}/{sChildPartNo}/{dQty}/{lLineNumber}/{sCheckInComments}/{sDispatchDocketNo}/{sTransactionDate}/{sComments}/{sProdOrderNo}/{dMoisturePercentage}/{sContainerId}/{sInvoiceStatus}/{sBatchNo}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string CreateMBAPartUsageLink(string sSessionId, string sUserId, string sFullName, string sParentPartNo, string sChildPartNo, string dQty, string lLineNumber, string sCheckInComments, 
                                      string sDispatchDocketNo, string sTransactionDate, string sComments, string sProdOrderNo, string dMoisturePercentage, string sContainerId,
                                      string sInvoiceStatus, string sBatchNo, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "updatembapartusagelinkfromdd/{sSessionId}/{sUserId}/{sFullname}/{sParentPartNo}/{sChildPartNo}/{dQty}/{lOldLineNumber}/{lNewLineNumber}/{sCheckInComments}/{sDispatchDocketNo}/{sTransactionDate}/{sComments}/{sProdOrderNo}/{sContainerId}/{dMoisturePercentage}/{sInvoiceStatus}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string UpdateMBAPartUsageLinkFromDD(string sSessionId, string sUserId, string sFullName, string sParentPartNo, string sChildPartNo, string dQty, 
                                            string lOldLineNumber, string lNewLineNumber, string sCheckInComments, string sDispatchDocketNo, string sTransactionDate,
                                            string sComments, string sProdOrderNo, string sContainerId, string dMoisturePercentage, string sInvoiceStatus, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "updatembapartusagelinkfrompo/{sSessionId}/{sUserId}/{sFullname}/{sParentPartNo}/{sChildPartNo}/{dQty}/{lOldLineNumber}/{lNewLineNumber}/{sCheckInComments}/{sDispatchDocketNo}/{sTransactionDate}/{sComments}/{sProdOrderNo}/{dMoisturePercentage}/{sInvoiceStatus}/{sBatchNo}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string UpdateMBAPartUsageLinkFromPO(string sSessionId, string sUserId, string sFullName, string sParentPartNo, string sChildPartNo, string dQty, 
                                            string lOldLineNumber, string lNewLineNumber, string sCheckInComments, string sDispatchDocketNo, string sTransactionDate,
                                            string sComments, string sProdOrderNo, string dMoisturePercentage, string sInvoiceStatus, string sBatchNo, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "updatembatransactioninvoicestatus/{sSessionId}/{sUserId}/{sFullname}/{sParentPartNo}/{sChildPartNo}/{sLineNumber}/{sInvoiceStatus}/{sInvoiceNo}/{sBatchList}/{sCutoffDate}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string UpdateMBATransactionInvoiceStatus(string sSessionId, string sUserId, string sFullName, string sParentPartNo, string sChildPartNo, string sLineNumber,
                                                 string sInvoiceStatus, string sInvoiceNo, string sBatchList, string sCutoffDate, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "updatembamultipletransactioninvoicestatus/{sSessionId}/{sUserId}/{sFullname}/{sParentPartNo}/{sChildPartNo}/{sLineNumber}/{sInvoiceStatus}/{sQtyInvoiced}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string UpdateMBAMultipleTransactionInvoiceStatus(string sSessionId, string sUserId, string sFullName, string sParentPartNo, string sChildPartNo, string sLineNumber,
                                                         string sInvoiceStatus, string sQtyInvoiced, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "deleteparttoopartlinkbydispatchdocket/{sSessionId}/{sUserId}/{sFullname}/{sDispatchDocketNo}/{lLineNumber}/{sParentPartNo}/{sChildPartNo}/{sCheckInComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string DeletePartToPartLinkByDispatchDocket(string sSessionId, string sUserId, string sFullName, string sDispatchDocketNo, string lLineNumber, string sParentPartNo,
                                                    string sChildPartNo, string sCheckInComments, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "deleteparttoopartlinkbyProductionOrder/{sSessionId}/{sUserId}/{sFullname}/{sProductionOrderNo}/{lLineNumber}/{sParentPartNo}/{sChildPartNo}/{sCheckInComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string DeletePartToPartLinkByProductionOrder(string sSessionId, string sUserId, string sFullName, string sProductionOrderNo, string lLineNumber,
                                                     string sParentPartNo, string sChildPartNo, string sCheckInComments, string sWebAppId);


        [OperationContract]
        [WebGet(UriTemplate = "deleteparttoopartlinkbylinenumber/{sSessionId}/{sUserId}/{sFullname}/{lLineNumber}/{sParentPartNo}/{sChildPartNo}/{sCheckInComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string DeletePartToPartLinkByLineNumber(string sSessionId, string sUserId, string sFullName, string lLineNumber,
                                                string sParentPartNo, string sChildPartNo, string sCheckInComments, string sWebAppId);


        [OperationContract]
        [WebGet(UriTemplate = "createbatch/{sSessionId}/{sUserId}/{sFullname}/{sBatchNo}/{sBatchName}/{sProductName}/{sFolder}/{sBatchType}/{sCheckInComments}/{iProdOrLibrary}/" +
                              "{dTargetQty}/{dActualQty}/{dMoisturePercentage}/{sQualityStatus}/" +
                              "{dTargetAl2O3}/{dActualAl2O3}/{dTargetCaO}/{dActualCaO}/{dTargetF}/{dActualF}/" +
                              "{dTargetFe2O3}/{dActualFe2O3}/{dTargetK2O}/{dActualK2O}/{dTargetMgO}/{dActualMgO}/" +
                              "{dTargetMnO}/{dActualMnO}/{dTargetNa2O3}/{dActualNa2O3}/{dTargetSiO2}/{dActualSiO2}/" +
                              "{dTargetC}/{dActualC}/{dTargetSO3}/{dActualSO3}/{dTargetCN}/{dActualCN}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string CreateBatch(string sSessionId, string sUserId, string sFullName, string sBatchNo, string sBatchName, string sProductName, string sFolder, string sBatchType, string sCheckInComments, string iProdOrLibrary,
                          string dTargetQty, string dActualQty, string dMoisturePercentage, string sQualityStatus,
                          string dTargetAl2O3, string dActualAl2O3, string dTargetCaO, string dActualCaO, string dTargetF, string dActualF,
                          string dTargetFe2O3, string dActualFe2O3, string dTargetK2O, string dActualK2O, string dTargetMgO, string dActualMgO,
                          string dTargetMnO, string dActualMnO, string dTargetNa2O3, string dActualNa2O3, string dTargetSiO2, string dActualSiO2,
                          string dTargetC, string dActualC, string dTargetSO3, string dActualSO3, string dTargetCN, string dActualCN, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "createmba/{sSessionId}/{sUserId}/{sFullname}/{sBatchNo}/{sBatchName}/{sProductName}/{sFolder}/" +
                              "{sBatchType}/{sCheckInComments}/{iProdOrLibrary}/{dMoisturePercentage}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string CreateMBA(string sSessionId, string sUserId, string sFullName, string sBatchNo, string sBatchName, string sProductName, string sFolder,
                         string sBatchType, string sCheckInComments, string iProdOrLibrary, string dMoisturePercentage, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "copypart/{sSessionId}/{sUserId}/{sSourcePartNo}/{sTargetPartNo}/{sTargetPartName}/{sProductName}/{sFolder}/" +
                              "{sPartType}/{iProdOrLibrary}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string CopyPart(string sSessionId, string sUserId, string sSourcePartNo, string sTargetPartNo, string sTargetPartName, string sProductName,
                        string sFolder, string sPartType, string iProdOrLibrary, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "updatebatch/{sSessionId}/{sUserId}/{sFullname}/{sBatchNo}/{sBatchName}/{sCheckinComments}/{dTargetQty}/{dActualQty}/{dMoisturePercentage}/{sQualityStatus}/" +
                              "{dTargetAl2O3}/{dActualAl2O3}/{dTargetCaO}/{dActualCaO}/{dTargetF}/{dActualF}/" +
                              "{dTargetFe2O3}/{dActualFe2O3}/{dTargetK2O}/{dActualK2O}/{dTargetMgO}/{dActualMgO}/" +
                              "{dTargetMnO}/{dActualMnO}/{dTargetNa2O3}/{dActualNa2O3}/{dTargetSiO2}/{dActualSiO2}/" +
                              "{dTargetC}/{dActualC}/{dTargetSO3}/{dActualSO3}/{dTargetCN}/{dActualCN}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string UpdateBatch(string sSessionId, string sUserId, string sFullName, string sBatchNo, string sBatchName, string sCheckinComments,
                                  string dTargetQty, string dActualQty, string dMoisturePercentage, string sQualityStatus,
                                  string dTargetAl2O3, string dActualAl2O3, string dTargetCaO, string dActualCaO, string dTargetF, string dActualF,
                                  string dTargetFe2O3, string dActualFe2O3, string dTargetK2O, string dActualK2O, string dTargetMgO, string dActualMgO,
                                  string dTargetMnO, string dActualMnO, string dTargetNa2O3, string dActualNa2O3, string dTargetSiO2, string dActualSiO2,
                                  string dTargetC, string dActualC, string dTargetSO3, string dActualSO3, string dTargetCN, string dActualCN, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "updatemba/{sSessionId}/{sUserId}/{sFullname}/{sBatchNo}/{sBatchName}/{sCheckinComments}/{dMoisturePercentage}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string UpdateMBA(string sSessionId, string sUserId, string sFullName, string sBatchNo, string sBatchName, string sCheckinComments, string dMoisturePercentage, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "createprodorder/{sSessionId}/{sUserId}/{sDocNo}/{sDocName}/{sProductName}/{sDocType}/{sFolderNameAndPath}/" +
                              "{sBatchNo}/{sTargetQty}/{sProdNoDate}/{sOriginator}/{sJobCode}/{sComments}/{sRevision}/{sCheckInComments}/" +
                              "{iProdOrLibrary}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string CreateProdOrder(string sSessionId, string sUserId, string sDocNo, string sDocName, string sProductName, string sDocType, string sFolderNameAndPath,
                                      string sBatchNo, string sTargetQty, string sProdNoDate, string sOriginator, string sJobCode, string sComments,
                                      string sRevision, string sCheckInComments, string iProdOrLibrary, string sWebAppId);


        [OperationContract]
        [WebGet(UriTemplate = "creatematerialcatalogitem/{sSessionId}/{sUserId}/{sFullname}/{sMatCatNo}/{sMatCatType}/{sName}/{sDesc}/{sLongDesc}/{sCheckInComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string CreateMaterialCatalogItem(string sSessionId, string sUserId, string sFullName, string sMatCatNo, string sMatCatType, string sName, string sDesc, string sLongDesc, string sCheckInComments, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "updatematerialcatalogitem/{sSessionId}/{sUserId}/{sFullname}/{sMatCatNo}/{sMatCatNewType}/{sMatCatOldType}/{sName}/{sDesc}/{sLongDesc}/{sCheckInComments}/{sWebAppId}/{sNewLink}", ResponseFormat = WebMessageFormat.Xml)]
        string UpdateMaterialCatalogItem(string sSessionId, string sUserId, string sFullName, string sMatCatNo, string sMatCatNewType, string sMatCatOldType, string sName, string sDesc, string sLongDesc, string sCheckInComments, string sWebAppId, string sNewLink);


        [OperationContract]
        [WebGet(UriTemplate = "createplantequipitem/{sSessionId}/{sUserId}/{sFullname}/{sPlantEquipNo}/{sPlantEquipType}/{sName}/{sDesc}/{sLongDesc}/{sContSysType}/{sDriveRating}/{sEquipRegFlag}/{sIPRegFlag}/{sIPAddress}/{sProduct}/{sFolder}/{sCheckInComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string CreatePlantEquipItem(string sSessionId, string sUserId, string sFullName, string sPlantEquipNo,
                                           string sPlantEquipType, string sName, string sDesc, string sLongDesc,
                                           string sContSysType, string sDriveRating, string sEquipRegFlag, string sIPRegFlag, string sIPAddress,
                                           string sProduct, string sFolder, string sCheckInComments, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "updateplantequipitem/{sSessionId}/{sUserId}/{sFullname}/{sPlantEquipNo}/{sName}/{sDesc}/{sLongDesc}/{sContSysType}/{sDriveRating}/{sEquipRegFlag}/{sIPRegFlag}/{sIPAddress}/{sCheckInComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string UpdatePlantEquipItem(string sSessionId, string sUserId, string sFullName, string sPlantEquipNo, 
                                    string sName, string sDesc, string sLongDesc, string sContSysType,
                                    string sDriveRating, string sEquipRegFlag, string sIPRegFlag, string sIPAddress,
                                    string sCheckInComments, string sWebAppId);


        [OperationContract]
        [WebGet(UriTemplate = "emailmessage/{sSessionId}/{sUserId}/{sSubject}/{sBody}/{sAttachments}/{sRecipients}/{sCCRecipients}/{sBCCRecipients}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string emailmessage(string sSessionId, string sUserId, string sSubject, string sBody, string sAttachments, string sRecipients, string sCCRecipients, string sBCCRecipients, string sWebAppId);


    }


}
