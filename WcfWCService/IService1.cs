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
        [WebGet(UriTemplate = "createwcdoc2/{sSessionId}/{sUserId}/{sDocNo}/{sDocName}/{sProductName}/{sDocType}/{sFolderNameAndPath}/{sDesc}/{sOriginator}/{sOriginatorDocId}/{sJobCode}/{sRevision}/{sCheckInComments}/{iProdOrLibrary}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string CreateWCDoc2(string sSessionId, string sUserId, string sDocNo, string sDocName, string sProductName, string sDocType, string sFolderNameAndPath,
                           string sDesc, string sOriginator, string sOriginatorDocId, string sJobCode, string sRevision, string sCheckInComments, string iProdOrLibrary, string sWebAppId);


        [OperationContract]
        [WebGet(UriTemplate = "createrequirementdoc/{sSessionId}/{sUserId}/{sDocNo}/{sDocName}/{sProductName}/{sDocType}/{sFolderNameAndPath}/{sDesc}/{sOriginator}/{sOriginatorDocId}/{sJobCode}/{sRevision}/" +
                                                   "{sTargetDate}/{sForecastDate}/{sActualDate}/{sDateBasis}/{sComments}/{sCheckInComments}/{iProdOrLibrary}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string CreateRequirementDoc(string sSessionId, string sUserId, string sDocNo, string sDocName, string sProductName, string sDocType, string sFolderNameAndPath,
                                          string sDesc, string sOriginator, string sOriginatorDocId, string sJobCode, string sRevision,
                                          string sTargetDate, string sForecastDate, string sActualDate, string sDateBasis, string sComments,
                                          string sCheckInComments, string iProdOrLibrary, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "createplantequipmaterialdoc/{sSessionId}/{sUserId}/{sDocNo}/{sDocName}/{sProductName}/{sDocType}/{sFolderNameAndPath}/{sDesc}/{sOriginator}/{sOriginatorDocId}/{sJobCode}/{sRevision}/" +
                                                   "{sFirstIssueDate}/{sIssueForUseDate}/{sFinalIssueDate}/{sStatusComments}/{sComments}/{sCheckInComments}/{iProdOrLibrary}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string CreatePlantEquipMaterialDoc(string sSessionId, string sUserId, string sDocNo, string sDocName, string sProductName, string sDocType, string sFolderNameAndPath,
                                                  string sDesc, string sOriginator, string sOriginatorDocId, string sJobCode, string sRevision,
                                                  string sFirstIssueDate, string sIssueForUseDate, string sFinalIssueDate, string sStatusComments, string sComments,
                                                  string sCheckInComments, string iProdOrLibrary, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "updaterequirementdoc/{sSessionId}/{sUserId}/{sDocNo}/{sDocName}/{sDesc}/{sOriginator}/{sOriginatorDocId}/" +
                                                   "{sTargetDate}/{sForecastDate}/{sActualDate}/{sDateBasis}/{sComments}/{sCheckInComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string UpdateRequirementDoc(string sSessionId, string sUserId, string sDocNo, string sDocName,
                                          string sDesc, string sOriginator, string sOriginatorDocId, 
                                          string sTargetDate, string sForecastDate, string sActualDate, string sDateBasis,
                                          string sComments, string sCheckInComments, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "updateplantequipmaterialdoc/{sSessionId}/{sUserId}/{sDocNo}/{sDocName}/{sDesc}/{sOriginator}/{sOriginatorDocId}/" +
                                                   "{sFirstIssueDate}/{sIssueForUseDate}/{sFinalIssueDate}/{sStatusComments}/{sComments}/{sCheckInComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string UpdatePlantEquipMaterialDoc(string sSessionId, string sUserId, string sDocNo, string sDocName,
                                          string sDesc, string sOriginator, string sOriginatorDocId,
                                          string sFirstIssueDate, string sIssueForUseDate, string sFinalIssueDate, string sStatusComments,
                                          string sComments, string sCheckInComments, string sWebAppId);


        [OperationContract]
        [WebGet(UriTemplate = "createworkexecutionpackage/{sSessionId}/{sUserId}/{sWorkItemId}/{sAssignedActivityId}/{sRoute}/{sPlannedWorkPackageNo}/{sWEDName}/{sProductName}/{sDocType}/{sFolderNameAndPath}/{sOriginator}/{sJobCode}/{sNew}/{sExistingWEDNo}/{sWebAppId}/{sSkipCompleteTask}", ResponseFormat = WebMessageFormat.Xml)]
        string CreateWorkExecutionPackage(string sSessionId, string sUserId, string sWorkItemId, string sAssignedActivityId, string sRoute, 
                                          string sPlannedWorkPackageNo, string sWEDName, string sProductName, string sDocType, string sFolderNameAndPath,
                                          string sOriginator, string sJobCode, string sNew, string sExistingWEDNo, string sWebAppId, string sSkipCompleteTask);

        [OperationContract]
        [WebGet(UriTemplate = "setdocpartdescribedbylink/{sSessionId}/{sUserId}/{sDocNo}/{sPartNo}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string SetDocPartDescribedByLink(string sSessionId, string sUserId, string sDocNo, string sPartNo, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "createprojectmaterialitem/{sSessionId}/{sUserId}/{sFullName}/{sPartNo}/{sPartName}/{sProductName}/{sPartType}/{sFolderNameAndPath}/" +
                                                        "{sCheckInComments}/{sPartDescription}/{sComments}/{iProdOrLibrary}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string CreateProjectMaterialItem(string sSessionId, string sUserId, string sFullName, string sPartNo, string sPartName,
                                        string sProductName, string sPartType, string sFolderNameAndPath,
                                        string sCheckInComments, string sPartDescription, string sComments,
                                        string iProdOrLibrary, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "createfunctionallocationbasepart/{sSessionId}/{sUserId}/{sFullName}/{sPartNo}/{sPartName}/{sProductName}/{sPartType}/{sFolderNameAndPath}/" +
                                                        "{sCheckInComments}/{sPartDescription}/{sComments}/{iProdOrLibrary}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string CreateFunctionalLocationBasePart(string sSessionId, string sUserId, string sFullName, string sPartNo, string sPartName,
                                        string sProductName, string sPartType, string sFolderNameAndPath,
                                        string sCheckInComments, string sPartDescription, string sComments,
                                        string iProdOrLibrary, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "updateprojectmaterialitem/{sSessionId}/{sUserId}/{sFullName}/{sPartNo}/{sPartName}/{sCheckInComments}/{sPartDescription}/{sComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string UpdateProjectMaterialItem(string sSessionId, string sUserId, string sFullName, string sPartNo, string sPartName,
                                                string sCheckInComments, string sPartDescription, string sComments,
                                                string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "createprojectworkitem/{sSessionId}/{sUserId}/{sFullName}/{sParentPartNo}/{sPartNo}/{sPartName}/{sProductName}/{sPartType}/{sPartUsageType}/{sPartUsageUnit}/{sFolderNameAndPath}/{sCheckInComments}/{sLineNumber}/{sPartDescription}/{sReqirementsInfo}/{sPreparationInfo}/{sReviewInfo}/{iProdOrLibrary}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string CreateProjectWorkItem(string sSessionId, string sUserId, string sFullName, string sParentPartNo, string sPartNo, string sPartName,
                                            string sProductName, string sPartType, string sPartUsageType, string sPartUsageUnit, string sFolderNameAndPath,
                                            string sCheckInComments, string sLineNumber, string sPartDescription,
                                            string sReqirementsInfo, string sPreparationInfo, string sReviewInfo, string iProdOrLibrary, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "insertexistingprojectworkitem/{sSessionId}/{sUserId}/{sFullName}/{sParentPartNo}/{sExistingPWIPartNo}/{sPartUsageType}/{sPartUsageUnit}/{sCheckInComments}/{sLineNumber}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string InsertExistingProjectWorkItem(string sSessionId, string sUserId, string sFullName, string sParentPartNo, string sExistingPWIPartNo,
                                                    string sPartUsageType, string sPartUsageUnit, string sCheckInComments, string sLineNumber, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "createproject/{sSessionId}/{sUserId}/{sFullName}/{sPartNo}/{sPartName}/{sProductName}/{sPartType}/{sFolderNameAndPath}/{sCheckInComments}/{sPartDescription}/{sReqirementsInfo}/{sPreparationInfo}/{sReviewInfo}/{iProdOrLibrary}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string CreateProject(string sSessionId, string sUserId, string sFullName, string sPartNo, string sPartName,
                                            string sProductName, string sPartType, string sFolderNameAndPath,
                                            string sCheckInComments, string sPartDescription,
                                            string sReqirementsInfo, string sPreparationInfo, string sReviewInfo, string iProdOrLibrary, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "createfronesisproject/{sSessionId}/{sUserId}/{sProjNo}/{sProjDesc}/{sProductName}/{sDocType}/{sPartType}/{sFolderNameAndPath}/{sClientDesc}/{sOriginator}/{sClientProjNo}/{sRevision}/{sCheckInComments}/{iProdOrLibrary}/{sWebAppId}/{sProjType}", ResponseFormat = WebMessageFormat.Xml)]
        string CreateFronesisProject(string sSessionId, string sUserId, string sProjNo, string sProjDesc, string sProductName, string sDocType, string sPartType, string sFolderNameAndPath,
                                  string sClientDesc, string sOriginator, string sClientProjNo, string sRevision, string sCheckInComments, string iProdOrLibrary, string sWebAppId, string sProjType);

        [OperationContract]
        [WebGet(UriTemplate = "createfronesisprojectchilddoc/{sSessionId}/{sUserId}/{sProjNo}/{sChildDocNo}/{sChildDocName}/{sProductName}/{sDocType}/{sFolderNameAndPath}/{sOriginator}/{sRevision}/{sCheckInComments}/{iProdOrLibrary}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string CreateFronesisProjectChildDoc(string sSessionId, string sUserId, string sProjNo, string sChildDocNo, string sChildDocName, string sProductName, string sDocType, string sFolderNameAndPath,
                                             string sOriginator, string sRevision, string sCheckInComments, string iProdOrLibrary, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "attachurl/{sSessionId}/{sUserId}/{sFullname}/{sDocNo}/{sURLDesc}/{sURL}/{bSecondary}/{sAttachComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string AttachURL(string sSessionId, string sUserId, string sFullName, string sDocNo, string sURLDesc, string sURL, string bSecondary, string sAttachComments, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "attachwcdoc/{sSessionId}/{sUserId}/{sFullname}/{sDocNo}/{sAttachDesc}/{sAttachPath}/{bSecondary}/{sAttachComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string AttachWCDoc(string sSessionId, string sUserId, string sFullName, string sDocNo, string sAttachDesc, string sAttachPath, string bSecondary, string sAttachComments, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "deletewcdoc/{sSessionId}/{sUserId}/{sFullname}/{sDocNo}/{sAttachFile}/{bSecondary}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string DeleteWCDoc(string sSessionId, string sUserId, string sFullName, string sDocNo, string sAttachFile, string bSecondary, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "deleteurl/{sSessionId}/{sUserId}/{sFullname}/{sDocNo}/{sURL}/{bSecondary}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string DeleteURL(string sSessionId, string sUserId, string sFullName, string sDocNo, string sURL, string bSecondary, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "setdocattributestrings/{sSessionId}/{sUserId}/{sDocNo}/{sDocName}/{sLongDesc}/{sOriginator}/{sOriginatorDocId}/{sJobCode}/{sCheckInComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string SetDocAttributeStrings(string sSessionId, string sUserId, string sDocNo, string sDocName, string sLongDesc, string sOriginator, string sOriginatorDocId, string sJobCode,
                                      string sCheckInComments, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "setdocattributestrings2/{sSessionId}/{sUserId}/{sDocNo}/{sDocName}/{sDesc}/{sOriginator}/{sOriginatorDocId}/{sJobCode}/{sCheckInComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string SetDocAttributeStrings2(string sSessionId, string sUserId, string sDocNo, string sDocName, string sDesc, string sOriginator, string sOriginatorDocId, string sJobCode,
                                      string sCheckInComments, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "setdoctodocref/{sSessionId}/{sUserId}/{sFullname}/{sDocNo}/{sReferencedDocNo}/{sCheckinComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string SetDocToDocRef(string sSessionId, string sUserId, string sFullName, string sDocNo, string sReferencedDocNo, string sCheckinComments, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "setdoctodocrefs/{sSessionId}/{sUserId}/{sFullname}/{sDocNo}/{sReferencedDocNos}/{sCheckinComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string SetDocToDocRefs(string sSessionId, string sUserId, string sFullName, string sDocNo, string sReferencedDocNos, string sCheckinComments, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "setdoctodoclink/{sSessionId}/{sUserId}/{sFullname}/{sParentDoc}/{sChildDocNo}/{sCheckinComments}/{sDocUsageType}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string setDocToDocLink(string sSessionId, string sUserId, string sFullName, string sParentDoc, string sChildDocNo, string sCheckInComments, string sDocUsageType, string sWebAppId);


        [OperationContract]
        [WebGet(UriTemplate = "setdocreviewer/{sSessionId}/{sUserId}/{sFullname}/{sDocNo}/{sReviewerNo}/{sCheckinComments}/{sReviewerTypeName}/{sCompletionDate}/{sCompletionStatus}/{sComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string SetDocReviewer(string sSessionId, string sUserId, string sFullName, string sDocNo, string sReviewerNo, string sCheckinComments, string sReviewerTypeName, string sCompletionDate, string sCompletionStatus, string sComments, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "setdocreviewerfordocrevision/{sSessionId}/{sUserId}/{sFullname}/{sDocNo}/{sDocRev}/{sReviewerNo}/{sCheckinComments}/{sReviewerTypeName}/{sCompletionDate}/{sCompletionStatus}/{sComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string SetDocReviewerForDocRevision(string sSessionId, string sUserId, string sFullName, string sDocNo, string sDocRev, string sReviewerNo, string sCheckinComments, string sReviewerTypeName, string sCompletionDate, string sCompletionStatus, string sComments, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "updatedocreviewer/{sSessionId}/{sUserId}/{sFullname}/{sDocNo}/{sReviewerNo}/{sCheckinComments}/{sReviewerTypeName}/{sCompletionDate}/{sCompletionStatus}/{sComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string UpdateDocReviewer(string sSessionId, string sUserId, string sFullName, string sDocNo, string sReviewerNo, string sCheckinComments, string sReviewerTypeName, string sCompletionDate, string sComments, string sCompletionStatus, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "updatedocreviewerfordocrevision/{sSessionId}/{sUserId}/{sFullname}/{sDocNo}/{sDocRev}/{sReviewerNo}/{sCheckinComments}/{sReviewerTypeName}/{sCompletionDate}/{sCompletionStatus}/{sComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string UpdateDocReviewerForDocRevision(string sSessionId, string sUserId, string sFullName, string sDocNo, string sDocRev, string sReviewerNo, string sCheckinComments, string sReviewerTypeName, string sCompletionDate, string sComments, string sCompletionStatus, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "setpartreviewer/{sSessionId}/{sUserId}/{sFullname}/{sPartNo}/{sReviewerNo}/{sPartRefLinkType}/{sCheckinComments}/{sReviewerTypeName}/{sCompletionDate}/{sCompletionStatus}/{sComments}/{sAccountableFlag}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string SetPartReviewer(string sSessionId, string sUserId, string sFullName, string sPartNo, string sReviewerNo, string sPartRefLinkType, string sCheckinComments, string sReviewerTypeName, 
                               string sCompletionDate, string sCompletionStatus, string sComments, string sAccountableFlag, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "updatepartreviewer/{sSessionId}/{sUserId}/{sFullname}/{sPartNo}/{sReviewerNo}/{sCheckinComments}/{sReviewerTypeName}/{sCompletionDate}/{sCompletionStatus}/{sComments}/{sAccountableFlag}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string UpdatePartReviewer(string sSessionId, string sUserId, string sFullName, string sPartNo, string sReviewerNo, string sCheckinComments, string sReviewerTypeName, 
                                  string sCompletionDate, string sCompletionStatus, string sComments, string sAccountableFlag, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "deletedoctodocusagelink/{sSessionId}/{sUserId}/{sFullname}/{sParentDocNo}/{sChildDocNo}/{sCheckinComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string DeleteDocToDocUsageLink(string sSessionId, string sUserId, string sFullName, string sParentDocNo, string sChildDocNo, string sCheckinComments, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "deletedoctodocusagelinkwithstringattribute/{sSessionId}/{sUserId}/{sFullname}/{sParentDocNo}/{sChildDocNo}/{sCheckinComments}/{sAttributeName}/{sAttributeValue}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string DeleteDocToDocUsageLinkWithStringAttribute(string sSessionId, string sUserId, string sFullName, string sParentDocNo, string sChildDocNo, string sCheckinComments, string sAttributeName, string sAttributeValue, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "deletedoctodocusagelinkwithstringattributefordocrevision/{sSessionId}/{sUserId}/{sFullname}/{sParentDocNo}/{sParentDocRev}/{sChildDocNo}/{sCheckinComments}/{sAttributeName}/{sAttributeValue}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string DeleteDocToDocUsageLinkWithStringAttributeForDocRevision(string sSessionId, string sUserId, string sFullName, string sParentDocNo, string sParentDocRev, string sChildDocNo, string sCheckinComments, string sAttributeName, string sAttributeValue, string sWebAppId);


        [OperationContract]
        [WebGet(UriTemplate = "deletedoctodocref/{sSessionId}/{sUserId}/{sFullname}/{sDocNo}/{sReferencedDocNo}/{sCheckinComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string DeleteDocToDocRef(string sSessionId, string sUserId, string sFullName, string sDocNo, string sReferencedDocNo, string sCheckinComments, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "deletedoctodocrefs/{sSessionId}/{sUserId}/{sFullname}/{sDocNo}/{sReferencedDocNos}/{sCheckinComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string DeleteDocToDocRefs(string sSessionId, string sUserId, string sFullName, string sDocNo, string sReferencedDocNos, string sCheckinComments, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "setdoctopartref/{sSessionId}/{sUserId}/{sFullname}/{sDocNo}/{sPartNo}/{sCheckinComments}/{sPartRefLinkType}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string SetDocToPartRef(string sSessionId, string sUserId, string sFullName, string sDocNo, string sPartNo, string sCheckinComments, string sPartRefLinkType, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "setdeliverabledoctopartref/{sSessionId}/{sUserId}/{sFullname}/{sDocNo}/{sPartNo}/{sCheckinComments}/{sPartRefLinkType}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string SetDeliverableDocToPartRef(string sSessionId, string sUserId, string sFullName, string sDocNo, string sPartNo, string sCheckinComments, string sPartRefLinkType, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "setdeliverableparttopartlink/{sSessionId}/{sUserId}/{sFullname}/{sParentPart}/{sChildPart}/{sCheckinComments}/{sPartUsageLinkType}/{sLineNumber}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string SetDeliverablePartToPartLink(string sSessionId, string sUserId, string sFullName, string sParentPart, string sChildPart, string sCheckinComments, string sPartUsageLinkType, string sLineNumber, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "setdoctopartrefs/{sSessionId}/{sUserId}/{sFullname}/{sDocNo}/{sPartNos}/{sCheckinComments}/{sPartRefLinkType}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string SetDocToPartRefs(string sSessionId, string sUserId, string sFullName, string sDocNo, string sPartNos, string sCheckinComments, string sPartRefLinkType, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "deletedoctopartref/{sSessionId}/{sUserId}/{sFullname}/{sDocNo}/{sPartNo}/{sCheckinComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string DeleteDocToPartRef(string sSessionId, string sUserId, string sFullName, string sDocNo, string sPartNo, string sCheckinComments, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "deletedoctopartrefs/{sSessionId}/{sUserId}/{sFullname}/{sDocNo}/{sPartNos}/{sCheckinComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string DeleteDocToPartRefs(string sSessionId, string sUserId, string sFullName, string sDocNo, string sPartNos, string sCheckinComments, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "deletedoctopartrefwithattribute/{sSessionId}/{sUserId}/{sFullname}/{sDocNo}/{sPartNo}/{sAttributeName}/{sAttributeValue}/{sCheckinComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string DeleteDocToPartRefWithAttribute(string sSessionId, string sUserId, string sFullName, string sDocNo, string sPartNo, string sAttributeName, string sAttributeValue, string sCheckinComments, string sWebAppId);

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
        [WebGet(UriTemplate = "updatprojectstatus/{sSessionId}/{sUserId}/{sFullname}/{sProjCode}/{sProjName}/{sProjStatus}/{sCheckInComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string UpdateProjectStatus(string sSessionId, string sUserId, string sFullName, string sProjCode, string sProjName, string sProjStatus, string sCheckinComments, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "renamepart/{sSessionId}/{sUserId}/{sFullname}/{sPartNo}/{sNewPartNo}/{sNewPartName}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string RenamePart(string sSessionId, string sUserId, string sFullName, string sPartNo, string sNewPartNo, string sNewPartName, string sWebAppId);


        [OperationContract]
        [WebGet(UriTemplate = "renamedocument/{sSessionId}/{sUserId}/{sFullname}/{sDocumentNo}/{sNewDocumentNo}/{sNewDocumentName}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string RenameDocument(string sSessionId, string sUserId, string sFullName, string sDocumentNo, string sNewDocumentNo, string sNewDocumentName, string sWebAppId);


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
        [WebGet(UriTemplate = "createtechnicalaction/{sSessionId}/{sUserId}/{sFullname}/{sTechActionNo}/{sTechActionName}/{sProductName}/{sPRType}/{sFolderNameAndPath}/{sPlantCode}/" +
                                                    "{sTechActionDesc}/{sComments}/{sNeedDate}/{iProdOrLibrary}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string CreateTechnicalAction(string sSessionId, string sUserId, string sFullName, string sTechActionNo, string sTechActionName, string sProductName, string sPRType, string sFolderNameAndPath,
                                           string sPlantCode, string sTechActionDesc, string sComments, string sNeedDate, string iProdOrLibrary, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "updatetechnicalaction/{sSessionId}/{sUserId}/{sFullname}/{sTechActionNo}/{sTechActionName}/{sTechActionDesc}/{sComments}/{sNeedDate}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string UpdateTechnicalAction(string sSessionId, string sUserId, string sFullName, string sTechActionNo, string sTechActionName, string sTechActionDesc,
                                     string sComments, string sNeedDate, string sWebAppId);

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
        [WebGet(UriTemplate = "settasknextelapseddateoncompletion/{sSessionId}/{sUserId}/{sWorkItemId}/{sAssignedActivityId}/{sRoute}/{sNextElapsedDate}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string SetTaskNextElapsedDateOnCompletion(string sSessionId, string sUserId, string sWorkItemId, string sAssignedActivityId, string sRoute, string sNextElapsedDate, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "settaskoperationalhoursoncompletion/{sSessionId}/{sUserId}/{sWorkItemId}/{sAssignedActivityId}/{sHoursOnCompletion}/{sRoute}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string SetTaskOperationalHoursOnCompletion(string sSessionId, string sUserId, string sWorkItemId, string sAssignedActivityId, string sHoursOnCompletion, string sRoute, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "settaskwocompletiondate/{sSessionId}/{sUserId}/{sWorkItemId}/{sAssignedActivityId}/{sRoute}/{sDateOnCompletion}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string SetTaskWOCompletionDate(string sSessionId, string sUserId, string sWorkItemId, string sAssignedActivityId, string sRoute, string sDateOnCompletion, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "settaskcontroldocexpirydate/{sSessionId}/{sUserId}/{sWorkItemId}/{sAssignedActivityId}/{sRoute}/{sExpiryDate}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string SetTaskControlDocExpiryDate(string sSessionId, string sUserId, string sWorkItemId, string sAssignedActivityId, string sRoute, string sExpiryDate, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "progresstask/{sSessionId}/{sUserId}/{sWorkItemId}/{sAssignedActivityId}/{sRoute}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string ProgressTask(string sSessionId, string sUserId, string sWorkItemId, string sAssignedActivityId, string sRoute,  string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "setprobrptaffectedobjects/{sSessionId}/{sUserId}/{sProdLossNo}/{sAffectdPartsString}/{sAffectdObjectTypesString}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string SetProbRptAffectedObjects(string sSessionId, string sUserId, string sProdLossNo, string sAffectdPartsString, string sAffectdObjectTypesString, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "setprobrptstate/{sSessionId}/{sUserId}/{sFullName}/{sProbRptNo}/{sProbRptName}/{sLifecycleState}/{sComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string SetProbRptState(string sSessionId, string sUserId, string sFullName, string sProbRptNo, string sProbRptName, string sLifecycleState, string sComments, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "deleteprobrptaffectedobjects/{sSessionId}/{sUserId}/{sProdLossNo}/{sAffectdPartsString}/{sAffectdObjectTypesString}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string DeleteProbRptAffectedObjects(string sSessionId, string sUserId, string sProdLossNo, string sAffectdPartsString, string sAffectdObjectTypesString, string sWebAppId);

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
        [WebGet(UriTemplate = "updatembapartusagelinkfrombatch/{sSessionId}/{sUserId}/{sFullname}/{sParentPartNo}/{sChildPartNo}/{dQty}/{lLineNumber}/{sCheckInComments}/{sDispatchDocketNo}/{sTransactionDate}/{sComments}/{sMoisturePercentage}/{sInvoiceStatus}/{sBatchNo}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string UpdateMBAPartUsageLinkFromBatch(string sSessionId, string sUserId, string sFullName, string sParentPartNo, string sChildPartNo,
                                               string dQty, string lLineNumber, string sCheckInComments, string sDispatchDocketNo,
                                               string sTransactionDate, string sComments, string sMoisturePercentage, string sInvoiceStatus,
                                               string sBatchNo, string sWebAppId);

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
                              "{dTargetC}/{dActualC}/{dTargetSO3}/{dActualSO3}/{dTargetCN}/{dActualCN}/{sProductCode}/{sBatchDate}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string CreateBatch(string sSessionId, string sUserId, string sFullName, string sBatchNo, string sBatchName, string sProductName, string sFolder, string sBatchType, string sCheckInComments, string iProdOrLibrary,
                          string dTargetQty, string dActualQty, string dMoisturePercentage, string sQualityStatus,
                          string dTargetAl2O3, string dActualAl2O3, string dTargetCaO, string dActualCaO, string dTargetF, string dActualF,
                          string dTargetFe2O3, string dActualFe2O3, string dTargetK2O, string dActualK2O, string dTargetMgO, string dActualMgO,
                          string dTargetMnO, string dActualMnO, string dTargetNa2O3, string dActualNa2O3, string dTargetSiO2, string dActualSiO2,
                          string dTargetC, string dActualC, string dTargetSO3, string dActualSO3, string dTargetCN, string dActualCN, 
                          string sProductCode, string sBatchDate, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "createmba/{sSessionId}/{sUserId}/{sFullname}/{sBatchNo}/{sBatchName}/{sProductName}/{sFolder}/" +
                              "{sBatchType}/{sCheckInComments}/{iProdOrLibrary}/{dMoisturePercentage}/{sProductCode}/{sBatchDate}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string CreateMBA(string sSessionId, string sUserId, string sFullName, string sBatchNo, string sBatchName, string sProductName, string sFolder,
                         string sBatchType, string sCheckInComments, string iProdOrLibrary, string dMoisturePercentage, string sProductCode, string sBatchDate, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "setshippingloaditem/{sSessionId}/{sUserId}/{sFullname}/{sBookingNo}/{sContainerNo}/{sContainerTare}/{sLoadNo}/{sSealNo}/" +
                              "{sBatchNo}/{sBatchLineNumber}/{sBatchQty}/{sItemComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string SetShippingLoadItem(string sSessionId, string sUserId, string sFullName, string sBookingNo, string sContainerNo, string sContainerTare, string sLoadNo, string sSealNo,
                                   string sBatchNo, string sBatchLineNumber, string sBatchQty, string sItemComments, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "createshippingload/{sSessionId}/{sUserId}/{sFullname}/{sBatchNo}/{sBatchName}/{sProductName}/{sFolder}/" +
                              "{sBatchType}/{sCheckInComments}/{iProdOrLibrary}/{dMoisturePercentage}/{sContainerSealNo}/{sProductCode}/{sBatchDate}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string CreateShippingLoad(string sSessionId, string sUserId, string sFullName, string sBatchNo, string sBatchName, string sProductName, string sFolder,
                         string sBatchType, string sCheckInComments, string iProdOrLibrary, string dMoisturePercentage, string sContainerSealNo,
                         string sProductCode, string sBatchDate, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "createshippingcontainer/{sSessionId}/{sUserId}/{sFullname}/{sBatchNo}/{sBatchName}/{sProductName}/{sFolder}/" +
                              "{sBatchType}/{sCheckInComments}/{iProdOrLibrary}/{dMoisturePercentage}/{dTareWeight}/{sProductCode}/{sBatchDate}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string CreateShippingContainer(string sSessionId, string sUserId, string sFullName, string sBatchNo, string sBatchName, string sProductName, string sFolder,
                                       string sBatchType, string sCheckInComments, string iProdOrLibrary, string dMoisturePercentage, string dTareWeight, 
                                       string sProductCode, string sBatchDate, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "createshippingbooking/{sSessionId}/{sUserId}/{sFullname}/{sBatchNo}/{sBatchName}/{sProductName}/{sFolder}/" +
                              "{sBatchType}/{sCheckInComments}/{iProdOrLibrary}/{sProductCode}/{sComments}/{sBatchDate}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string CreateShippingBooking(string sSessionId, string sUserId, string sFullName, string sBatchNo, string sBatchName, string sProductName, string sFolder,
                         string sBatchType, string sCheckInComments, string iProdOrLibrary,
                         string sProductCode, string sComments, string sBatchDate, string sWebAppId);


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
                              "{dTargetC}/{dActualC}/{dTargetSO3}/{dActualSO3}/{dTargetCN}/{dActualCN}/{sProductCode}/{sBatchDate}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string UpdateBatch(string sSessionId, string sUserId, string sFullName, string sBatchNo, string sBatchName, string sCheckinComments,
                                  string dTargetQty, string dActualQty, string dMoisturePercentage, string sQualityStatus,
                                  string dTargetAl2O3, string dActualAl2O3, string dTargetCaO, string dActualCaO, string dTargetF, string dActualF,
                                  string dTargetFe2O3, string dActualFe2O3, string dTargetK2O, string dActualK2O, string dTargetMgO, string dActualMgO,
                                  string dTargetMnO, string dActualMnO, string dTargetNa2O3, string dActualNa2O3, string dTargetSiO2, string dActualSiO2,
                                  string dTargetC, string dActualC, string dTargetSO3, string dActualSO3, string dTargetCN, string dActualCN, string sProductCode, 
                                  string sBatchDate, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "updatebatchqty/{sSessionId}/{sUserId}/{sFullname}/{sBatchNo}/{sBatchName}/{sCheckinComments}/{dActualQty}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string UpdateBatchQty(string sSessionId, string sUserId, string sFullName, string sBatchNo, string sBatchName, string sCheckinComments,
                                  string dActualQty, string sWebAppId);


        [OperationContract]
        [WebGet(UriTemplate = "updatemba/{sSessionId}/{sUserId}/{sFullname}/{sBatchNo}/{sBatchName}/{sCheckinComments}/{dMoisturePercentage}/{sBatchDate}/{sComments}/{sProductCode}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string UpdateMBA(string sSessionId, string sUserId, string sFullName, string sBatchNo, string sBatchName, string sCheckinComments, string dMoisturePercentage, string sBatchDate, string sComments, string sProductCode, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "updateshippingload/{sSessionId}/{sUserId}/{sFullname}/{sBatchNo}/{sBatchName}/{sCheckinComments}/{dMoisturePercentage}/{sBatchDate}/{sComments}/{sContainerSealNo}/{sProductCode}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string UpdateShippingLoad(string sSessionId, string sUserId, string sFullName, string sBatchNo, string sBatchName, string sCheckinComments, string dMoisturePercentage, string sBatchDate, string sComments, string sContainerSealNo, string sProductCode, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "updateshippingcontainer/{sSessionId}/{sUserId}/{sFullname}/{sBatchNo}/{sBatchName}/{sCheckinComments}/{dMoisturePercentage}/{sBatchDate}/{sComments}/{dTareWeight}/{sProductCode}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string UpdateShippingContainer(string sSessionId, string sUserId, string sFullName, string sBatchNo, string sBatchName, string sCheckinComments, string dMoisturePercentage, string sBatchDate, string sComments, string dTareWeight, string sProductCode, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "updateshippingbooking/{sSessionId}/{sUserId}/{sFullname}/{sBatchNo}/{sBatchName}/{sCheckinComments}/{sBatchDate}/{sComments}/{sProductCode}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string UpdateShippingBooking(string sSessionId, string sUserId, string sFullName, string sBatchNo, string sBatchName, string sCheckinComments, string sBatchDate, string sComments,string sProductCode, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "createprodorder/{sSessionId}/{sUserId}/{sDocNo}/{sDocName}/{sProductName}/{sDocType}/{sFolderNameAndPath}/" +
                              "{sBatchNo}/{sTargetQty}/{sProdNoDate}/{sOriginator}/{sJobCode}/{sComments}/{sRevision}/{sCheckInComments}/" +
                              "{iProdOrLibrary}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string CreateProdOrder(string sSessionId, string sUserId, string sDocNo, string sDocName, string sProductName, string sDocType, string sFolderNameAndPath,
                                      string sBatchNo, string sTargetQty, string sProdNoDate, string sOriginator, string sJobCode, string sComments,
                                      string sRevision, string sCheckInComments, string iProdOrLibrary, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "createcableschedule/{sSessionId}/{sUserId}/{sDocNo}/{sDocName}/{sProductName}/{sDocType}/{sFolderNameAndPath}/" +
                              "{sOriginator}/{sJobCode}/{sComments}/{sRevision}/{sCheckInComments}/{iProdOrLibrary}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string CreateCableSchedule(string sSessionId, string sUserId, string sDocNo, string sDocName, string sProductName, string sDocType, string sFolderNameAndPath,
                                          string sOriginator, string sJobCode, string sComments, string sRevision, string sCheckInComments, string iProdOrLibrary, string sWebAppId);


        [OperationContract]
        [WebGet(UriTemplate = "createcablescheduleitem/{sSessionId}/{sUserId}/{sCSNo}/{sProductName}/{sFolderNameAndPath}/{sCableNo}/" +
                              "{sCableName}/{sFromFL}/{sToFL}/{dLength}/{sFromLineNumber}/{sToLineNumber}/" +
                              "{sMaterialCableCode}/{sOriginator}/{sCableComments}/{sCableCheckInComments}/{iProdOrLibrary}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string CreateCableScheduleItem(string sSessionId, string sUserId, string sCSNo, string sProductName, string sFolderNameAndPath,
                                      string sCableNo, string sCableName, string sFromFL, string sToFL, string dLength, string sFromLineNumber,
                                      string sToLineNumber, string sMaterialCableCode,
                                      string sOriginator, string sCableComments, string sCableCheckInComments, string iProdOrLibrary, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "createcableitem/{sSessionId}/{sUserId}/{sProductName}/{sFolderNameAndPath}/{sCableNo}/" +
                              "{sCableName}/{sFromFL}/{sToFL}/{dLength}/{sFromLineNumber}/{sToLineNumber}/" +
                              "{sMaterialCableCode}/{sOriginator}/{sCableComments}/{sCableCheckInComments}/{iProdOrLibrary}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string CreateCableItem(string sSessionId, string sUserId, string sProductName, string sFolderNameAndPath,
                                string sCableNo, string sCableName, string sFromFL, string sToFL, string dLength, string sFromLineNumber,
                                string sToLineNumber, string sMaterialCableCode,
                                string sOriginator, string sCableComments, string sCableCheckInComments, string iProdOrLibrary, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "updatecableitem/{sSessionId}/{sUserId}/{sCableNo}/{sCableName}/{sOriginator}/{sCableComments}/{sCableCheckInComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string UpdateCableItem(string sSessionId, string sUserId, string sCableNo, string sCableName, string sOriginator, string sCableComments, string sCableCheckInComments, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "updatecablematerial/{sSessionId}/{sUserId}/{sFullName}/{sCableNo}/{sLength}/{sMaterialCode}/{sOldMaterialCode}/{sCableCheckInComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string UpdateCableMaterial(string sSessionId, string sUserId, string sFullName, string sCableNo,
                                          string sLength, string sMaterialCode, string sOldMaterialCode,
                                          string sCableCheckInComments, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "setpartstate/{sSessionId}/{sUserId}/{sPartNo}/{sLifecycleState}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string SetPartState(string sSessionId, string sUserId, string sPartNo, string sLifecycleState, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "setdocumentstate/{sSessionId}/{sUserId}/{sDocumentNo}/{sLifecycleState}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string SetDocumentState(string sSessionId, string sUserId, string sDocumentNo, string sLifecycleState, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "createcablepartlink/{sSessionId}/{sUserId}/{sCableNo}/{sFuncLoc}/{sLineNumber}/{sToOrFrom}/{sCheckInComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string CreateCablePartLink(string sSessionId, string sUserId, string sCableNo, string sFuncLoc, string sLineNumber, string sToOrFrom, string sCheckInComments, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "updatecablefromdetails/{sSessionId}/{sUserId}/{sCableNo}/{sNewFuncLoc}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string UpdateCableFromDetails(string sSessionId, string sUserId, string sCableNo, string sNewFuncLoc, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "createcableterminationlink/{sSessionId}/{sUserId}/{sCableNo}/{sFuncLoc}/{sLineNumber}/{sToOrFrom}/{sTermination}/{sWireNo}/{sCoreNo}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string CreateCableTerminationLink(string sSessionId, string sUserId, string sCableNo, string sFuncLoc, string sLineNumber, string sToOrFrom, string sTermination, string sWireNo, string sCoreNo, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "createcableterminationlink2/{sSessionId}/{sUserId}/{sCableNo}/{sFuncLoc}/{sLineNumber}/{sToOrFrom}/{sTermination}/{sWireNo}/{sCoreNo}/{sCoreLabel}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string CreateCableTerminationLink2(string sSessionId, string sUserId, string sCableNo, string sFuncLoc, string sLineNumber, string sToOrFrom, string sTermination, 
                                           string sWireNo, string sCoreNo, string sCoreLabel, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "updatecableterminationlink/{sSessionId}/{sUserId}/{sCableNo}/{sFuncLoc}/{sLineNumber}/{sToOrFrom}/{sTermination}/{sWireNo}/{sCoreNo}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string UpdateCableTerminationLink(string sSessionId, string sUserId, string sCableNo, string sFuncLoc, string sLineNumber, string sToOrFrom, string sTermination, string sWireNo, string sCoreNo, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "updatecableterminationlink2/{sSessionId}/{sUserId}/{sCableNo}/{sFuncLoc}/{sLineNumber}/{sToOrFrom}/{sTermination}/{sWireNo}/{sCoreNo}/{sCoreLabel}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string UpdateCableTerminationLink2(string sSessionId, string sUserId, string sCableNo, string sFuncLoc, string sLineNumber, string sToOrFrom, string sTermination, 
                                           string sWireNo, string sCoreNo, string sCoreLabel, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "createtestandtagitem/{sSessionId}/{sUserId}/{sProductName}/{sFolderNameAndPath}/{sGroupNo}/{sTestAndTagItemNo}/" +
                                                   "{sTestAndTagName}/{sTestAndTagDate}/{sTestAndTagResult}/{sTestAndTagTagNumber}/" +
                                                   "{sTestAndTagMaintenanceActionNo}/{sCommonActionNo}/{sNextTestDate}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string CreateTestAndTagItem(string sSessionId, string sUserId, string sProductName, string sFolderNameAndPath, string sGroupNo,
                                           string sTestAndTagItemNo, string sTestAndTagName, string sTestAndTagDate, string sTestAndTagResult,
                                           string sTestAndTagTagNumber, string sTestAndTagMaintenanceActionNo, string sCommonActionNo,
                                           string sNextTestDate, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "updatetestandtagitem/{sSessionId}/{sUserId}/{sTestAndTagItemNo}/" +
                                                   "{sTestAndTagName}/{sTestAndTagDate}/{sTestAndTagResult}/{sTestAndTagTagNumber}/" +
                                                   "{sNextTestDate}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string UpdateTestAndTagItem(string sSessionId, string sUserId,
                                           string sTestAndTagItemNo, string sTestAndTagName, string sTestAndTagDate, string sTestAndTagResult,
                                           string sTestAndTagTagNumber,
                                           string sNextTestDate, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "deletetestandtagitem/{sSessionId}/{sUserId}/{sFullName}/" +
                                                   "{sGroupNo}/{sTestAndTagItemNo}/{sCheckInComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string DeleteTestAndTagItem(string sSessionId, string sUserId, string sFullName, string sGroupNo, string sTestAndTagItemNo, string sCheckInComments, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "creatematerialcatalogitem/{sSessionId}/{sUserId}/{sFullname}/{sMatCatNo}/{sMatCatType}/{sName}/{sDesc}/{sLongDesc}/{sDrivekW}/{sFullLoadCurrent}/{sUnitWeight}/{sLeadTime}/{sRepairable}/{sSpareRqd}/{sCheckInComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string CreateMaterialCatalogItem(string sSessionId, string sUserId, string sFullName, string sMatCatNo, string sMatCatType, string sName, 
                                         string sDesc, string sLongDesc, string sDrivekW, string sFullLoadCurrent,
                                         string sUnitWeight, string sLeadTime, string sRepairable, string sSpareRqd, string sCheckInComments, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "updatematerialcatalogitem/{sSessionId}/{sUserId}/{sFullname}/{sMatCatNo}/{sMatCatNewType}/{sMatCatOldType}/{sName}/{sDesc}/{sLongDesc}/{sDrivekW}/{sFullLoadCurrent}/{sUnitWeight}/{sLeadTime}/{sRepairable}/{sSpareRqd}/{sCheckInComments}/{sWebAppId}/{sNewLink}", ResponseFormat = WebMessageFormat.Xml)]
        string UpdateMaterialCatalogItem(string sSessionId, string sUserId, string sFullName, string sMatCatNo, string sMatCatNewType, string sMatCatOldType, string sName, 
                                         string sDesc, string sLongDesc, string sDrivekW, string sFullLoadCurrent,
                                         string sUnitWeight, string sLeadTime, string sRepairable, string sSpareRqd, string sCheckInComments, string sWebAppId, string sNewLink);


        [OperationContract]
        [WebGet(UriTemplate = "createplantequipitem/{sSessionId}/{sUserId}/{sFullname}/{sPlantEquipNo}/{sPlantEquipType}/{sName}/{sDesc}/{sLongDesc}/{sContSysType}/{sDriveRating}/" +
                                                   "{sEquipRegFlag}/{sIPRegFlag}/{sIPAddress}/{sComments}/{sOpZone}/{sProduct}/{sFolder}/" +
                                                   "{sPowerCable}/{sControlCable}/{sInstrumentationCable}/{sDataCable}/{sEarthCable}/" +
                                                   "{sInstRegFlag}/{sFullLoadCurrent}/{sConstructionDate}/{sFLGrouping}/" +
                                                   "{sCheckInComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string CreatePlantEquipItem(string sSessionId, string sUserId, string sFullName, string sPlantEquipNo,
                                           string sPlantEquipType, string sName, string sDesc, string sLongDesc,
                                           string sContSysType, string sDriveRating, string sEquipRegFlag, string sIPRegFlag, string sIPAddress,
                                           string sComments, string sOpZone,
                                           string sProduct, string sFolder,
                                           string sPowerCable, string sControlCable, string sInstrumentationCable,
                                           string sDataCable, string sEarthCable,
                                           string sInstRegFlag, string sFullLoadCurrent, string sConstructionDate, string sFLGrouping,
                                           string sCheckInComments, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "updateplantequipitem/{sSessionId}/{sUserId}/{sFullname}/{sPlantEquipNo}/{sName}/{sDesc}/{sLongDesc}/{sContSysType}/{sDriveRating}/{sEquipRegFlag}/" +
                                                   "{sIPRegFlag}/{sIPAddress}/{sComments}/{sOpZone}/" +
                                                   "{sPowerCable}/{sControlCable}/{sInstrumentationCable}/{sDataCable}/{sEarthCable}/" +
                                                   "{sInstRegFlag}/{sFullLoadCurrent}/{sConstructionDate}/{sFLGrouping}/" +
                                                   "{sCheckInComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string UpdatePlantEquipItem(string sSessionId, string sUserId, string sFullName, string sPlantEquipNo, 
                                    string sName, string sDesc, string sLongDesc, string sContSysType,
                                    string sDriveRating, string sEquipRegFlag, string sIPRegFlag, string sIPAddress,
                                    string sComments, string sOpZone,
                                    string sPowerCable, string sControlCable, string sInstrumentationCable, 
                                    string sDataCable, string sEarthCable,
                                    string sInstRegFlag, string sFullLoadCurrent, string sConstructionDate, string sFLGrouping,
                                    string sCheckInComments, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "createworkpackageitem/{sSessionId}/{sUserId}/{sFullname}/{sWPNo}/{sPartType}/{sName}/{sDesc}/{sMaintenanceType}/{sTrigThreshold}/" +
                                                    "{sElapsedNextDate}/{sMonitoredPart}/{sAccumThreshold}/{sWarningAlert}/{sProduct}/{sFolder}/" +
                                                    "{sCheckInComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string CreateWorkPackageItem(string sSessionId, string sUserId, string sFullName, string sWPNo,
                                     string sPartType, string sName, string sDesc, string sMaintenanceType, string sTrigThreshold, string sElapsedNextDate,
                                     string sMonitoredPart, string sAccumThreshold, string sWarningAlert,
                                     string sProduct, string sFolder, string sCheckInComments, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "updateworkpackageitem/{sSessionId}/{sUserId}/{sFullname}/{sWPNo}/{sName}/{sDesc}/{sMaintenanceType}/{sTrigThreshold}/" +
                                                    "{sElapsedNextDate}/{sMonitoredPart}/{sAccumThreshold}/{sWarningAlert}/" +
                                                    "{sCheckInComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string UpdateWorkPackageItem(string sSessionId, string sUserId, string sFullName, string sWPNo,
                                     string sName, string sDesc, string sMaintenanceType, string sTrigThreshold, string sElapsedNextDate,
                                     string sMonitoredPart, string sAccumThreshold, string sWarningAlert,
                                     string sCheckInComments, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "createrequiredactionitem/{sSessionId}/{sUserId}/{sFullname}/{sReqdActionNo}/{sPlantEquipType}/{sName}/{sDesc}/{sComments}/{sCompletionStatus}/{sCompletionDate}/{sProduct}/{sFolder}/" +
                                                       "{sCheckInComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string CreateRequiredActionItem(string sSessionId, string sUserId, string sFullName, string sReqdActionNo,
                                        string sPlantEquipType, string sName, string sDesc,
                                        string sComments, string sCompletionStatus, string sCompletionDate,
                                        string sProduct, string sFolder,
                                        string sCheckInComments, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "updaterequiredactionitem/{sSessionId}/{sUserId}/{sFullname}/{sReqdActionNo}/{sName}/{sDesc}/{sComments}/{sCompletionStatus}/" +
                                                   "{sCompletionDate}/{sCheckInComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string UpdateRequiredActionItem(string sSessionId, string sUserId, string sFullName, string sReqdActionNo,
                                        string sName, string sDesc, string sComments, string sCompletionStatus,
                                        string sCompletionDate,
                                        string sCheckInComments, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "createbatchactionitem/{sSessionId}/{sUserId}/{sFullname}/{sBatchNo}/{sBatchActionNo}/{sDesc}/{sVerify}/{sCompletionStatus}/{sCompletionDate}/{sCompletedBy}/" +
                                                    "{sActionedDate}/{sActionedBy}/{sProduct}/{sFolder}/{sCheckInComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string CreateBatchActionItem(string sSessionId, string sUserId, string sFullName, string sBatchNo,
                                            string sBatchActionNo, string sDesc, string sVerify, string sCompletionStatus,
                                            string sCompletionDate, string sCompletedBy, string sActionedDate, string sActionedBy,
                                            string sProduct, string sFolder,
                                            string sCheckInComments, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "updatebatchactionitem/{sSessionId}/{sUserId}/{sFullname}/{sBatchNo}/{sBatchActionNo}/{sDesc}/{sVerify}/{sCompletionStatus}/{sCompletionDate}/{sCompletedBy}/" +
                                                    "{sActionedDate}/{sActionedBy}/{sCheckInComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string UpdateBatchActionItem(string sSessionId, string sUserId, string sFullName, string sBatchNo,
                                               string sBatchActionNo, string sDesc, string sVerify, string sCompletionStatus,
                                               string sCompletionDate, string sCompletedBy, string sActionedDate, string sActionedBy,
                                               string sCheckInComments, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "createorganisationitem/{sSessionId}/{sUserId}/{sFullname}/{sOrganisationNo}/{sDocType}/{sName}/{sDesc}/{sEmail}/{sProduct}/{sFolder}/" +
                                                       "{sCheckInComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string CreateOrganisationItem(string sSessionId, string sUserId, string sFullName, string sOrganisationNo,
                                        string sDocType, string sName, string sDesc, string sEmail,
                                        string sProduct, string sFolder,
                                        string sCheckInComments, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "updateorganisationitem/{sSessionId}/{sUserId}/{sFullname}/{sOrganisationNo}/{sName}/{sDesc}/{sEmail}/{sCheckInComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string UpdateOrganisationItem(string sSessionId, string sUserId, string sFullName, string sOrganisationNo,
                                        string sName, string sDesc, string sEmail, string sCheckInComments, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "createpersonitem/{sSessionId}/{sUserId}/{sFullname}/{sPersonNo}/{sDocType}/{sName}/{sDesc}/{sEmail}/{sProduct}/{sFolder}/" +
                                                       "{sCheckInComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string CreatePersonItem(string sSessionId, string sUserId, string sFullName, string sPersonNo,
                                        string sDocType, string sName, string sDesc, string sEmail,
                                        string sProduct, string sFolder,
                                        string sCheckInComments, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "updatepersonitem/{sSessionId}/{sUserId}/{sFullname}/{sPersonNo}/{sName}/{sDesc}/{sEmail}/{sCheckInComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string UpdatePersonItem(string sSessionId, string sUserId, string sFullName, string sPersonNo,
                                        string sName, string sDesc, string sEmail, string sCheckInComments, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "setmaintenancetemplates/{sSessionId}/{sUserId}/{sWONo}/{sWOName}/{sTemplateIndex}/{sWPNo}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string SetMaintenanceTemplates(string sSessionId, string sUserId, string sWONo, string sWOName, string sTemplateIndex, string sWPNo, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "processiospreadsheet/{sSessionId}/{sUserId}/{sFile}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string ProcessIOSpreadsheet(string sSessionId, string sUserId, string sFile, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "processiopreallocationspreadsheet/{sSessionId}/{sUserId}/{sFile}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string ProcessIOPreallocatedSpreadsheet(string sSessionId, string sUserId, string sFile, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "processcablespreadsheet/{sSessionId}/{sUserId}/{sFile}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string ProcessCableSpreadsheet(string sSessionId, string sUserId, string sFile, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "processterminationspreadsheet/{sSessionId}/{sUserId}/{sFile}/{sWebAppId}/{sFLOrMat}", ResponseFormat = WebMessageFormat.Xml)]
        string ProcessTerminationSpreadsheet(string sSessionId, string sUserId, string sFile, string sWebAppId, string sFLOrMat);

        [OperationContract]
        [WebGet(UriTemplate = "emailmessage/{sSessionId}/{sUserId}/{sSubject}/{sBody}/{sAttachments}/{sRecipients}/{sCCRecipients}/{sBCCRecipients}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string emailmessage(string sSessionId, string sUserId, string sSubject, string sBody, string sAttachments, string sRecipients, string sCCRecipients, string sBCCRecipients, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "setfuncdoctopartref/{sSessionId}/{sUserId}/{sFullname}/{sFuncDocNo}/{sPartNo}/{sSequenceNo}/{sPrimaryPart}/{sPartDocRefLinkType}/{sCheckinComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string SetFuncDocToPartRef(string sSessionId, string sUserId, string sFullName, string sFuncDocNo, string sPartNo, string sSequenceNo, string sPrimaryPart, string sPartDocRefLinkType, string sCheckinComments, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "updatefuncdoctopartref/{sSessionId}/{sUserId}/{sFullname}/{sFuncDocNo}/{sPartNo}/{sSequenceNo}/{sPrimaryPart}/{sCheckinComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string UpdateFuncDocToPartRef(string sSessionId, string sUserId, string sFullName, string sFuncDocNo, string sPartNo, string sSequenceNo, string sPrimaryPart, string sCheckinComments, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "setsuppliertopartref/{sSessionId}/{sUserId}/{sFullname}/{sSupplierNo}/{sPartNo}/{sSupplierPartNo}/{sPartDocRefLinkType}/{sCheckinComments}/{sWebAppId}/{sManufacturerFlag}", ResponseFormat = WebMessageFormat.Xml)]
        string SetSupplierToPartRef(string sSessionId, string sUserId, string sFullName, string sSupplierNo, string sPartNo, string sSupplierPartNo, string sPartDocRefLinkType, string sCheckinComments, string sWebAppId, string sManufacturerFlag);

        [OperationContract]
        [WebGet(UriTemplate = "updatesuppliertopartref/{sSessionId}/{sUserId}/{sFullname}/{sSupplierNo}/{sPartNo}/{sSupplierPartNo}/{sCheckinComments}/{sWebAppId}/{sManufacturerFlag}", ResponseFormat = WebMessageFormat.Xml)]
        string UpdateSupplierToPartRef(string sSessionId, string sUserId, string sFullName, string sSupplierNo, string sPartNo, string sSupplierPartNo, string sCheckinComments, string sWebAppId, string sManufacturerFlag);

        [OperationContract]
        [WebGet(UriTemplate = "revisedocumentandremoveattachments/{sSessionId}/{sUserId}/{sFullname}/{sDocNo}/{sDocName}/{sRevision}/{sLongDesc}/{sOriginator}/" +
                              "{sOriginatorDocId}/{sJobCode}/{sCheckinComments}/{sIncludeHyperlinks}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string ReviseDocumentAndRemoveAttachments(string sSessionId, string sUserId, string sFullname, string sDocNo, string sDocName, string sRevision,
                                                         string sLongDesc, string sOriginator, string sOriginatorDocId, string sJobCode,
                                                         string sCheckInComments, string sIncludeHyperlinks, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "setmaterialioparttopartlink/{sSessionId}/{sUserId}/{sFullname}/{sParentPartNo}/{sChildPartNo}/{sLineNumber}/{sIOType}/{sIOTag}/{sCheckinComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string SetMaterialIOPartToPartLink(string sSessionId, string sUserId, string sFullName, string sParentPartNo, string sChildPartNo, string sLineNumber, string sIOType, string sIOTag, string sCheckinComments, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "updateioparttopartlink/{sSessionId}/{sUserId}/{sFullname}/{sParentPartNo}/{sChildPartNo}/{sLineNumber}/{sIOType}/{sIOTag}/{sCheckinComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string UpdateIOPartToPartLink(string sSessionId, string sUserId, string sFullName, string sParentPartNo, string sChildPartNo, string sLineNumber, string sIOType, string sIOTag, string sCheckinComments, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "removeownaccreditationitem/{sSessionId}/{sUserId}/{sFullName}/{sOrgOrPersonNo}/{sOwnAccreditationNo}/{sWorkflowId}/{sCheckinComments}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string RemoveOwnAccreditationItem(string sSessionId, string sUserId, string sFullName,  string sOrgOrPersonNo, string sOwnAccreditationNo, string sWorkflowId, string sCheckinComments, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "reassignpartlifecycle/{sSessionId}/{sUserId}/{sPartNo}/{sLifecycleName}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string ReassignPartLifecycle(string sSessionId, string sUserId, string sPartNo, string sLifecycleName, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "reassigndocumentlifecycle/{sSessionId}/{sUserId}/{sDocumentNo}/{sLifecycleName}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string ReassignDocumentLifecycle(string sSessionId, string sUserId, string sDocumentNo, string sLifecycleName, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "processfixactionsupportlink/{sSessionId}/{sUserId}/{sFile}/{sWebAppId}", ResponseFormat = WebMessageFormat.Xml)]
        string ProcessFixActionSupportLink(string sSessionId, string sUserId, string sFile, string sWebAppId);

        [OperationContract]
        [WebGet(UriTemplate = "processbulkupdatedocumentlifecycle/{sSessionId}/{sUserId}/{sFile}/{sWebAppId}/{sLatestOrHistory}", ResponseFormat = WebMessageFormat.Xml)]
        string ProcessBulkUpdateDocumentLifecycle(string sSessionId, string sUserId, string sFile, string sWebAppId, string sLatestOrHistory);
    }


}
