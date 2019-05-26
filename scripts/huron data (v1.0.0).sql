
/*
*********************************************************************************
*
*	Insert default languages(s)
*
*********************************************************************************
*/

INSERT INTO tblLanguages (LangName, Localized, IsRTL, Encoding, ISO639, IsActive)
	VALUES ('English', 'English', 0, 'UTF-8', 'EN', 1)
GO


/*
*********************************************************************************
*
*	Insert default role(s)
*
*********************************************************************************
*/

INSERT INTO tblRoles (RoleName, RoleDesc, RoleMask, IsActive)
	VALUES ('Administrator', 'Liberum Support Desk Administrators', 120026087, 1)
GO
INSERT INTO tblRoles (RoleName, RoleDesc, RoleMask, IsActive)
	VALUES ('Technician', 'Liberum Support Desk Technicians', 119676898, 1)
GO
INSERT INTO tblRoles (RoleName, RoleDesc, RoleMask, IsActive)
	VALUES ('User', 'Liberum Support Desk Users', 17039585, 1)
GO


/*
*********************************************************************************
*
*	Insert default lists
*
*********************************************************************************
*/


/*
*	Insert Status List Items
*/

INSERT INTO tblLists (ItemOrder, ItemName, IsActive)
	VALUES (5, 'STATUS_LIST', 1)
GO


DECLARE @LIST_ID Int
SET @LIST_ID = (SELECT ListItemPK FROM tblLists WHERE ItemName='STATUS_LIST')

INSERT INTO tblLists (ParentListItemFK, ItemOrder, ItemName, IsActive)
	VALUES (@LIST_ID, 0, 'Open', 1)

INSERT INTO tblLists (ParentListItemFK, ItemOrder, ItemName, IsActive)
	VALUES (@LIST_ID, 25, 'Pending', 1)

INSERT INTO tblLists (ParentListItemFK, ItemOrder, ItemName, IsActive)
	VALUES (@LIST_ID, 75, 'Cancelled', 1)

INSERT INTO tblLists (ParentListItemFK, ItemOrder, ItemName, IsActive)
	VALUES (@LIST_ID, 100, 'Closed', 1)
GO

/*
*	Insert Priority List Items
*/

INSERT INTO tblLists (ItemOrder, ItemName, IsActive)
	VALUES (5, 'PRIORITY_LIST', 1)
GO

DECLARE @LIST_ID Int
SET @LIST_ID = (SELECT ListItemPK FROM tblLists WHERE ItemName='PRIORITY_LIST')

INSERT INTO tblLists (ParentListItemFK, ItemOrder, ItemName, IsActive)
	VALUES (@LIST_ID, 0, 'Low', 1)

INSERT INTO tblLists (ParentListItemFK, ItemOrder, ItemName, IsActive)
	VALUES (@LIST_ID, 25, 'Normal', 1)

INSERT INTO tblLists (ParentListItemFK, ItemOrder, ItemName, IsActive)
	VALUES (@LIST_ID, 50, 'High', 1)

INSERT INTO tblLists (ParentListItemFK, ItemOrder, ItemName, IsActive)
	VALUES (@LIST_ID, 75, 'Urgent', 1)
GO

/*
*	Insert Contact Type Items
*/

INSERT INTO tblLists (ItemOrder, ItemName, IsActive)
	VALUES (5, 'CONTACT_TYPE_LIST', 1)
GO

DECLARE @LIST_ID Int
SET @LIST_ID = (SELECT ListItemPK FROM tblLists WHERE ItemName='CONTACT_TYPE_LIST')

INSERT INTO tblLists (ParentListItemFK, ItemOrder, ItemName, IsActive)
	VALUES (@LIST_ID, 0, 'Contractor', 1)

INSERT INTO tblLists (ParentListItemFK, ItemOrder, ItemName, IsActive)
	VALUES (@LIST_ID, 50, 'Staff', 1)
GO

/*
*	Insert Organisation Type Items
*/

INSERT INTO tblLists (ItemOrder, ItemName, IsActive)
	VALUES (5, 'ORG_TYPE_LIST', 1)
GO

DECLARE @LIST_ID Int
SET @LIST_ID = (SELECT ListItemPK FROM tblLists WHERE ItemName='ORG_TYPE_LIST')

INSERT INTO tblLists (ParentListItemFK, ItemOrder, ItemName, IsActive)
	VALUES (@LIST_ID, 25, 'IT', 1)

GO

/*
*********************************************************************************
*
*	Insert default Organisation(s)
*
*********************************************************************************
*/

INSERT INTO tblOrganisations (OrgTypeFK, OrgShortName, OrgName, IsActive)
	VALUES ( 1, 'Liberum', 'Liberum', 1)
GO


/*
*********************************************************************************
*
*	Insert default Department(s)
*
*********************************************************************************
*/

INSERT INTO tblDepartments (OrgFk, DeptName, DeptDesc, IsActive)
	VALUES (1, 'IT Support', 'IT Support', 1)
GO


/*
*********************************************************************************
*
*	Insert default user(s)
*
*********************************************************************************
*/

DECLARE @ROLE_ID Int
SET @ROLE_ID = (SELECT RolePK FROM tblRoles WHERE RoleName='Administrator')

INSERT INTO tblContacts (UserName, PW, FName, LName, DeptFK, LangFK, RoleFK, IsActive)
	VALUES ('Admin', 'password', 'Liberum', 'Administrator', 1, 1, @ROLE_ID, 1)
GO



/*
*********************************************************************************
*
*	Insert default Message Types
*
*********************************************************************************
*/

INSERT INTO tblEMailMsgs (EMailMsgType, Subject, Body, LangFK, IsActive)
	VALUES ('CASE_ASSIGNED_REP', 'Case #[CASEID] Assigned', 'Case #[CASEID] has been assigned to you for action.  Please action at your earliest convienance. ', 1, 1)
GO
INSERT INTO tblEMailMsgs (EMailMsgType, Subject, Body, LangFK, IsActive)
	VALUES ('CASE_CANCELLED_USER', 'Case #[CASEID] Cancelled', 'Case #[CASEID] has been cancelled.', 1, 1)
GO
INSERT INTO tblEMailMsgs (EMailMsgType, Subject, Body, LangFK, IsActive)
	VALUES ('CASE_CLOSED_USER', 'Case #[CASEID] Closed', 'Case #[CASEID] has been closed.', 1, 1)
GO
INSERT INTO tblEMailMsgs (EMailMsgType, Subject, Body, LangFK, IsActive)
	VALUES ('CASE_REASSIGNED_REP', 'Case #[CASEID] Reassigned', 'Case #[CASEID] has been reassigned to you for action.  Please action at your earliest convienance.', 1, 1)
GO
INSERT INTO tblEMailMsgs (EMailMsgType, Subject, Body, LangFK, IsActive)
	VALUES ('CASE_REOPENED_USER', 'Case #[CASEID] Reopened', 'Case #[CASEID] has been re-opened.', 1, 1)
GO
INSERT INTO tblEMailMsgs (EMailMsgType, Subject, Body, LangFK, IsActive)
	VALUES ('CASE_SUBMITTED_REP', 'Case #[CASEID] Submitted', 'Case #[CASEID] has been re-opened.', 1, 1)
GO
INSERT INTO tblEMailMsgs (EMailMsgType, Subject, Body, LangFK, IsActive)
	VALUES ('CASE_SUBMITTED_USER', 'Case #[CASEID] Submitted', 'Thank you, your case had been received and will be actioned as soon as possible.  The progress of your case can be monitored via the Main Menu or by clicking on the link provided below. [URL]', 1, 1)
GO
INSERT INTO tblEMailMsgs (EMailMsgType, Subject, Body, LangFK, IsActive)
	VALUES ('CASE_UPDATED_REP', 'Case #[CASEID] Updated', 'Case #[CASEID] has been updated by the requestor of this case.  Please review these updates at your earliest convienance. [URL]', 1, 1)
GO
INSERT INTO tblEMailMsgs (EMailMsgType, Subject, Body, LangFK, IsActive)
	VALUES ('CASE_UPDATED_USER', 'Case #[CASEID] Updated', 'Case #[CASEID] has been updated.  You can view these updates by clicking the link below. [URL]', 1, 1)
GO
INSERT INTO tblEMailMsgs (EMailMsgType, Subject, Body, LangFK, IsActive)
	VALUES ('CASE_UNASSIGNED_REP', 'Case #[CASEID] Unassigned', '', 1, 1)
GO


/*
*********************************************************************************
*
*	Insert default Groups
*
*********************************************************************************
*/

INSERT INTO tblGroups (GroupName, GroupDesc, IsActive)
	VALUES ('IT Support Reps', 'IT Support Reps', 1)
GO


/*
*********************************************************************************
*
*	Insert default Case Types
*
*********************************************************************************
*/

INSERT INTO tblCaseTypes (CaseTypeName, CaseTypeDesc, CaseTypeOrder, RepGroupFK, IsActive)
	VALUES ('IT Support', 'IT Support', 5, 1, 1)
GO


/*
*********************************************************************************
*
*	Insert default Catagories
*
*********************************************************************************
*/



/*
*********************************************************************************
*
*	Insert default parameters
*
*********************************************************************************
*/

INSERT INTO tblParameters (ParamName, ParamValue)
	VALUES ('AUTH_TYPE', '2')
GO
INSERT INTO tblParameters (ParamName, ParamValue)
	VALUES ('BASE_URL', 'http://')
GO
INSERT INTO tblParameters (ParamName, ParamValue)
	VALUES ('DATE_FORMAT', 'dd-mmm-yyyy')
GO
INSERT INTO tblParameters (ParamName, ParamValue)
	VALUES ('DEFAULT_LANGUAGE', 1)
GO
DECLARE @PRIORITY_ID Int
SET @PRIORITY_ID = (SELECT ListItemPK FROM tblLists WHERE ItemName='Normal' AND ParentListItemFK=(SELECT ListItemPK FROM tblLists WHERE ItemName='PRIORITY_LIST'))
INSERT INTO tblParameters (ParamName, ParamValue)
	VALUES ('DEFAULT_PRIORITY', @PRIORITY_ID)
GO
DECLARE @ROLE_ID Int
SET @ROLE_ID = (SELECT RolePK FROM tblRoles WHERE RoleName='User')
INSERT INTO tblParameters (ParamName, ParamValue)
	VALUES ('DEFAULT_ROLE', @ROLE_ID)
GO
DECLARE @STATUS_ID Int
SET @STATUS_ID = (SELECT ListItemPK FROM tblLists WHERE ItemName='Open' AND ParentListItemFK=(SELECT ListItemPK FROM tblLists WHERE ItemName='STATUS_LIST'))
INSERT INTO tblParameters (ParamName, ParamValue)
	VALUES ('DEFAULT_STATUS', @STATUS_ID)
GO
INSERT INTO tblParameters (ParamName, ParamValue)
	VALUES ('ENABLE_ATTACHMENTS', '1')
GO
INSERT INTO tblParameters (ParamName, ParamValue)
	VALUES ('ENABLE_EMAIL', '0')
GO
INSERT INTO tblParameters (ParamName, ParamValue)
	VALUES ('ENABLE_INOUT', '0')
GO
INSERT INTO tblParameters (ParamName, ParamValue)
	VALUES ('ENABLE_KB', '1')
GO
INSERT INTO tblParameters (ParamName, ParamValue)
	VALUES ('ENABLE_REPORTS', '1')
GO
INSERT INTO tblParameters (ParamName, ParamValue)
	VALUES ('EMAIL_METHOD', '2')
GO
INSERT INTO tblParameters (ParamName, ParamValue)
	VALUES ('ITEMS_PER_PAGE', '25')
GO
INSERT INTO tblParameters (ParamName, ParamValue)
	VALUES ('MAX_ATTACHMENT_SIZE', '1024')
GO
INSERT INTO tblParameters (ParamName, ParamValue)
	VALUES ('SITE_NAME', 'Liberum Help desk')
GO
INSERT INTO tblParameters (ParamName, ParamValue)
	VALUES ('SMTP_SERVER', 'smtp.domain.com')
GO
DECLARE @STATUS_ID Int
SET @STATUS_ID = (SELECT ListItemPK FROM tblLists WHERE ItemName='Cancelled' AND ParentListItemFK=(SELECT ListItemPK FROM tblLists WHERE ItemName='STATUS_LIST'))
INSERT INTO tblParameters (ParamName, ParamValue)
	VALUES ('STATUS_CANCELLED', @STATUS_ID)
GO
DECLARE @STATUS_ID Int
SET @STATUS_ID = (SELECT ListItemPK FROM tblLists WHERE ItemName='Closed' AND ParentListItemFK=(SELECT ListItemPK FROM tblLists WHERE ItemName='STATUS_LIST'))
INSERT INTO tblParameters (ParamName, ParamValue)
	VALUES ('STATUS_CLOSED', @STATUS_ID)
GO
DECLARE @STATUS_ID Int
SET @STATUS_ID = (SELECT ListItemPK FROM tblLists WHERE ItemName='Open' AND ParentListItemFK=(SELECT ListItemPK FROM tblLists WHERE ItemName='STATUS_LIST'))
INSERT INTO tblParameters (ParamName, ParamValue)
	VALUES ('STATUS_OPEN', @STATUS_ID)
GO
INSERT INTO tblParameters (ParamName, ParamValue)
	VALUES ('SYSTEM_EMAIL', 'Liberum@domain.com')
GO
INSERT INTO tblParameters (ParamName, ParamValue)
	VALUES ('TIME_FORMAT', 'hh:mm')
GO
INSERT INTO tblParameters (ParamName, ParamValue)
	VALUES ('VERSION', '1.0.0.b1')
GO


/*
*********************************************************************************
*
*	Insert default permissions
*
*********************************************************************************
*/

/*
*	1st Binary Bit
*/

INSERT INTO tblPermissions (PermByte, PermLabel, PermDesc)
	VALUES (1, 'PERM_VIEW_USER', 'View User Fields')
GO
INSERT INTO tblPermissions (PermByte, PermLabel, PermDesc)
	VALUES (2, 'PERM_VIEW_TECH', 'View Technician Fields')
GO
INSERT INTO tblPermissions (PermByte, PermLabel, PermDesc)
	VALUES (4, 'PERM_VIEW_ADMIN', 'View Admin Fields')
GO

/*
*	2nd Binary Bit
*/

INSERT INTO tblPermissions (PermByte, PermLabel, PermDesc)
	VALUES (32, 'PERM_CREATE_OWN', 'Create Rights for Own Cases')
GO
INSERT INTO tblPermissions (PermByte, PermLabel, PermDesc)
	VALUES (64, 'PERM_READ_OWN', 'Read Rights for Own Cases')
GO
INSERT INTO tblPermissions (PermByte, PermLabel, PermDesc)
	VALUES (128, 'PERM_MODIFY_OWN', 'Modify Rights for Own Cases')
GO

/*
*	3rd Binary Bit
*/

INSERT INTO tblPermissions (PermByte, PermLabel, PermDesc)
	VALUES (256, 'PERM_READ_GROUP', 'Read Rights for Group Cases')
GO
INSERT INTO tblPermissions (PermByte, PermLabel, PermDesc)
	VALUES (512, 'PERM_MODIFY_GROUP', 'Modify Rights for Group Cases')
GO
INSERT INTO tblPermissions (PermByte, PermLabel, PermDesc)
	VALUES (1024, 'PERM_READ_ASSIGNED', 'Read Rights for Cases Assigned to the Contact')
GO
INSERT INTO tblPermissions (PermByte, PermLabel, PermDesc)
	VALUES (2048, 'PERM_MODIFY_ASSIGNED', 'Modify Rights for Cases Assigned to the Contact')
GO

/*
*	4th Binary Bit
*/

INSERT INTO tblPermissions (PermByte, PermLabel, PermDesc)
	VALUES (4096, 'PERM_CREATE_ALL','Create Rights on Behalf of Others')
GO
INSERT INTO tblPermissions (PermByte, PermLabel, PermDesc)
	VALUES (8192, 'PERM_READ_ALL', 'Read Rights for All Cases')
GO
INSERT INTO tblPermissions (PermByte, PermLabel, PermDesc)
	VALUES (16384, 'PERM_MODIFY_ALL', 'Modify Rights for All Cases')
GO

/*
*	5th Binary Bit
*/

INSERT INTO tblPermissions (PermByte, PermLabel, PermDesc)
	VALUES (65536, 'PERM_ACCESS_ADMIN', 'Manage & Configure System')
GO
INSERT INTO tblPermissions (PermByte, PermLabel, PermDesc)
	VALUES (131072, 'PERM_ACCESS_TECH', 'Technician Access')
GO
INSERT INTO tblPermissions (PermByte, PermLabel, PermDesc)
	VALUES (262144, 'APERM_CCESS_USER', 'User Access')
GO

/*
*	6th Binary Bit
*/

INSERT INTO tblPermissions (PermByte, PermLabel, PermDesc)
	VALUES (1048576, 'PERM_REOPEN_CASES', 'Ability to Re-open Cases')
GO
INSERT INTO tblPermissions (PermByte, PermLabel, PermDesc)
	VALUES (2097152, 'PERM_ACCESS_REPORTS', 'Access Reports Module')
GO

/*
*	7th Binary Bit
*/

INSERT INTO tblPermissions (PermByte, PermLabel, PermDesc)
	VALUES (16777216, 'PERM_KB_READ', 'Read Knowledgebase')
GO
INSERT INTO tblPermissions (PermByte, PermLabel, PermDesc)
	VALUES (33554432, 'PERM_KB_MODIFY', 'Modify Knowledgebase')
GO
INSERT INTO tblPermissions (PermByte, PermLabel, PermDesc)
	VALUES (67108864, 'PERM_KB_CREATE', 'Add/Create Knowledgebase Records')
GO


/*
*********************************************************************************
*
*	Insert Languages Labels
*
*********************************************************************************
*/

INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Actions')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Active')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Add')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Add_To_Knowledgebase')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Add-Ons')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Administration')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Administration_Menu')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('All')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Alternate_Email')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Assigned')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Assigned_Rep')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Assignment')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Assignment_Saved')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Assignments')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Attach')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Attachments_Enabled')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Attach_Files')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Authentication')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Authentication_Type')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Between_Dates')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Body')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Cancelled')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Cancelled_Status')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Case')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Case_Criteria')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Case_Detail')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Case_Options')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Case_Submitted')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Case_Type')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Case_Type_Saved')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Case_Types')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Case_Updated')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Categories')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Category')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Category_Saved')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Cause')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Cc')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('City')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Close')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Closed')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Closed_Date')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Closed_Status')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Contact')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Contact_Detail')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Contact_Details')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Contact_Saved')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Contact_Type')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Contacts')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Country')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Courier_Address')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Custom')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Date_Criteria')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Date_Format')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Default_Language')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Default_Priority')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Default_Role')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Default_Status')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Department')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Department_Saved')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Departments')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Description')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Detailed_Description')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Edit')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Edit_My_Profile')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Email')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Email_Enabled')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Email_Message')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Email_Message_Saved')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Email_Message_Type')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Email_Messages')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Email_Type')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Encoding')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Entered_By')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Entered_Date')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Fax')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('First_Name')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('From')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Full_Name')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Group')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Group_Name')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Group_Saved')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Groups')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('In/Out_Status')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('In/Out_Status_Date')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('In/Out_Status_Text')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Is_Active')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Is_RTL')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Is_User')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('ISO_639')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Issue')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Item_Name')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Item_Order')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Items')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Items_Per_Page')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Job_Function')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Job_Title')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Keyword')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Keyword_Criteria')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Keywords')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Knowledgebase')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Knowledgebase_Enabled')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Knowledgebase_Record')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Language')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Language_Name')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Languages')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Last_Updated')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Last_Access_Time')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Last_Name')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Last_Updated_By')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('List')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('List_All_Un-Assigned_Cases')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('List_Assigned_Cases_For')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('List_Item')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('List_Item_Saved')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('List_Items')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('List_My_Active_Cases')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('List_My_Assigned_Cases')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('List_My_Cases')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('List_Name')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('List_Saved')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('List_Unassigned_Cases')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Lists')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Localised')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Location')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Log_Off')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Log_On')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Logon')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Mail_Address')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Main_Menu')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Manage')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Manage_&_Configure_System')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Manage_Assignments')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Manage_Case_Types')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Manage_Categories')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Manage_Contacts')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Manage_Departments')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Manage_Email_Messages')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Manage_Groups')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Manage_Languages')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Manage_List_Items')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Manage_Lists')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Manage_Organisations')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Manage_Roles')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Manage_Strings')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Max_Attachment_Size')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Members')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Modify')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Modify_Assignment')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Modify_Case_Type')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Modify_Category')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Modify_Contact')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Modify_Department')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Modify_Email')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Modify_Email_Message')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Modify_Group')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Modify_Language')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Modify_List')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Name')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('New')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('New_Assignment')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('New_Case')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('New_Category')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('New_Contact')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('New_Department')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('New_Email')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('New_Email_Message')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('New_Knowledgebase_Record')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('New_User')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Next')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('No')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Non_Members')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Note')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Notes')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Notification_Settings')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Notify_User')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Office_Address')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('On_Behalf_Of')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Open')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Open_Status')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Options')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Or')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Order')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Order_By')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Organisation')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Organisation_Name')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Organisation_Short_Name')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Organisation_Type')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Organisations')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Pager_Email')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Password')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Pending')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Permissions')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Phone')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Phone_Home')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Phone_Mobile')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Phone_Work')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Photo_File')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Photo_File_ID')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Previous')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Primary_Contact')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Printer_Friendly_View')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Priority')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Private_Note')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Raised_Date')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Reference_ID')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Register')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Remove')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Reopen')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Rep')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Rep_Group')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Rep_Options')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Report_Criteria')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Reports')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Reports_Enabled')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Reports_Menu')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Resolution')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Results')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Resume')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Role')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Role_Name')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Roles')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Save')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Saved')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Search')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Search_Case_Archives')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Search_Cases')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Search_Knowledgebase')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Search_Options')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Search_Results')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Site_Information')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Site_Name')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('SMTP_Server')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Start_Date')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('State')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Status')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Status_Summary_Report')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Subject')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Submit')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Submit_New_Case')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('System_Configuration')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('System_Configuration_Saved')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('System_Defaults')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('System_Email_Address')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Through')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Time_Format')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Time_Spent')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Timezone_Offset')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Title')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('To')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Type')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Update')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('User_Name')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('User_Permissions')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Version')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('View')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('View/Edit_Attachments')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('View_Case')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('View_Contact_Details')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('View_License')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('View_Reports')
INSERT INTO tblLanguageLabels (LangLabel) VALUES ('Yes')

GO
                                                     