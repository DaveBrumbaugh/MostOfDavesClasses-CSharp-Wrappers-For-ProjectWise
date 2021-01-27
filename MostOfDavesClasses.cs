using System;
using System.Collections.Generic;
using System.Text;
using System.Collections;
using System.Configuration;
using System.IO;
using System.Data;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Threading;
using System.Xml;
//using System.Net.Http;
//using System.Net.Http.Headers;
using System.Linq;
using System.Globalization;

public class PWWrapper
{
    static PWWrapper()
    {
        try
        {
            aaApi_Initialize(512); // trying to avoid loading WRE
        }
        catch (DllNotFoundException dlfne)
        {
            string sMessage = dlfne.Message;
            System.Diagnostics.Debug.WriteLine("Fixing path...");
            AppendProjectWiseDllPathToEnvironmentPath();
        }
        catch (BadImageFormatException badImage)
        {
            string sMessage = badImage.Message;
            System.Diagnostics.Debug.WriteLine("Fixing path...");
            AppendProjectWiseDllPathToEnvironmentPath();
        }
        finally
        {
            aaApi_Initialize(512); // trying to avoid loading WRE
        }
    }

    public enum ProjectTypes : int
    {
        AADMS_PROJECT_TYPE_NORMAL = 0,     /**< Specifies a normal project */
        AADMS_PROJECT_TYPE_RICH = 2      /**< Specifies a rich project that may contain ODS properties */
    }

    public enum WorkflowTypes : int
    {
        AADMS_WORKFLOW_PROJECT = 1,
        AADMS_WORKFLOW_DOCUMENT = 2,
        AADMS_WORKFLOW_BOTH = 3
    }

    // dww - to support aaApi_ChangeDocumentFile4

    [Flags]
    public enum DocumentFileOp : uint
    {
        AADMS_DOCUMENT_FILE_COMMAND = 0x0000000f,
        AADMS_DOCUMENT_FILE_REPLACE = 0x0,
        AADMS_DOCUMENT_FILE_RENAME = 0X1,
        AADMS_DOCUMENT_FILE_ADD = 0x2,
        AADMS_DOCUMENT_FILE_KEEP_ITYPE = 0x00000010
    }

    public enum AccessFolderRights : int
    {
        SECURITYOBJECT = 1,
        FULLCONTROL = 2,
        CHANGEPERMISSIONS = 3,
        CREATESUBFOLDER = 4,
        DELETE = 5,
        READ = 6,
        WRITE = 7,
        NOACCESS = 8,
        FOLDER_ACCESS = 1,
        DOCUMENT_ACCESS = 2
    }

    public enum AccessDocumentRights : int
    {
        SECURITYOBJECT = 2,
        DOCFULLCONTROL = 3,
        DOCCHANGEPERMISSIONS = 4,
        DOCDELETE = 5,
        DOCREAD = 7,
        DOCWRITE = 6,
        DOCFILEREAD = 8,
        DOCFILEWRITE = 9,
        DOCNOACCESS = 10,
    }

    public enum DepartmentProperty : int
    {
        ID = 1,
        Name = 2,
        Desc = 3,
        DisplayName = 4,
    }

    public enum ViewProperties
    {
        VIEW_PROP_ID = 1,
        VIEW_PROP_VIEWTYPE = 2,
        VIEW_PROP_USERID = 3,
        VIEW_PROP_NAME = 4,
        VIEW_PROP_CONTEXT = 5,
        VIEW_PROP_IS_GLOBAL = 6
    }

    public enum ViewColumnProperties
    {
        VIEWCOLS_PROP_ID = 1,
        VIEWCOLS_PROP_VIEWID = 2,
        VIEWCOLS_PROP_USERID = 3,
        VIEWCOLS_PROP_FIELDTYPE = 4,
        VIEWCOLS_PROP_CONTEXT = 5,
        VIEWCOLS_PROP_DATATYPE = 6,
        VIEWCOLS_PROP_FIELDNAME = 7,
        VIEWCOLS_PROP_ORDERNO = 8,
        VIEWCOLS_PROP_ALIGNMENT = 9,
        VIEWCOLS_PROP_WIDTH = 10,
        VIEWCOLS_PROP_SORTORDER = 11,
        VIEWCOLS_PROP_SORTDIR = 12,
        VIEWCOLS_PROP_DISPLAY_NAME = 13
    }

    public enum DocumentProperty : int
    {
        ID = 1,
        VersionNumber = 2,
        ProposalNumber = 3,
        CreatorID = 4,
        UpdaterID = 5,
        UserID = 6,
        Size = 7,
        FileType = 8,
        ItemType = 9,
        StorageID = 10,

        SetID = 11,
        SetType = 12,
        WorkFlowID = 13,
        StateID = 14,
        ApplicationID = 15,
        DepartmentID = 16,
        OriginalNumber = 18,
        IsOutToMe = 19,

        Name = 20,
        FileName = 21,
        Desc = 22,
        Version = 23,
        CreateTime = 24,
        UpdateTime = 25,
        DMSStatus = 26,
        DMSDate = 27,
        Node = 28,

        ProjectID = 29,
        Access = 30,
        IsLogicalSetMaster = 31,
        IsRedlineMaster = 32,
        IsRefMaster = 33,
        HasFinalStatus = 35,
        Manager = 36,
        FileUpdater = 37,
        LastRtLocker = 38,
        ItemFlags = 39,
        FileUpdateTime = 40,
        LastRtLockTime = 41,

        Is3DFile = 51, // 42, // old values
        Is2DFile = 52, // 43, // old values
        MgrType = 44,
        IsUrl = 45,
        UrlName = 46,
        DocGuid = 47,
        ProjGuid = 48,
        OrigGuid = 49,
        WSpaceProfID = 50,

        FileRevision = 53, // string
        Overlaps = 54, // numeric
        LocationID = 55, // guid
        MIMEType = 56, // string
        LocationSource = 58 // numeric
    }

    public enum ProjectProperty : int
    {
        ID = 1,
        VersionNo = 2,
        ManagerID = 3,
        StorageID = 4,
        CreatorID = 5,
        UpdaterID = 6,
        WorkflowID = 7,
        StateID = 8,
        Type = 9,
        ArchiveID = 10,
        IsParent = 11,

        Name = 12,
        Desc = 13,
        Code = 14,
        Version = 15,
        CreateTime = 16,
        UpdateTime = 17,
        Config = 18,
        Table = 19,

        EnvironmentID = 21,
        ParentID = 22,
        MgrType = 23,
        Access = 24,
        ProjGuid = 25,
        PprjGuid = 26,
        WSpaceProfID = 27,

        ComponentClassId = 28,
        Flags = 30,
        ComponentInstanceId = 31
    }

    public enum LinkDataSpecialColsIds : int
    {
        AADMSLINKDATACOL_PROJECTNO = -2,
        AADMSLINKDATACOL_DOCUMENTNO = -3,
        AADMSLINKDATACOL_UNIQUEVALUE = -4,
        AADMSLINKDATACOL_CREATORNO = -5,
        AADMSLINKDATACOL_CREATETIME = -6,
        AADMSLINKDATACOL_UPDATERNO = -7,
        AADMSLINKDATACOL_UPDATETIME = -8,
        AADMSLINKDATACOL_DOCGUID = -9
    }

    public enum ObjectTypeForLinkData : int
    {
        None = 0,
        DocumentByProject = 1,
        Document = 2,
        DocumentByWorkspace = 3,
        DocumentBySet = 4,
        DocumentByAttrRec = 5
    }


    [Flags]
    public enum LinkDataSelectFlags : uint
    {
        Creator = 0x00000001,
        CreateTime = 0x00000002,
        Updater = 0x00000004,
        UpdateTime = 0x00000008,
        SysColsOnly = 0x00000010,
        DocGuid = 0x00000020
    }

    //[Flags]
    //public enum AttributeParameterFlags : uint
    //{
    //    AADMS_ATTRDEF_FLD_UNIQUE = 0x00000001,
    //    AADMS_ATTRDEF_FLD_REQUIRED = 0x00000002,
    //    AADMS_ATTRDEF_FLD_EDITABLE_IF_FINAL = 0x00000010,
    //    AADMS_ATTRDEF_FLD_FORCE_UPDATE = 0x00000020,
    //      // Force call of update trigger (if set), regardless of the value changes. 
    //    AADMS_ATTRDEF_FLD_MULTISEL = 0x00000040,
    //      // Attribute value list is multi-selectable. 
    //    AADMS_ATTRDEF_FLD_LIMIT2LIST = 0x00000080,
    //      // Only the values from the list can be chosen. 
    //    AADMS_ATTRDEF_FLD_COPY_CLR_NEWSHEET = 0x00001000,
    //      // Clear attribute when new sheet created. 
    //    AADMS_ATTRDEF_FLD_COPY_CLR_INENV = 0x00002000,
    //      // Clear attribute when copying inside environment. 
    //    AADMS_ATTRDEF_FLD_COPY_CLR_OUTENV = 0x00004000,
    //      // Clear attribute when copying outside environment. 
    //    AADMS_ATTRDEF_FLD_COPY_CLR_OUTDB = 0x00008000
    //      // Clear attribute when copying outside the database. 
    //}

    [Flags]
    public enum FetchDocumentFlags : uint
    {
        CheckOut = 0x00000000,
        Export = 0x00000001,
        CopyOut = 0x00000002,
        Refresh = 0x00000004,
        Lock = 0x00000008,
        UseUpToDateCopy = 0x00000010,
        AcceptCheckouts = 0x00000020,
        CopyOutMasters = 0x00000040,
        AsSetMembers = 0x00001000,
        ExportReferences = 0x00002000,
        ChangeSetId = 0x00004000,
        UseVaultDirs = 0x00008000,
        IgnoreMasters = 0x00010000,
        GiveOut = 0x00020000,
        MarkAsView = 0x00040000,
        View = 0x00080000,
        NoAuditTrail = 0x00080000,
        AddToMRU = 0x00100000,
        IgnoreExport = 0x00400000,
        DO_NOT_CHANGE_SET_ID_FOR_CHECKED_OUT_REFERENCES = 0x00200000,
        SHARED_CHECKOUT = 0x01000000,
        MASTER_AS_SET = 0x10000000,
        IGNORE_REDLINE_REL = 0x20000000,
        NESTED_REFERENCES = 0x40000000,
        REDLINED_REFERENCES = 0x80000000,
        SEND_TO_FOLDER = 0x00020100,
        SHARED_EXPORT = 0x01000001
    }

    public enum LinkDataProperty : int
    {
        TableID = 1,
        ColumnID = 2,
        ColumnType = 3,
        ColumnLength = 4,
        ColumnName = 5,
        ColumnFormat = 6,
        ColumnDescription = 7,
        ColumnNativeType = 8
    }

    public enum EnvTriggerProperty : int
    {
        EnvironmentID = 1,
        TableID = 2,
        ColumnID = 3,
        TriggeredColumnID = 4,
        OrderNumber = 5,
        ValueType = 6,
        Value = 7
    }

    public enum LinkProperty : int
    {
        VaultID = 1,
        DocumentID = 2,
        TableID = 3,
        ColumnID = 4,
        ColumnValue = 5,
        DocGuid = 6
    }


    public enum VaultType : int
    {
        Normal = 0,
        Workspace = 1
    }


    public enum VaultDescriptorFlags : uint
    {
        VaultID = 0x00000001,
        EnvironmentID = 0x00000002,
        ParentID = 0x00000004,
        StorageID = 0x00000008,
        ManagerID = 0x00000010,
        TypeID = 0x00000020,
        Workflow = 0x00000040,
        Name = 0x00000080,
        Description = 0x00000100,
        Configuration = 0x00000200,
        ManagerType = 0x00000400,
        WSpaceProfID = 0x00000800
    }


    public enum EnvironmentProperty : int
    {
        ID = 1,
        TableID = 2,
        Flags = 3,
        AttrNo = 4,
        Name = 5,
        Desc = 6,
        ViewID = 7
    }


    [Flags]
    public enum DocumentCreationFlag : uint
    {
        Default = 0x00000000,
        NoAttributeRecord = 0x00000001,
        CreateAttributeRecord = 0x00000002,
        NoAuditTrail = 0x00000004
    }

    public enum DocumentDeleteMasks : int
    {
        None = 0x00000000,
        NoSetChild = 0x00000001,
        NoSetParent = 0x00000002,
        MoveAction = 0x00000004,
        IncludeVersions = 0x00000008
    }

    [Flags]
    public enum MenuItemStateFlag : uint
    {
        Show = 0x00000000,
        Hidden = 0x00000001,
        GrayedOut = 0x00000002,
        Checked = 0x00000004,
        ForceShow = 0x00000008,
        Popup = 0x00001000,
        Separator = 0x00002000,
        Undefined = 0x0000FFFF,
        StateMask = 0x00000FFF
    }

    [Flags]
    public enum SetTypeMasks : uint
    {
        Flat = 0x00010000,
        Logical = 0x00020000,
        Redline = 0x00080000,
        Ref = 0x00100000,
        Satellite = 0x00200000,
        All = 0x00FF0000
    }

    [Flags]
    public enum NewVersionCreationFlags : uint
    {
        None = 0x00000000,
        CopyAttrs = 0x00000001,
        KeepRelations = 0x00000002
    }

    [Flags]
    public enum CreateVersionsFromSourceFlags : uint
    {

        //brief  Source document attributes are to be added to the new target document version.
        AARULEO_ADD_SOURCE_ATTRS = 0x00000001,
        // Target document attributes are to be removed from the new version.
        AARULEO_REMOVE_TARGET_ATTRS = 0x00000002,
        //document name is to be applied to the new target document version.
        AARULEO_APPLY_SOURCE_NAME = 0x00000004,
        //Source document file name is to be applied to the new target document version.
        AARULEO_APPLY_SOURCE_FNAME = 0x00000008,
        //Do not send notifications when new versions are created.
        AARULEO_NO_UPDATE = 0x00000010,
        //Do not use wizards to create new versions.
        AARULEO_NO_WIZARDS = 0x00000020,
        //Recreate set relations between created document versions.
        /* SKIP_SAME_TABLE_ATTRS used with ADD_SOURCE_ATTRS, skips attributes if source
           and target documents are in same environment*/
        AARULEO_RECREATE_RELATIONS = 0x00000040,
        AARULEO_SKIP_SAME_TABLE_ATTRS = 0x00000080,
        AARULEO_INCLUDE_VERSIONS = 0x00000100
    }

    public enum DocumentType : int
    {
        Normal = 10,
        History = 11,
        Set = 12,
        Redline = 13,
        ModelerBRP = 14,
        Abstract = 15,
        Unknown = 0
    }


    public enum ApplicationProperty : int
    {
        ID = 1,
        Name = 2,
        ViewerId = 3
    }

    public enum ApplActionProperty : int
    {
        APPLACTION_PROP_APPLICATION_ID = 1,
        APPLACTION_PROP_USER_ID = 2,
        APPLACTION_PROP_ACTION_TYPE = 3,
        APPLACTION_PROP_FLAGS = 4,
        APPLACTION_PROP_PROGRAM_CLASS = 5,
        APPLACTION_PROP_EXECUTABLE_PATH = 6,
        APPLACTION_PROP_PROGRAM_NAME = 7,
        APPLACTION_PROP_ARGUMENTS = 8,
        APPLACTION_PROP_IS_DEFAULT = 9
    }

    public enum CodeDefinitionType : int
    {
        All = -1,
        PartOfCode = 1,
        DocumentCodePlaceHolder = 2,
        AdditionalDocumentCode = 3,
        RevisionPlaceHolder = 4
    }

    public enum DocumentCodeDefinitionProperty : int
    {
        EnvironmentID = 1,
        TableID = 2,
        ColumnID = 3,
        Type = 4,
        SerialType = 5,
        Flags = 6,
        OrderNumber = 7,
        ConnectString = 8
    }

    public enum DocumentCodeDefinitionFlag : uint
    {
        AllowEmpty = 0x00000002
    }

    public enum DocumentCodeSerialType : int
    {
        None = 0,
        Number = 1,
        UsedWith = 2
    }

    public enum AttributeDefinitionProperty : int
    {
        EnvironmentID = 1,
        TableID = 2,
        ColumnID = 3,
        ControlType = 4,
        EditFontHeight = 5,
        FieldFlags = 6,
        DefaultValueType = 7,
        FieldLength = 8,
        FieldAccess = 9,
        ValueListType = 10,
        EditFont = 11,
        DefaultValue = 12,
        FieldFormat = 13,
        ValueListSource = 14,
        Extra1 = 15,
        Extra2 = 16,
        Extra3 = 17,
        Extra4 = 18,
        Extra5 = 19
    }

    public enum DatasourceStatisticsProperty : int
    {
        STAT_PROP_LAST_UPDATED = 0,   //< Last time the statistics were calculated. This value may be cast to a __time64_t value 
        STAT_PROP_USER_COUNT = 1,   //< The number of users records in the datasource. 
        STAT_PROP_GROUP_COUNT = 2,   //< The number of user groups in the datasource. 
        STAT_PROP_MIN_USERS_PER_GROUP = 3,   //< The smallest number of users in a group. 
        STAT_PROP_MAX_USERS_PER_GROUP = 4,   //< The largest number of users in a group. 
        STAT_PROP_AVG_USERS_PER_GROUP = 5,   //< The average number of users in a group. 
        STAT_PROP_POPULATED_FOLDER_COUNT = 6,   //< The number of folders/projects that have at least one document. 
        STAT_PROP_MIN_DOCS_PER_FOLDER = 7,   //< The smallest number of documents in a folder. 
        STAT_PROP_MAX_DOCS_PER_FOLDER = 8,   //< The largest number of documents in a folder. 
        STAT_PROP_AVG_DOCS_PER_FOLDER = 9,   //< The average number of documents in a folder. 
        STAT_PROP_ITEMS_WITH_AUDIT_RECS = 10,   //< The number of objects in the datasource with at least one audit trail record. 
        STAT_PROP_MIN_AUDIT_RECS_PER_FOLDER = 11,   //< The smallest number of audit trail records for a folder. 
        STAT_PROP_MAX_AUDIT_RECS_PER_FOLDER = 12,   //< The largest number of audit trail records for a folder. 
        STAT_PROP_AVG_AUDIT_RECS_PER_FOLDER = 13,   //< The average number of audit trail records per folder. 
        STAT_PROP_MIN_REFERENCE_ATTACHMENTS = 14,   //< The smallest number of reference attachments managed by ProjectWise. 
        STAT_PROP_MAX_REFERENCE_ATTACHMENTS = 15,   //< The largest number of reference attachments managed by ProjectWise. 
        STAT_PROP_AVG_REFERENCE_ATTACHMENTS = 16,   //< The average number of reference attachments managed by ProjectWise. 
        STAT_PROP_MAX_FOLDER_DEPTH = 17,   //< The number of folders in the "deepest" folder hierarchy. 
        STAT_PROP_MIN_FOLDER_DEPTH = 18,   //< The number of folders in the "shallowest" folder hierarchy. 
        STAT_PROP_AVG_FOLDER_DEPTH = 19,   //< The average number of folders in a folder hierarchy. 
        STAT_PROP_STORAGE_COUNT = 20,   //< The number of storage area records in the datasource. 
        STAT_PROP_FOLDER_COUNT = 21,   //< The number of folders in the datasource. 
        STAT_PROP_PROJECT_COUNT = 22,   //< The number of projects in the datasource. 
        STAT_PROP_RICHPROJECT_COUNT = 23,   //< The number of rich projects in the datasource. 
        STAT_PROP_DOC_COUNT = 24,   //< The number of documents in the datasource. 
        STAT_PROP_AUDIT_COUNT = 25,   //< The number of audit trail records in the datasource. 
        STAT_PROP_ENVIRONMENT_COUNT = 26,   //< The number of defined environments in the datasource. 
        STAT_PROP_VIEW_COUNT = 27,   //< The number of defined views (user and global) in the datasource. 
        STAT_PROP_PROPERTY_COUNT = 28,   //< The number of defined extraction properties in the datasource. 
        STAT_PROP_WORKFLOW_COUNT = 29,   //< The number of workflows defined in the datasource. 
        STAT_PROP_THUMB_COUNT = 30,   //< The number of thumbnail records in the datasource. 
        STAT_PROP_QUERY_COUNT = 31,   //< The number of saved queries in the datasource. 
        STAT_PROP_ACCESSCONTROL_COUNT = 32,   //< The number of access control records in the datasource. 
        STAT_PROP_MASTERFILE_COUNT = 33,   //< The number of set masters in the datasource. 
        STAT_PROP_DEPARTMENT_COUNT = 34,   //< The number of departments in the datasource. 
        STAT_PROP_STATE_COUNT = 35,   //< The number of states definitions in the datasource. 
        STAT_PROP_ODS_CLASS_COUNT = 36,   //< The number of defined ODS classes in the datasource. 
        STAT_PROP_MAX_DOC_FILE_SIZE = 37,   //< The size of the largest document file in the datasource. 
        STAT_PROP_MIN_DOC_FILE_SIZE = 38,   //< The size of the smallest document file in the datasource. 
        STAT_PROP_AVG_DOC_FILE_SIZE = 39,   //< The average size of a document file in the datasource. 
        STAT_PROP_DOCS_PROCESSING = 40,   //< The number of documents marked as being currently processed by the extraction system. 
        STAT_PROP_DOCS_PROCESSED = 41,   //< The number of documents marked as being completed processing by the extraction system. 
        STAT_PROP_DOCS_WITH_THUMB = 42,   //< The number of documents in the datasource with thumbnail records. 
        STAT_PROP_DOCS_WITHOUT_THUMB = 43,   //< The number of documents in the datasource without thumbnail records. 
        STAT_PROP_DOCS_NOT_PROCESSED = 44,   //< The number of documents in the datasource that have not been processed by the extraction system. 
        STAT_PROP_EMPTY_FOLDER_COUNT = 45,   //< The number of folders in the datasource containing no objects. 
        STAT_PROP_ATTRIBUTE_COUNT = 46,   //< The number of defined attributes in the datasource. 
        STAT_PROP_MIN_ATTRIBUTES_PER_ENV = 47,   //< The smallest number of attributes in an environment. 
        STAT_PROP_MAX_ATTRIBUTES_PER_ENV = 48,   //< The largest number of attributes in an environment. 
        STAT_PROP_AVG_ATTRIBUTES_PER_ENV = 49,   //< The average number of attributes in an environment. 
        STAT_PROP_CHKL_REC_COUNT = 50,   //< The number of document check out location records in the datasource. 
        STAT_PROP_ORPHANED_THUMB_COUNT = 51,   //< The number of thumbnail records in the datasource that do not correspond to a document record. 
        STAT_PROP_DOCS_WITH_FPROPS_COUNT = 52,   //< The number of document records in the datasource that have extracted file properties. 
        STAT_PROP_AFP_FTR_DOCS_PROCESSED = 53,   //< The number of document records in the datasource that have been processed by the full-text-retrieval engine. 
        STAT_PROP_AFP_FTR_DOCS_UNPROCESSED = 54,   //< The number of document records in the datasource that have not been processed by the full-text-retrieval engine. 
        STAT_PROP_AFP_FTR_DOCS_PROCESSING = 55,   //< The number of document records in the datasource that are being processed by the full-text-retrieval engine. 
        STAT_PROP_AFP_THUMB_DOCS_PROCESSED = 56,   //< The number of document records in the datasource that have been processed by the thumbnail extraction engine. 
        STAT_PROP_AFP_THUMB_DOCS_UNPROCESSED = 57,   //< The number of document records in the datasource that have not been processed by the thumbnail extraction engine. 
        STAT_PROP_AFP_THUMB_DOCS_PROCESSING = 58,   //< The number of document records in the datasource that are being processed by the thumbnail extraction engine. 
        STAT_PROP_AFP_FPROP_DOCS_PROCESSED = 59,   //< The number of document records in the datasource that have been processed by the file property extraction engine. 
        STAT_PROP_AFP_FPROP_DOCS_UNPROCESSED = 60,   //< The number of document records in the datasource that have not been processed by the file property extraction engine. 
        STAT_PROP_AFP_FPROP_DOCS_PROCESSING = 61,  //< The number of document records in the datasource that are being processed by the file property extraction engine. 

        // Added by Brian Flaherty on 4/30/2018
        //   Adding additional statistics properties
        STAT_PROP_USER_NOLOGIN_1MONTH = 62, //< The number of users who have not logged in to the datasource in the last month. 
        STAT_PROP_USER_NOLOGIN_3MONTHS = 63, //< The number of users who have not logged in to the datasource in the last 3 months. 
        STAT_PROP_USER_NOLOGIN_4MONTHS = 64, //< The number of users who have not logged in to the datasource in the last 4 months. 
        STAT_PROP_USER_NOLOGIN_1YEAR = 65, //< The number of users who have not logged in to the datasource in the last year. 
        STAT_PROP_USER_NEVER_LOGGED_IN = 66, //< The number of users who have not logged in to the datasource ever. 
        STAT_PROP_TOTAL_DOC_FILE_SIZE = 67, //< The total file size of all documents. 
        STAT_PROP_DB_STATS_LAST_UPDATED = 68, //< The database statistics last update time.
        // End addition

        STAT_PROP_LAST = STAT_PROP_DB_STATS_LAST_UPDATED
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct AAEATTRDEF
    {
        public int iControlType;
        public int iEditFontHeight;
        public int iFieldFlags;
        public int iDefaultValueType;
        public int iFieldLength;
        public int iFieldAccess;
        public int iValueListType;
        public string sEditFont;
        public string sDefaultValue;
        public string sFieldFormat;
        public string sValueListSource;
        public string sExtra1;
        public string sExtra2;
        public string sExtra3;
        public string sExtra4;
        public string sExtra5;
    }

    public enum AttributeParameterFlags : int
    {
        Unique = 0x00000001,
        Required = 0x00000002,
        EditableIfFinal = 0x00000010,
        MultiSelect = 0x00000040,
        LimitToList = 0x00000080,
        CopyClearNewSheet = 0x00001000,
        CopyClearInEnvironment = 0x00002000,
        CopyClearOutEnvironment = 0x00004000,
        CopyClearOutDatabase = 0x00008000
    }


    public enum AttributeDefaultValueTypes : int
    {
        None = 0,
        Fixed = 1,
        Select = 2,
        SystemVariable = 3,
        Function = 4
    }




    // dww 2013-10-02 matched SS4 documentation
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct VaultDescriptor
    {
        public uint Flags;
        public int VaultID;
        public int EnvironmentID;
        public int ParentID;
        public int StorageID;
        public int ManagerID;
        public int TypeID;
        public int WorkflowID;
        public string Name;
        public string Description;
        public int ManagerType;
        public int WorkspaceProfileId;
        public Guid GuidVault;
        public int ComponentClassId;
        public int ComponentinstanceId;
        public uint projFlagMask;
        public uint projFlags;
    }

    public enum DialogBoxCommandId : int
    {
        IDOK = 1,
        IDCANCEL = 2,
        DABORT = 3,
        IDRETRY = 4,
        IDIGNORE = 5,
        IDYES = 6,
        IDNO = 7,
        IDCLOSE = 8,
        IDHELP = 9
    }

    // dww 2014-01-24 updated to include force doc name locking & others
    public enum DatasourceGenericSettings : int
    {
        AADMS_GEN = 100,
        AADMS_GEN_CREATE_ON_FIRST_STATE = (100 + 2),
        AADMS_GEN_CAN_CHANGE_VERSION_STATE = (100 + 3),
        AADMS_GEN_USE_RECYCLE_BIN = (100 + 4),
        AADMS_GEN_CREATE_DOC_ON_STATE = (100 + 5),
        AADMS_GEN_GDR = (100 + 6),
        AADMS_GEN_COMMON_ENVIRONMENT = (100 + 7),
        AADMS_GEN_VERSION_FLAGS = (100 + 8),
        AADMS_GEN_FIRST_VERSION_NO = (100 + 9),
        AADMS_GEN_DOC_STATE_CHANGE_FLAGS = (100 + 10),
        AADMS_GEN_NOTIFY_ABOUT_EVENT_SUCCESS = (100 + 11),
        AADMS_GEN_AUDT_DOC_TRACE = (100 + 12),
        AADMS_GEN_AUDT_VLT_TRACE = (100 + 13),
        AADMS_GEN_AUDT_SET_TRACE = (100 + 14),
        AADMS_GEN_AUDT_TRUNCATE = (100 + 15),
        AADMS_GEN_AUDT_TRUNCATE_PARAM = (100 + 16),
        AADMS_GEN_AUDT_TRUNCATE_TIME_UNITS = (100 + 17),
        AADMS_GEN_AUDT_TRUNCATE_USE_TABLE = (100 + 18),
        AADMS_GEN_AUDT_TRUNCATE_TABLE_NAME = (100 + 19),
        AADMS_GEN_FORCE_CASE_INSENS_FIND_DOCS = (100 + 20),
        AADMS_GEN_CAN_MOVE_VERSIONS = (100 + 21),
        AADMS_GEN_ENG_COMP_ENABLED = (100 + 22),
        AADMS_GEN_SAVED_SEARCH_FOLDER = (100 + 23),
        AADMS_GEN_IMAGES_FOLDER = (100 + 24),
        AADMS_GEN_MRU_ITEMS_TO_SHOW = (100 + 25),
        AADMS_GEN_MRU_TRUNCATE_VALUE = (100 + 26),
        AADMS_GEN_MRU_TRUNCATE_UNITS = (100 + 27),
        AADMS_GEN_DATASOURCE_GUID = (100 + 28),
        AADMS_GEN_DEFAULT_RIGHTS_ADD_EVERYONE = (100 + 29),
        AADMS_GEN_DEFAULT_RIGHTS_CLASS_MASK = (100 + 30),
        AADMS_GEN_DEFAULT_RIGHTS_ATTRIBUTE_MASK = (100 + 31),
        AADMS_GEN_DEFAULT_RIGHTS_METHOD_MASK = (100 + 32),
        AADMS_GEN_SPATIAL_NEW_FOLDER_INHERIT = (100 + 38),
        AADMS_GEN_SPATIAL_NEW_DOC_INHERIT = (100 + 39),
        AADMS_GEN_WEB_DEFAULT_SERVER = (100 + 40),
        AADMS_GEN_USE_UNION_BASED_FIND_DOCS = (100 + 41),
        AADMS_GEN_SHAREPOINT_DEFAULT_SERVER = (100 + 42),
        AADMS_GEN_PUBLISHER_DEFAULT_SERVER = (100 + 43),
        AADMS_GEN_TEMPLATES_FOLDER = (100 + 44),
        AADMS_GEN_SHAREABLE_DOCUMENTS = (100 + 45),
        AADMS_GEN_DEFAULT_WEB_VIEWER = (100 + 46),
        AADMS_GEN_ENABLE_DELTA_FILE_XFER = (100 + 48),
        AADMS_GEN_ENABLE_REQUEST_COMPRESSION = (100 + 49),
        AADMS_GEN_DESIGN_COMPARE_ADDRESS = (100 + 50),
        AADMS_GEN_ALLOW_SAME_WORKING_DIR = (100 + 52),
        AADMS_GEN_UPDATE_ACC_ON_DOC_DEL = (100 + 53),
        AADMS_GEN_VERSIONING_RULES = (100 + 56),
        AADMS_GEN_COPY_AUDIT_TRAIL = (100 + 57),
        AADMS_GEN_PROJ_COPY_EVENTS = (100 + 59),
        AADMS_GEN_DOC_COPY_EVENTS = (100 + 60),
        AADMS_GEN_ENABLE_FT_CONNECTION_CACHE = (100 + 61),
        AADMS_GEN_RESTRICTED_ADMIN_GROUP = (100 + 62),
        AADMS_GEN_ENABLE_ACL_EDITOR = (100 + 63),
        AADMS_GEN_FORCE_DOC_NAME_LOCKING = (100 + 64),
        AADMS_GEN_GRANT_IMPLICIT_OWNER_ACL = (100 + 65)
    }

    public enum DocumentOperations : int
    {
        AAOPER_DOC_FIRST = 1000,
        AAOPER_DOC_CREATE = AAOPER_DOC_FIRST + 0,
        AAOPER_DOC_CREATE_LEAVE_OUT = AAOPER_DOC_FIRST + 1,
        AAOPER_DOC_COPY = AAOPER_DOC_FIRST + 2,
        AAOPER_DOC_MOVE = AAOPER_DOC_FIRST + 3,
        AAOPER_DOC_DELETE = AAOPER_DOC_FIRST + 4,
        AAOPER_DOC_MODIFY = AAOPER_DOC_FIRST + 5,
        AAOPER_DOC_CHECKOUT = AAOPER_DOC_FIRST + 6,
        AAOPER_DOC_COPYOUT = AAOPER_DOC_FIRST + 7,
        AAOPER_DOC_EXPORT = AAOPER_DOC_FIRST + 8,
        AAOPER_DOC_CHECKIN = (AAOPER_DOC_FIRST + 9),
        AAOPER_DOC_CHECKIN_LEAVE_COPY = (AAOPER_DOC_FIRST + 10),      /**< Specifies to check in the document and leave local copy */

        AAOPER_DOC_CREATE_LINK_DATA = (AAOPER_DOC_FIRST + 30),      /**<  */
        AAOPER_DOC_UPDATE_LINK_DATA = (AAOPER_DOC_FIRST + 31),      /**<  */
        AAOPER_DOC_DELETE_LINK_DATA = (AAOPER_DOC_FIRST + 32),      /**<  */
    }

    public enum HookActions : int
    {
        AAHOOK_SUCCESS = 0,
        AAHOOK_ERROR = 1,
        AAHOOK_CALL_NEXT_IN_CHAIN = 2,
        AAHOOK_CALL_DEFAULT = 3
    }

    // dwww 2013-10-08 added a bunch of hook identifiers for SS4
    public enum HookIdentifiers : int
    {
        AADMSHOOK_FIRST = 1002,
        AAHOOK_LOGIN = (AADMSHOOK_FIRST + 0),
        AAHOOK_LOGOUT = (AADMSHOOK_FIRST + 1),
        AAHOOK_CREATE_PROJECT = (AADMSHOOK_FIRST + 100),
        AAHOOK_MOVE_PROJECT = (AADMSHOOK_FIRST + 101),
        AAHOOK_DELETE_PROJECT = (AADMSHOOK_FIRST + 102),
        AAHOOK_MODIFY_PROJECT = (AADMSHOOK_FIRST + 103),
        AAHOOK_PROJECT_WORKFLOW = (AADMSHOOK_FIRST + 104),
        AAHOOK_CHECKOUT_PROJECT = (AADMSHOOK_FIRST + 105),
        AAHOOK_COPYOUT_PROJECT = (AADMSHOOK_FIRST + 106),
        AAHOOK_PURGE_PROJECT = (AADMSHOOK_FIRST + 107),
        AAHOOK_EXPORT_PROJECT = (AADMSHOOK_FIRST + 108),
        AAHOOK_UPGRADE_PROJECT_TO_RICHPRJ = (AADMSHOOK_FIRST + 109),
        AAHOOK_DOWNGRADE_PROJECT_TO_FOLDER = (AADMSHOOK_FIRST + 110),
        AAHOOK_CREATE_DOCUMENT = (AADMSHOOK_FIRST + 200),
        AAHOOK_MOVE_DOCUMENT = (AADMSHOOK_FIRST + 201),
        AAHOOK_DELETE_DOCUMENT = (AADMSHOOK_FIRST + 202),
        AAHOOK_MODIFY_DOCUMENT = (AADMSHOOK_FIRST + 203),
        AAHOOK_CHECKOUT_DOCUMENT = (AADMSHOOK_FIRST + 204),
        AAHOOK_COPYOUT_DOCUMENT = (AADMSHOOK_FIRST + 205),
        AAHOOK_EXPORT_DOCUMENT = (AADMSHOOK_FIRST + 207),
        AAHOOK_CHECKIN_DOCUMENT = (AADMSHOOK_FIRST + 208),
        AAHOOK_PURGE_DOCUMENT_COPY = (AADMSHOOK_FIRST + 209),
        AAHOOK_FREE_DOCUMENT = (AADMSHOOK_FIRST + 210),
        AAHOOK_REFRESH_DOC_SERV_COPY = (AADMSHOOK_FIRST + 211),
        AAHOOK_REFRESH_DOCUMENT_COPY = (AADMSHOOK_FIRST + 212),
        AAHOOK_CHANGE_DOC_VERSION = (AADMSHOOK_FIRST + 213),
        AAHOOK_CHANGE_DOC_STATE = (AADMSHOOK_FIRST + 214),
        AAHOOK_CREATE_REDLINE_DOC = (AADMSHOOK_FIRST + 215),
        AAHOOK_UPDATE_LINK_DATA = (AADMSHOOK_FIRST + 216),
        AAHOOK_DELETE_LINK_DATA = (AADMSHOOK_FIRST + 217),
        AAHOOK_LOCK_DOCUMENT = (AADMSHOOK_FIRST + 218),
        AAHOOK_ADD_DOCUMENT_FILE = (AADMSHOOK_FIRST + 219),
        AAHOOK_DELETE_DOCUMENT_FILE = (AADMSHOOK_FIRST + 220),
        AAHOOK_CHANGE_DOCUMENT_FILE = (AADMSHOOK_FIRST + 221),
        AAHOOK_FETCH_MULTIDOCS = (AADMSHOOK_FIRST + 222),
        AAHOOK_DELETE_DOCUMENT_EXT = (AADMSHOOK_FIRST + 223),
        AAHOOK_DELETE_DOCUMENTS = (AADMSHOOK_FIRST + 224),
        AAHOOK_COPY_DOCUMENT_CROSS_DS = (AADMSHOOK_FIRST + 225),
        AAHOOK_CREATE_SET = (AADMSHOOK_FIRST + 300),
        AAHOOK_ADD_SET_MEMBER = (AADMSHOOK_FIRST + 301),
        AAHOOK_DELETE_SET_MEMBER = (AADMSHOOK_FIRST + 302),
        AAHOOK_VERIFY_VERSION = (AADMSHOOK_FIRST + 400),
        AAHOOK_VERIFY_TABLES = (AADMSHOOK_FIRST + 401),
        AAHOOK_CREATE_TABLES = (AADMSHOOK_FIRST + 402),
        AAHOOK_COPY_DOC_ATTRIBUTES = (AADMSHOOK_FIRST + 500),
        AAHOOK_DELETE_USER = (AADMSHOOK_FIRST + 600),
        AAHOOK_SET_DOC_FINAL_STATUS = (AADMSHOOK_FIRST + 601),
        AAHOOK_DELETE_GROUP = (AADMSHOOK_FIRST + 602),
        AAHOOK_DELETE_WORKFLOW = (AADMSHOOK_FIRST + 603),
        AAHOOK_DELETE_STATE = (AADMSHOOK_FIRST + 604),
        AAHOOK_DEL_WORKFLOW_STATE = (AADMSHOOK_FIRST + 605),
        AAHOOK_DELETE_ENVIRONMENT = (AADMSHOOK_FIRST + 606),
        AAHOOK_INVALIDATE_CACHE = (AADMSHOOK_FIRST + 607),
        AAHOOK_ACTIVATE_INTERFACE = (AADMSHOOK_FIRST + 608),
        AAHOOK_COPY_DOCUMENTS = (AADMSHOOK_FIRST + 700),
        AAHOOK_DELETE_USERLIST = (AADMSHOOK_FIRST + 701),
        AAHOOK_COPY_DOCUMENTS_CROSS_DS = (AADMSHOOK_FIRST + 702),
        AAHOOK_CREATE_VIEW = (AADMSHOOK_FIRST + 715),
        AAHOOK_MODIFY_VIEW = (AADMSHOOK_FIRST + 716),
        AAHOOK_DELETE_VIEW = (AADMSHOOK_FIRST + 717),
        AAHOOK_ENUMERATE_VIEWS = (AADMSHOOK_FIRST + 718),
        AAHOOK_GET_VIEWCOLUMN_NAME = (AADMSHOOK_FIRST + 719),
        AAHOOK_GEN_SETTING_SET_VALUE = (AADMSHOOK_FIRST + 730),
        AAHOOK_USER_SETTING_SET_VALUE = (AADMSHOOK_FIRST + 731),
        AAHOOK_GROUP_MEMBER_CHANGE = (AADMSHOOK_FIRST + 732),
        AADMSHOOK_LAST = (AAHOOK_GROUP_MEMBER_CHANGE),
        AADMSHOOK_LAST_RESERVED = 2000,
        AAWINDMSHOOK_FIRST = 3001,
        AAHOOK_OPEN_DOCUMENT = (AAWINDMSHOOK_FIRST + 0),
        AAHOOK_PRINT_DOCUMENT = (AAWINDMSHOOK_FIRST + 1),
        AAHOOK_START_APPLICATION = (AAWINDMSHOOK_FIRST + 2),
        AAHOOK_DOC_SEND_MAIL = (AAWINDMSHOOK_FIRST + 4),
        AAHOOK_VALIDATE_FILE = (AAWINDMSHOOK_FIRST + 5),
        AAHOOK_LOGIN_DLG = (AAWINDMSHOOK_FIRST + 7),
        AAHOOK_CREATE_PROJECT_DLG = (AAWINDMSHOOK_FIRST + 8),
        AAHOOK_MODIFY_PROJECT_DLG = (AAWINDMSHOOK_FIRST + 9),
        AAHOOK_PROJECT_PROPERTY_DLG = (AAWINDMSHOOK_FIRST + 10),
        AAHOOK_SELECT_PROJECT_DLG = (AAWINDMSHOOK_FIRST + 11),
        AAHOOK_CREATE_DOCUMENT_DLG = (AAWINDMSHOOK_FIRST + 12),
        AAHOOK_SAVE_DOCUMENT_DLG = (AAWINDMSHOOK_FIRST + 13),
        AAHOOK_OPEN_DOCUMENT_DLG = (AAWINDMSHOOK_FIRST + 14),
        AAHOOK_MODIFY_DOCUMENT_DLG = (AAWINDMSHOOK_FIRST + 15),
        AAHOOK_DOCUMENT_PROPERTY_DLG = (AAWINDMSHOOK_FIRST + 16),
        AAHOOK_SELECT_DOCUMENT_DLG = (AAWINDMSHOOK_FIRST + 17),
        AAHOOK_FIND_DOCUMENT_DLG = (AAWINDMSHOOK_FIRST + 19),
        AAHOOK_DOCUMENT_VERSION_DLG = (AAWINDMSHOOK_FIRST + 20),
        AAHOOK_WORKFLOW_DLG = (AAWINDMSHOOK_FIRST + 21),
        AAHOOK_CREATE_SET_DLG = (AAWINDMSHOOK_FIRST + 22),
        AAHOOK_MODIFY_SET_DLG = (AAWINDMSHOOK_FIRST + 23),
        AAHOOK_USER_SETTINGS_DLG = (AAWINDMSHOOK_FIRST + 24),
        /*AAHOOK_CHECKIN_COMMENT_DLG          = (AAWINDMSHOOK_FIRST + 25), < Not used. */
        AAHOOK_VIEW_EDITOR_DLG = (AAWINDMSHOOK_FIRST + 26),
        AAHOOK_VIEW_DOCUMENTS = (AAWINDMSHOOK_FIRST + 27),
        AAHOOK_VIEW_FILE = (AAWINDMSHOOK_FIRST + 28),
        AAHOOK_CLOSE_VIEWER = (AAWINDMSHOOK_FIRST + 29),
        AAHOOK_SHOW_NOTICE_WND = (AAWINDMSHOOK_FIRST + 30),
        AAHOOK_SPLASH_WINDOW = (AAWINDMSHOOK_FIRST + 31),
        AAHOOK_CREATE_REDLINE_DOC_DLG = (AAWINDMSHOOK_FIRST + 32),
        AAHOOK_SELECT_REDLINE_DOC_DLG = (AAWINDMSHOOK_FIRST + 33),
        AAHOOK_START_REDLINE = (AAWINDMSHOOK_FIRST + 34),
        AAHOOK_REDLINE_FIND_FILE = (AAWINDMSHOOK_FIRST + 35),
        AAHOOK_IMPORTBYDROPHANDLE = (AAWINDMSHOOK_FIRST + 37),
        AAHOOK_EXEC_MENU_COMMAND = (AAWINDMSHOOK_FIRST + 38),
        AAHOOK_INIT_POPUPMENU = (AAWINDMSHOOK_FIRST + 39),
        AAHOOK_SELECT_INTERFACE_DLG = (AAWINDMSHOOK_FIRST + 40),
        /* AAHOOK_SAVE_DOCUMENT_DLG2           = (AAWINDMSHOOK_FIRST + 41), */
        AAHOOK_OPEN_DOCUMENT_DLG2 = (AAWINDMSHOOK_FIRST + 42),
        AAHOOK_PROJECT_EXPORT_WZRD = (AAWINDMSHOOK_FIRST + 43),
        AAHOOK_CREATE_DOCUMENTS_DLG = (AAWINDMSHOOK_FIRST + 44),
        AAHOOK_TRANSFER_DOCUMENT_DLG = (AAWINDMSHOOK_FIRST + 45),
        AAHOOK_SELECT_ENVIRONMENT_DLG = (AAWINDMSHOOK_FIRST + 46),
        AAHOOK_CODE_RESERVATION_DLG = (AAWINDMSHOOK_FIRST + 47),
        AAHOOK_DOCUMENT_EXPORT_WZRD = (AAWINDMSHOOK_FIRST + 48),
        AAHOOK_OPEN_DOCUMENTS_DLG2 = (AAWINDMSHOOK_FIRST + 49),
        AAHOOK_SHOW_DOC_PROP_PAGE = (AAWINDMSHOOK_FIRST + 50),
        AAHOOK_SHOW_PROJ_PROP_PAGE = (AAWINDMSHOOK_FIRST + 51),
        AAHOOK_EXECUTE_DOC_ACTION = (AAWINDMSHOOK_FIRST + 52),
        AAHOOK_RELOAD_UPDATED_DOCS_DLG = (AAWINDMSHOOK_FIRST + 53),
        AAHOOK_CODE_GENERATION_DLG = (AAWINDMSHOOK_FIRST + 54),
        AAHOOK_SAVE_DOCUMENT_DLG3 = (AAWINDMSHOOK_FIRST + 55),
        AAHOOK_DOCUMENT_IN_USE_CHECK = (AAWINDMSHOOK_FIRST + 56),
        AAHOOK_IMPORTBYSTGMEDIUM = (AAWINDMSHOOK_FIRST + 57),
        AAHOOK_DLG_APPCHANGED = (AAWINDMSHOOK_FIRST + 58),
        AAHOOK_OPEN_MULTI_DOCUMENTS_DLG = (AAWINDMSHOOK_FIRST + 59),
        AAWINDMSHOOK_LAST = (AAHOOK_OPEN_MULTI_DOCUMENTS_DLG),
        AAWINDMSHOOK_LAST_RESERVED = 4000
    }

    public enum HookTypes : int
    {
        AAPREHOOK = 1,
        AAACTIONHOOK = 2,
        AAPOSTHOOK = 3,
        AAPOSTHOOK_FAIL = 4
    }

    public enum MenuCommandIds : uint
    {
        AAMENU_PROJECT_FIRST = 30050,	// dww - added - for consistency

        IDMP_CREATE = 30051,
        IDMP_MODIFY = 30052,
        IDMP_DELETE = 30053,
        IDMP_PURGE = 30054,
        IDMP_WORKFL = 30055,
        IDMP_REFRESH = 30056,
        IDMP_PROPERTY = 30057,
        IDMP_COPYOUT = 30058,
        IDMP_PROACCESS = 30059,
        IDMP_COPY = 30060,
        IDMP_PASTE = 30061,
        IDMP_UFLDRREMOVE = 30068,
        IDMP_PROPERTYREAD = 30070,
        IDMP_EXPORT_TO = 30071,
        IDMP_COPYOUT_TO = 30072,
        IDMP_CLEAN_AUDITTRAIL = 30073,
        IDMP_RENAME = 30074,
        IDMP_EXPORT = 30075,
        IDMP_FIND_DOCS = 30076,
        IDMP_UPGRADE = 30077,
        IDMP_DOWNGRADE = 30078,
        IDMP_SCAN_REFERENCES = 30079,       // dww added for SS3
        IDMP_CREATE_RICHPROJECT = 30080,    // dww - was incorrectly identified as 30079

        AAMENU_DOC_FIRST = 30500,			// dww - added - for consistency

        IDMD_OPEN = 30501,
        IDMD_OPEN_WITH = 30502,
        IDMD_PRINT = 30503,
        IDMD_QUICKVIEW = 30504,
        IDMD_NEW = 30506,
        IDMD_MODIFY = 30507,
        IDMD_SAVE_AS = 30508,
        IDMD_DELETE = 30509,
        IDMD_CHECKIN = 30511,
        IDMD_CHECKOUT = 30512,
        IDMD_COPYOUT = 30513,
        IDMD_PURGECOPY = 30514,
        IDMD_FREE = 30515,
        IDMD_SETMODIFY = 30517,
        IDMD_SETCREATE = 30518,
        IDMD_REMOVE = 30519,
        IDMD_VERSION = 30521,
        IDMD_STATE = 30522,
        IDMD_PROPERTIES = 30525,
        IDMD_EXPORT = 30526,
        IDMD_REDLINE = 30527,
        IDMD_REFRESH_COPY = 30528,
        IDMD_UPDATE_SERVER = 30529,
        IDMD_SEND_MAIL = 30530,
        IDMD_SEND_NOTICE = 30531,
        IDMD_COPY_OUT_TO = 30532,
        IDMD_SET_FINAL = 30533,
        IDMD_REMOVE_FINAL = 30534,
        IDMD_OPEN_READONLY = 30535,
        IDMD_UFLDRREMOVE = 30536,
        IDMD_CHANGE_PREV_STATE = 30539,
        IDMD_CHANGE_NEXT_STATE = 30540,
        IDMD_MODIFY_ATTR = 30541,
        IDMD_COPY_LIST = 30542,
        IDMD_PRINT_LIST = 30543,
        IDMD_COPYLIST_TABSEPARATED = 30545,
        IDMD_COPYLIST_SPACESEPARATED = 30546,
        IDMD_IMPORT = 30547,
        IDMD_GENERATE_CODE = 30548,
        IDMD_DELETE_SHEET = 30549,
        IDMD_CREATE_MULTIPLE = 30550,
        IDMD_OPENVAULT = 30551,
        IDMD_COPY = 30552,
        IDMD_MOVE = 30553,
        IDMD_MEMBERIN = 30554,
        IDMD_MASTER_LINKS = 30555,
        IDMD_RELEASE_MASTER = 30556,
        IDMD_RELEASE_REF = 30557,
        IDMD_COPY_ATTR_DATA = 30558,
        IDMD_PASTE_ATTR_DATA = 30559,
        IDMD_ADD_SHEET = 30560,
        IDMD_CLEAN_AUDITTRAIL = 30561,
        IDMD_RENAME = 30562,
        IDMD_CODE_RESERVATION = 30563,
        IDMD_SEND_MAIL_AS_LINK = 30564,
        IDMD_SHARED_CHECKOUT = 30565,
        IDMD_IMPORT_FROM = 30566,
        IDMD_ADD_COMMENT = 30567,
        IDMD_SCAN_REFERENCES = 30568,	// dww - added - missing
        IDMD_SHOW_MARKUPS = 30569,		// dww - added - first appeared in ss1 documentation

        AAMENU_ITEM_TOOLS_FIRST = 30850,

        IDMT_NOTICES = (AAMENU_ITEM_TOOLS_FIRST + 0),
        IDMT_SETTING_QUERY_DLG = (AAMENU_ITEM_TOOLS_FIRST + 1),
        IDMT_QUERY_DLG = (AAMENU_ITEM_TOOLS_FIRST + 2),
        IDMT_ICON_ASSOCIATION = (AAMENU_ITEM_TOOLS_FIRST + 3),
        IDMT_PROGRAM_ASSOCIATION = (AAMENU_ITEM_TOOLS_FIRST + 4),
        IDMT_EXTENSION_ASSOCIATION = (AAMENU_ITEM_TOOLS_FIRST + 5),
        IDMT_SETTING_LAZY_REFRESH = (AAMENU_ITEM_TOOLS_FIRST + 6),
        IDMT_SETTING_ALL_TABLES = (AAMENU_ITEM_TOOLS_FIRST + 7),
        IDMT_SETTING_SET_OPEN = (AAMENU_ITEM_TOOLS_FIRST + 8),
        IDMT_USER_SETTINGS = (AAMENU_ITEM_TOOLS_FIRST + 9),
        IDMT_SHOW_CHECKIN_ONEXIT = (AAMENU_ITEM_TOOLS_FIRST + 10),
        IDMT_SETTING_UPDATE_NEVER = (AAMENU_ITEM_TOOLS_FIRST + 11),
        IDMT_SETTING_UPDATE_AFTER = (AAMENU_ITEM_TOOLS_FIRST + 12),
        IDMT_SETTING_UPDATE_ONEACH = (AAMENU_ITEM_TOOLS_FIRST + 13),
        IDMT_CHECKIN_DLG = (AAMENU_ITEM_TOOLS_FIRST + 14),
        IDMT_SET_INTERFACE = (AAMENU_ITEM_TOOLS_FIRST + 15),
        IDMT_SCAN_REFERENCES = (AAMENU_ITEM_TOOLS_FIRST + 16),
        IDMT_SHOW_DLG_ON_ERROR = (AAMENU_ITEM_TOOLS_FIRST + 17),
        IDMT_WIZARD_MANAGER = (AAMENU_ITEM_TOOLS_FIRST + 19),
        IDMT_CODE_RESERVATION = (AAMENU_ITEM_TOOLS_FIRST + 20),
        IDMT_USE_LASTVAULT = (AAMENU_ITEM_TOOLS_FIRST + 21),
        IDMT_USE_LASTPAGE = (AAMENU_ITEM_TOOLS_FIRST + 22),
        IDMT_CUSTOMIZE = (AAMENU_ITEM_TOOLS_FIRST + 23),
        IDMT_NETWORK_CONFIG = (AAMENU_ITEM_TOOLS_FIRST + 24),	// dww - added - missing
        IDMT_USER_MANAGEMENT = (AAMENU_ITEM_TOOLS_FIRST + 25)	// dww - added - first appeared in SS3 documentation
    }

    public enum MenuCommandsStateFlags : uint
    {
        AAMF_SEL_MORE_THAN_ONE = 0x00000100
    }

    public enum MenuItemStateFlags : int
    {
        AAMS_GRAYED = 0x0002
    }

    public enum EnvironmentInfo : int
    {
        ENV_PROP_ID = 1,
        //Numeric property. 

        ENV_PROP_TABLEID = 2,
        //Numeric property. 

        ENV_PROP_FLAGS = 3,
        //Numeric property. For the list of possible flag values see Environment Creation Flags. 

        ENV_PROP_ATTRNO = 4,
        //Numeric property. 

        ENV_PROP_NAME = 5,
        //String property. 

        ENV_PROP_DESC = 6,
        //String property. 

        ENV_PROP_VIEWID = 7
        //Not used.  

    }

    public enum DocumentListDefinitions : uint
    {
        AADLSTF_ATTACH_TO_ATTR_SHEET = 0x00000001,
        //Created project combo box contains the list of all datasources and allows user to log-in/activate the datasource. 

        AADLSTF_VIRTUAL_LIST = 0x00000002,
        //Virtual document list will be created. 

        AADLSTF_CHECKBOXES = 0x00000004,
        //Show check boxes in a document list. 

        AADLSTF_ISTYPEWINDOW = 0x00000008,
        //Document list is asociated with some DSC object type. 

        AADLSTF_PUT_NAME_FIRST = 0x00000010,
        //Makes 'Name' the first column. 

        AADLSTF_ENABLE_BACK_FWD_BROWSING = 0x00000080
        //Enable forward/backward functionality.

    }

    public enum DocListUpdateTypeMasks : uint
    {
        AADLUISF_NOMASTERCHECK = 0x00000002,

        AADLUISF_NOSETCHECK = 0x00000001
    }

    public enum AttributeLinkageTypes : int
    {
        AADMS_EALNK_DOCUMENT = 1,
        AADMS_EALNK_ENVIRONMENT = 2

    }

    // Seems to be unused.  Taking out.  2011-01-12  DAB
    //[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
    //public unsafe struct _AADOCTODIR_PARAM
    //{
    //    /** Used for compatibility with older ProjectWise API versions. Specifies which added members in the extended structure can be accessed. This can be a combination of the following flags:
    //    <table>
    //    <tr><td>#AAPARAMMASK_FLAGS</td><td>the \em ulFlags member is present.</td></tr>
    //    <tr><td>#AAPARAMMASK_COMMENT</td><td>The \em lpctstrComment member is present.</td></tr>
    //    <tr><td>#AAPARAMMASK_CHKLSET</td><td>The \em hChklSet member is present.</td></tr>
    //    </table> */
    //    public UInt32 ulMask;                /* 0 or <AAPARAMMASK_FLAGS|AAPARAMMASK_COMMENT|AAPARAMMASK_CHKLSET> */
    //    /** Specifies the number of document items in the \em lpDocuments. */
    //    public Int32 lCount;                /* number of documents in lpDocuments        */
    //    /** Specifies the documents to operate with. */
    //    public _AADOC_ITEM* lpDocuments;           /* Document identifiers                      */
    //    /** Pointer to a null-terminated string specifying the output directory for the documents. If this member is NULL then user's working directory is used. */
    //    [MarshalAs(UnmanagedType.LPWStr)]
    //    public char* lpctstrWorkdir;        /* Working directory                         */
    //    /** Pointer to a string buffer receiving the output file name. */
    //    [MarshalAs(UnmanagedType.LPWStr)]
    //    public char* lptstrFileName;        /* File name with full path                  */
    //    /** Specifies the size of lptstrFileName buffer in character symbols (including the terminating null character). */
    //    public Int32 lBufferSize;           /* Buffer size for file name                 */
    //    /** Specifies the operation. See \ref aadmsdef_DocumentDefinitions_DocumentFetchFlags for more information. This member can be accessed only if ulMask has #AAPARAMMASK_FLAGS set. */
    //    public UInt32 ulFlags;               /* operation flags                           */
    //    /** Audit trail comment. This member can be accessed only if ulMask has #AAPARAMMASK_COMMENT set. */
    //    [MarshalAs(UnmanagedType.LPWStr)]
    //    public char* lpctstrComment;
    //    /* Audit trail comment (valid only if ulMask
    //                                               has AAPARAMMASK_COMMENT set)             */
    //    /** Information about fetched documents provided for Post hook. This member can be accessed only if ulMask has #AAPARAMMASK_CHKLSET set. */
    //    public IntPtr hChklSet;              /* Information about fetched documents. Should be expected in PostHook during Fetch. */
    //    /* Valid only if ulMask has AAPARAMMASK_CHKLSET set! */
    //}

    public /*unsafe*/ struct _AADOC_ITEM
    {
        public Int32 lProjectId;     /**< Specifies the unique document item project identifier. */
        public Int32 lDocumentId;    /**< Specifies the unique document item identifier. */
    }

    public struct _AAEALINKAGE
    {
        public int lLinkageType;           /**< Specifies the type of linkage. See \ref aadmsdef_TablesColumnsSQLstatements_AttributeLinkageTypes "Attribute Linkage Types" for possible values. */
        //union
        //{
        public _AADOC_ITEM documentId;    /**< Specifies the structure containing the information about the document. This data member is valid only if the \em lLinkageType parameter is #AADMS_EALNK_DOCUMENT. */
        public int lEnvironmentId;      /**< Specifies the identifier of the environment. This data member is valid only if \em lLinkageType parameter is #AADMS_EALNK_ENVIRONMENT. */
        //};
    }

    // seems to be unused.  Taking out.  2011-01-12  DAB
    //[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
    //public unsafe struct _AADOCCREATE_PARAM
    //{
    //    /// <summary>
    //    /// [in] 0 or see Document Creation Mask Flags for more information.
    //    /// </summary>
    //    [MarshalAs(UnmanagedType.U4)]
    //    public UInt32 ulMask;

    //    /// <summary>
    //    /// [in] Specifies the project identifier.
    //    /// </summary>
    //    public int lProjectId;

    //    /// <summary>
    //    /// [out] Specifies the document identifier.
    //    /// </summary>
    //    public int lDocumentId;

    //    /// <summary>
    //    /// [in] Specifies the file type.
    //    /// </summary>
    //    public int lFileType;

    //    /// <summary>
    //    /// [in] Specifies the item type See Document Types for more information.
    //    /// </summary>
    //    public int lItemType;

    //    /// <summary>
    //    /// [in] Specifies the identifier of document's application.
    //    /// </summary>
    //    public int lApplicationId;

    //    /// <summary>
    //    /// [in] Specifies the identifier of document's department.
    //    /// </summary>
    //    public int lDepartmentId;

    //    /// <summary>
    //    /// [in] Pointer to a null-terminated string specifying the document file name.
    //    /// </summary>
    //    [MarshalAs(UnmanagedType.LPWStr)]
    //    public char* lpctstrFileName;

    //    /// <summary>
    //    /// [in] Pointer to a null-terminated string specifying the document name.
    //    /// </summary>
    //    [MarshalAs(UnmanagedType.LPWStr)]
    //    public char* lpctstrName;

    //    /// <summary>
    //    /// [in] Pointer to a null-terminated string specifying the document description.
    //    /// </summary>
    //    [MarshalAs(UnmanagedType.LPWStr)]
    //    public char* lpctstrDesc;

    //    /// <summary>
    //    /// [in] Specifies the document's storage identifier.
    //    /// </summary>
    //    public int lStorageId;

    //    /// <summary>
    //    /// [in] Pointer to a null-terminated string specifying the source file to be used as a base for the document's file. 
    //    /// </summary>
    //    [MarshalAs(UnmanagedType.LPWStr)]
    //    public char* lpctstrSourceFile;

    //    /// <summary>
    //    /// [in] Pointer to a null-terminated string specifying the document version string.
    //    /// </summary>
    //    [MarshalAs(UnmanagedType.LPWStr)]
    //    public char* lpctstrVersion;

    //    /// <summary>
    //    /// [in] Pointer to a string buffer receiving the file name of the created document.
    //    /// </summary>
    //    [MarshalAs(UnmanagedType.LPWStr)]
    //    public char* lptstrWorkingFile;

    //    /// <summary>
    //    /// [in] Specifies size of buffer lptstrWorkingFile in character symbols.
    //    /// </summary>
    //    public int lBufferSize;

    //    /// <summary>
    //    /// [in] Specifies creation flags.
    //    /// </summary>
    //    [MarshalAs(UnmanagedType.U4)]
    //    public UInt32 ulFlags;

    //    /// <summary>
    //    /// [in] Specifies the identifier of workspace profile.
    //    /// </summary>
    //    public int lWorkspaceProfileId;

    //    /// <summary>
    //    /// [in] Specifies if the new document will be checked in or will be left checked out by the current user. [in] Leave out flag
    //    /// </summary>
    //    public bool bLeaveOut;

    //    /// <summary>
    //    /// [in] Specifies the identifier of attribute if created.
    //    /// </summary>
    //    public int lAttributeId;

    //    /// <summary>
    //    /// [in] Specifies the GUID of project, if project ID not passed.
    //    /// </summary>
    //    public Guid guidProject;

    //    /// <summary>
    //    /// [out] Specifies the GUID of document.
    //    /// </summary>
    //    public Guid guidDocument;

    //    /// <summary>
    //    /// [out] MIME type
    //    /// </summary>
    //    [MarshalAs(UnmanagedType.LPWStr)]
    //    public char* pMimeType;

    //    /// <summary>
    //    /// [out] Handled relations (AARELATION_*)
    //    /// </summary>
    //    [MarshalAs(UnmanagedType.U4)]
    //    public UInt32 handledRelations;

    //}

    // seems to be unused.  Taking out.  2011-01-12 DAB
    //[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
    //public unsafe struct _AADOCMOVE_PARAM
    //{
    //    /// <summary>
    //    /// Reserved for the future use.
    //    /// </summary>
    //    [MarshalAs(UnmanagedType.U4)]
    //    public UInt32 ulMask;

    //    /// <summary>
    //    /// Specifies the project identifier of the source document.
    //    /// </summary>
    //    public int lSourceProjectId;

    //    /// <summary>
    //    /// Specifies the source document identifier.
    //    /// </summary>
    //    public int lSourceDocumentId;

    //    /// <summary>
    //    /// Specifies the project identifier of the target document.
    //    /// </summary>
    //    public int lTargetProjectId;

    //    /// <summary>
    //    /// Specifies the target document identifier.
    //    /// </summary>
    //    public int lTargetDocumentId;

    //    /// <summary>
    //    /// Pointer to a null-terminated string specifying the target document name.
    //    /// </summary>
    //    [MarshalAs(UnmanagedType.LPWStr)]
    //    public char* lpctstrName;

    //    /// <summary>
    //    /// Pointer to a null-terminated string specifying the target document description.
    //    /// </summary>
    //    [MarshalAs(UnmanagedType.LPWStr)]
    //    public char* lpctstrDesc;

    //    /// <summary>
    //    /// Pointer to a null-terminated string specifying the name of target document's file.
    //    /// </summary>
    //    [MarshalAs(UnmanagedType.LPWStr)]
    //    public char* lpctstrFileName;

    //    /// <summary>
    //    /// Pointer to a null-terminated string specifying the temporary directory to be used during copy or move operation. 
    //    /// </summary>
    //    [MarshalAs(UnmanagedType.LPWStr)]
    //    public char* lpctstrWorkdir;

    //    /// <summary>
    //    /// Specifies operation options.
    //    /// </summary>
    //    [MarshalAs(UnmanagedType.U4)]
    //    public UInt32 ulFlags;

    //}

    public delegate int HookFunction
    (int hookId,
        int hookType,
        int aParam1,
        int aParam2,
        ref int pResult
     );

    public delegate int DoumentHookFunction
    (int hookId,
        int hookType,
        ref AaDocumentsParam aParam1,
        int aParam2,
        ref int pResult
     );

    public delegate int GenericHookFunction
    (
        int hookId,
        int hookType,
        [In, Out] IntPtr ppDocsParam,
        int aParam2,
        ref int pResult
    );

    //new function wrapped by MDS
    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_SelectDatasourceStatistics();

    [DllImport("dmawin.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_ValidateMSFile
(
        string sFileName
);

    [DllImport("DGNPlatformSCUtils.dll", CharSet = CharSet.Unicode)]
    private static extern bool GetDgnFileType(string sFileName, ref int iFileType, ref int iMajorVersion, ref int iMinorVersion);

    [DllImport("DGNPlatformSCUtilsX64.dll", CharSet = CharSet.Unicode, EntryPoint = "GetDgnFileType")]
    private static extern bool GetDgnFileTypeX64(string sFileName, ref int iFileType, ref int iMajorVersion, ref int iMinorVersion);

    private static void FixDgnPlatformPath()
    {
        string sPath = PWWrapper.GetProjectWisePath();
        string sEnvPath = Environment.GetEnvironmentVariable("PATH");
        sPath = $"{sPath}\\bin\\DgnPlatform";

        if (!sEnvPath.ToLower().Contains(sPath.ToLower()))
            System.Environment.SetEnvironmentVariable("PATH", string.Format("{0};{1}", sPath, sEnvPath), EnvironmentVariableTarget.Process);
    }

    public static int GetDgnFileType(string sFileName)
    {
        FixDgnPlatformPath();

        try
        {
            int iFileType = 0, iMajorVersion = 0, iMinorVersion = 0;

            if (PWWrapper.Is64Bit())
            {
                if (GetDgnFileTypeX64(sFileName, ref iFileType, ref iMajorVersion, ref iMinorVersion))
                {
                    BPSUtilities.WriteLog($"Type: {iFileType}, Major: {iMajorVersion}, Minor: {iMinorVersion}");
                    return iFileType;
                }

                BPSUtilities.WriteLog($"Error: {iFileType}");
            }
            else
            {
                if (GetDgnFileType(sFileName, ref iFileType, ref iMajorVersion, ref iMinorVersion))
                {
                    BPSUtilities.WriteLog($"Type: {iFileType}, Major: {iMajorVersion}, Minor: {iMinorVersion}");
                    return iFileType;
                }

                BPSUtilities.WriteLog($"Error: {iFileType}");
            }
        }
        catch (Exception ex)
        {
            BPSUtilities.WriteLog($"{ex.Message}");
        }

        return 0;
    }

    [DllImport("dmawin.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaApi_DscTreeSelectSearchResultsItem
(
   IntPtr hWndDscTree,         /* i  Handle of the tree  */
   IntPtr hDataSource          /* i  Datasource handle   */
);

    [DllImport("dmawin.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_FindDocumentsToDocumentList
(
    IntPtr hWndDocList,           /* i  Document list window handle   */
    IntPtr hCriteriaBuf   /* i  Criteria to search for        */
);

    //BOOL aaApi_StringsToMonikers  ( LONG const   count,  
    //  HMONIKER *  monikers,  
    //  LPCWSTR *  strings,  
    //  DWORD  flags   
    // ) 

    //BOOL aaApi_MonikersToStrings  ( LONG_PTR const   monikerCount,  
    //  HMONIKER const *  monikers,  
    //  LPWSTR **  strings,  
    //  DWORD  flags   
    // ) 

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_MonikersToStrings([In]int lCount, [In]IntPtr[] pMonikers,
        [Out]StringBuilder[] sArrayMonikers, System.Int16 flags);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_StringsToMonikers([In]int lCount, [Out]IntPtr[] pMonikers,
        [In]string[] sArrayMonikers, System.Int16 flags);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_BuildMonikerStringByDocGuid(IntPtr hDatasource,
        ref Guid pDocGuid, ref IntPtr pMonikerStr);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_BuildMonikerStringByProjectGuid(IntPtr hDatasource,
        ref Guid pProjectGuid, ref IntPtr pMonikerStr);

    //aaApi_BuildMonikerStringByDocGuid  ( HDSOURCE  hDatasource,  
    // LPCGUID  pDocGuid,  
    // LPWSTR *  ppMonikerStr   
    //) 

    public static string GetDatasourceNameFromMonikerString(string sMoniker)
    {
        IntPtr[] pMonikerArray = new IntPtr[1];
        string[] sMonikerArray = new string[1];
        sMonikerArray[0] = sMoniker;

        try
        {
            if (aaApi_StringsToMonikers(1, pMonikerArray, sMonikerArray, (short)8)) // don't validate
            {
                try
                {
                    return aaApi_GetDatasourceNameFromMoniker(pMonikerArray[0]);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(string.Format("Error: {0}\n{1}", ex.Message, ex.StackTrace));

                    string sTransformMoniker = sMoniker.Replace('\\', '/');

                    sTransformMoniker = sTransformMoniker.Replace("pw://", "");

                    int iOccurence = sTransformMoniker.IndexOf("/Documents/");

                    if (iOccurence > 0)
                    {
                        string sDatasource = sTransformMoniker.Substring(0, iOccurence);

                        return sDatasource;
                    }
                }
            }
        }
        catch (Exception ex)
        {
            BPSUtilities.WriteLog("Error: {0}\n{1}", ex.Message, ex.StackTrace);
        }

        return string.Empty;
    }

    public static int GetProjectIdFromMonikerString(string sMoniker)
    {
        IntPtr[] pMonikerArray = new IntPtr[1];
        string[] sMonikerArray = new string[1];
        sMonikerArray[0] = sMoniker;

        int iProjectId = 0;

        try
        {
            if (aaApi_StringsToMonikers(1, pMonikerArray, sMonikerArray, (short)8)) // don't validate
            {
                Guid tempGuid = new Guid();

                // make sure it's big enough
                byte[] bArray = tempGuid.ToByteArray();

                IntPtr pProjGuid = PWWrapper.aaApi_GetProjectGuidFromMoniker(pMonikerArray[0]);

                Marshal.Copy(pProjGuid, bArray, 0, bArray.Length);

                Guid projGuid = new Guid(bArray);

                IntPtr hBuf = aaApi_GUIDSelectProjectDataBuffer(ref projGuid);

                if (hBuf != IntPtr.Zero)
                {
                    if (1 == aaApi_DmsDataBufferGetCount(hBuf))
                    {
                        iProjectId = aaApi_DmsDataBufferGetNumericProperty(hBuf, (int)ProjectProperty.ID, 0);
                    }

                    aaApi_DmsDataBufferFree(hBuf);
                }
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(string.Format("Error: {0}\n{1}", ex.Message, ex.StackTrace));
        }

        return iProjectId;
    }

    public static bool GetDocumentIdsFromMonikerString(string sMoniker, ref int iProjectId, ref int iDocumentId)
    {
        IntPtr[] pMonikerArray = new IntPtr[1];
        string[] sMonikerArray = new string[1];
        sMonikerArray[0] = sMoniker;

        bool bRetVal = false;

        try
        {
            if (aaApi_StringsToMonikers(1, pMonikerArray, sMonikerArray, (short)8)) // don't validate
            {
                Guid tempGuid = new Guid();

                // make sure it's big enough
                byte[] bArray = tempGuid.ToByteArray();

                IntPtr pDocGuid = PWWrapper.aaApi_GetDocumentGuidFromMoniker(pMonikerArray[0]);

                Marshal.Copy(pDocGuid, bArray, 0, bArray.Length);

                Guid docGuid = new Guid(bArray);

                PWWrapper.AaDocItem docItem = new PWWrapper.AaDocItem();
                Guid[] guids = new Guid[1];

                try
                {
                    guids[0] = docGuid;

                    if (PWWrapper.aaApi_GetDocumentIdsByGUIDs(1, guids, ref docItem))
                    {
                        iProjectId = docItem.lProjectId;
                        iDocumentId = docItem.lDocumentId;
                        bRetVal = true;
                    }
                }
                finally
                {
                }
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(string.Format("Error: {0}\n{1}", ex.Message, ex.StackTrace));
        }

        return bRetVal;
    }

    public static string GetMonikerStringFromDocumentGuidString(string sDocGuid)
    {
        try
        {
            Guid docGuid = new Guid(sDocGuid);

            return GetMonikerStringFromDocumentGuid(docGuid);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(string.Format("Error: {0}\n{1}", ex.Message, ex.StackTrace));
        }

        return string.Empty;
    }

    public static string GetMonikerStringFromProjectGuidString(string sProjGuid)
    {
        try
        {
            Guid projGuid = new Guid(sProjGuid);

            return GetMonikerStringFromProjectGuid(projGuid);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(string.Format("Error: {0}\n{1}", ex.Message, ex.StackTrace));
        }

        return string.Empty;
    }

    public static string GetMonikerStringFromDocumentGuid(Guid docGuid)
    {
        StringBuilder sbMonikerReturn = new StringBuilder();

        try
        {
            IntPtr pMonikerString = IntPtr.Zero;

            if (aaApi_BuildMonikerStringByDocGuid(aaApi_GetActiveDatasource(), ref docGuid, ref pMonikerString))
            {
                string sMonikerString = Marshal.PtrToStringUni(pMonikerString);
                sbMonikerReturn.Append(sMonikerString);
                aaApi_Free(pMonikerString);
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(string.Format("Error: {0}\n{1}", ex.Message, ex.StackTrace));
        }

        return sbMonikerReturn.ToString();
    }

    public static string GetMonikerStringFromProjectGuid(Guid projGuid)
    {
        StringBuilder sbMonikerReturn = new StringBuilder();

        try
        {
            IntPtr pMonikerString = IntPtr.Zero;

            if (aaApi_BuildMonikerStringByProjectGuid(aaApi_GetActiveDatasource(), ref projGuid, ref pMonikerString))
            {
                string sMonikerString = Marshal.PtrToStringUni(pMonikerString);
                sbMonikerReturn.Append(sMonikerString);
                aaApi_Free(pMonikerString);
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(string.Format("Error: {0}\n{1}", ex.Message, ex.StackTrace));
        }

        return sbMonikerReturn.ToString();
    }

    public static string GetMonikerStringFromProjectId(int iProjectId)
    {
        StringBuilder sbMonikerReturn = new StringBuilder();

        string sProjectGuid = GetProjectGuidStringFromId(iProjectId);

        try
        {
            Guid projGuid = new Guid(sProjectGuid);

            IntPtr pMonikerString = IntPtr.Zero;

            if (aaApi_BuildMonikerStringByProjectGuid(aaApi_GetActiveDatasource(), ref projGuid, ref pMonikerString))
            {
                string sMonikerString = Marshal.PtrToStringUni(pMonikerString);
                sbMonikerReturn.Append(sMonikerString);
                aaApi_Free(pMonikerString);
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(string.Format("Error: {0}\n{1}", ex.Message, ex.StackTrace));
        }

        return sbMonikerReturn.ToString();
    }

    public static string GetMonikerStringFromDocumentIds(int iProjectId, int iDocumentId)
    {
        StringBuilder sbMonikerReturn = new StringBuilder();

        string sDocGuid = GetGuidStringFromIds(iProjectId, iDocumentId);

        try
        {
            Guid docGuid = new Guid(sDocGuid);

            IntPtr pMonikerString = IntPtr.Zero;

            if (aaApi_BuildMonikerStringByDocGuid(aaApi_GetActiveDatasource(), ref docGuid, ref pMonikerString))
            {
                string sMonikerString = Marshal.PtrToStringUni(pMonikerString);
                sbMonikerReturn.Append(sMonikerString);
                aaApi_Free(pMonikerString);
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(string.Format("Error: {0}\n{1}", ex.Message, ex.StackTrace));
        }

        return sbMonikerReturn.ToString();
    }

    [DllImport("dmawin.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SetAttributeSheetSelection(int iProjectId, int iDocumentId, int iAttrRecId);

    [DllImport("dmawin.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_GetCurrentAttributeSheetDocument(ref int iProjectId, ref int iDocumentId, ref int iAttrRecId);

    [DllImport("dmawin.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_DocListSetMenuType(IntPtr hWndDocList, bool bShowMenu, bool bShowModify, bool bShowWorkflow, bool bShowView, bool bShowCustom);

    [DllImport("dmawin.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaApi_GetMainDocumentList();

    [DllImport("dmawin.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaApi_GetDocListMoniker(IntPtr docListP);

    [DllImport("dmawin.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaApi_GetDscTreeMoniker(IntPtr dscTreeP);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaApi_GetDocumentGuidFromMoniker(IntPtr hMoniker);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern string aaApi_GetDatasourceNameFromMoniker(IntPtr hMoniker);

    [DllImport("dmawin.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_DocListSelectDocument(IntPtr hWndList, int iProjectId, int iDocumentId, int iAttrId);

    [DllImport("dmawin.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_DocListSetProject(IntPtr hWndList, int iProjectId);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaApi_GetProjectGuidFromMoniker(IntPtr hMoniker);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectParentProject(int projectId);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GetDocumentCount(int lProjectId);


    [DllImport("dmawin.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectDocumentDlg(IntPtr hWndParent, string lpctstrTitle, int lApplicationId, ref int lplProjectId, ref int lplDocumentId);

    #region hide
#if false

    // this works, but because it's unsafe, taking out of PWWrapper.  Put it directly in app.  2011-01-12  DAB
    [DllImport("dmawin.dll", CharSet = CharSet.Unicode, EntryPoint = "aaApi_SelectDocumentsDlg")]
    private static unsafe extern int _aaApi_SelectDocumentsDlg(int hWndParent,
        string lpctstrTitle, int ProjectId,
        int ApplicationId, ref int plRealCount, ref IntPtr ppAaDocItem);

    public static unsafe int aaApi_SelectDocumentsDlg(int hWndParent, string lpctstrTitle, int ProjectId,
        int ApplicationId, ref int plRealCount, ref PWWrapper.AaDocItem[] aaDocItemArray)
    {
        IntPtr docs = IntPtr.Zero;

        int returnVal = _aaApi_SelectDocumentsDlg(hWndParent, lpctstrTitle,
            ProjectId, ApplicationId, ref plRealCount, ref docs);

        if (docs != IntPtr.Zero)
        {
            aaDocItemArray = new PWWrapper.AaDocItem[plRealCount];
            int pDocs = docs.ToInt32();

            int iSize = Marshal.SizeOf(new PWWrapper.AaDocItem());

            for (int i = 0; i < plRealCount; i++)
            {
                aaDocItemArray[i] = new PWWrapper.AaDocItem();

                // aaDocItemArray[i].lProjectId = Marshal.ReadInt32(new IntPtr(pDocs + (i * 532)));
                // aaDocItemArray[i].lDocumentId = Marshal.ReadInt32(new IntPtr(pDocs + (i * 532) + 4));

                aaDocItemArray[i].lProjectId =
                    Marshal.ReadInt32(new IntPtr(pDocs + (i * iSize)));
                aaDocItemArray[i].lDocumentId =
                    Marshal.ReadInt32(new IntPtr(pDocs + (i * iSize) + 4));
            }

            PWWrapper.aaApi_Free(docs);
        }

        return returnVal;
    }
    
#endif
    #endregion

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GetApplicationNumericProperty(ApplicationProperty lPropertyId, int lIndex);

    // this appears to be used in PCMV8iNewDrawing, but truthfully, I do not remember writing.  Will take out for now.  2011-01-12  DAB
    //[DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    //public static unsafe extern bool aaApi_CreateEnvAttr(int lTableId, _AAEALINKAGE* lpLinkage, ref int lplAttrRecId);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_StrToNumber(string lpctstrNumber);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_ModifyProject2(ref VaultDescriptor vaultDescriptor);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_Initialize(int init);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_Uninitialize();


    [DllImport("dmscli.dll", EntryPoint = "aaApi_Login", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_Login(int iDSType, string lptstrDataSource,
        string lpctstrUsername, string lpctstrPassword, string lpctstrSchema);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_LoginWithSecurityToken(string dataSource, string securityToken, bool asAdmin, string hostname, long[] productIds);


    internal static Hashtable sessionHandles = new Hashtable();

    /// *****************************************************************************
    /// <summary>Use MD5 to create a one-way hash token for the session parameters</summary>                
    /// @author                                             AdamKlatzkin 02/03
    /// ****************************************************************************
    private static string GetCredentialHash(PWWrapper.DataSourceType lDSType, string lptstrDataSource,
        string lpctstrUsername, string lpctstrPassword, string lpctstrSchema)
    {
        string loginToken = lptstrDataSource + "_" + lpctstrUsername + "_" + lpctstrPassword + "_" + lpctstrSchema + "_" + lDSType.ToString();
        byte[] inputBytes = Encoding.Unicode.GetBytes(loginToken);
        byte[] hash = MD5.Create().ComputeHash(inputBytes);
        return Convert.ToBase64String(hash);
    }


    /// *****************************************************************************
    /// <summary>Same as aaApi_Login except if bReuseExistingConnection is set to 
    /// false a login will occur regardless of whether a session already exists for the
    /// given parameters.</summary>                
    /// @author                                             AdamKlatzkin 02/03
    /// ****************************************************************************
    public static bool aaApi_Login(PWWrapper.DataSourceType lDSType, string lptstrDataSource,
        string lpctstrUsername, string lpctstrPassword, string lpctstrSchema, bool bReuseExistingConnection)
    {
        IntPtr handle = IntPtr.Zero;
        string hash = GetCredentialHash(lDSType, lptstrDataSource, lpctstrUsername, lpctstrPassword, lpctstrSchema);
        if (bReuseExistingConnection)
        {
            if (sessionHandles.Contains(hash))
                handle = (IntPtr)sessionHandles[hash];
        }
        if (handle == IntPtr.Zero)
        {
            bool retVal = aaApi_Login((int)lDSType, lptstrDataSource, lpctstrUsername, lpctstrPassword, lpctstrSchema);
            if (retVal)
            {
                handle = aaApi_GetActiveDatasource();
                sessionHandles[hash] = handle;
            }
            return retVal;
        }
        else
        {
            if (aaApi_ActivateDatasourceByHandle(handle) == IntPtr.Zero)
            {
                bool retVal = aaApi_Login((int)lDSType, lptstrDataSource, lpctstrUsername, lpctstrPassword, lpctstrSchema);
                if (retVal)
                {
                    handle = aaApi_GetActiveDatasource();
                    sessionHandles[hash] = handle;
                }
                return retVal;
            }
            return true;
        }
    }


    [DllImport("dmscli.dll", EntryPoint = "aaApi_Logout", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_Logout(string lptstrDataSource);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GetDocumentNumericProperty(DocumentProperty PropertyId, int lIndex);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern UInt64 aaApi_GetDocumentUint64Property(DocumentProperty PropertyId, int lIndex);

    [DllImport("dmscli.dll", EntryPoint = "aaApi_GetDocumentGuidProperty", CharSet = CharSet.Unicode)]
    private static extern IntPtr intPtr_aaApi_GetDocumentGuidProperty(DocumentProperty lPropertyId, int lIdxRow);

    public static Guid aaApi_GetDocumentGuidProperty(DocumentProperty PropertyId, int lIndex)
    {
        return (Guid)Marshal.PtrToStructure(intPtr_aaApi_GetDocumentGuidProperty(PropertyId, lIndex), Type.GetType("System.Guid"));
    }

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern UInt64 aaApi_GetDatasourceStatisticsNumericProperty(DatasourceStatisticsProperty PropertyId);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectLinkDataByObject(int lTableId, ObjectTypeForLinkData lItemType, int lItemId1, int lItemId2,
        string lpctstrWhere, ref int lplColumnCount, int[] lplColumnIds, LinkDataSelectFlags ulFlags);


    [DllImport("dmscli.dll", EntryPoint = "aaApi_GetEnvTableInfoByProject", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_GetEnvTableInfoByProject
        (
        int lProjectId,         /* i  Project  id                    */
        ref int lplEnvironmentId,   /* o  Environment id (or NULL)       */
        ref int lplTableId,         /* o  Table id (or NULL)             */
        ref int lplIdColumnId       /* o  Identifier column id (or NULL) */
        );

    [DllImport("dmscli.dll", EntryPoint = "aaApi_GetEnvTableIdColumnId", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GetEnvTableIdColumnId(int lTableId);

    [DllImport("dmscli.dll", EntryPoint = "aaApi_SelectLinks", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectLinks
        (
        int lProjectId,         /* i  Project  id                    */
        int lDocumentId
        );

    [DllImport("dmscli.dll", EntryPoint = "aaApi_SelectLinksByAttr", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectLinksByAttr
        (
        int iTableId,         /* i  Table  id                    */
        int iColumnId,
        string sAttrVal
        );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_FreeLinkDataUpdateDesc();

    /// <summary>
    /// Get the basic datsource type. Returns 1 for Oracle, 2 for SQL Server, 0 for Unknown.
    /// </summary>
    /// <returns>1 for Oracle, 2 for SQL Server, 0 for Unknown</returns>
    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GetActiveDatasourceType();


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_GetActiveDatasourceName
(
StringBuilder lptstrName,    /* o  Datasource name                    */
int lNameSize      /* i  lptstrName size in characters      */
);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_UpdateLinkDataColumnValue(int tableID, int columnID, string columnValue);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_UpdateLinkData(int tableID, int columnID, string columnValue);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_UpdateEnvAttr(int tableID, int attrRecId);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GetLinkDataColumnNumericProperty(LinkDataProperty property, int index);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GetLinkDataDataBufferNumericColumnValue(IntPtr hDataBuffer, int rowIndex, int columnIndex);

    [DllImport("dmscli.dll", EntryPoint = "aaApi_GetLinkStringProperty", CharSet = CharSet.Unicode)]
    private static extern IntPtr unsafe_aaApi_GetLinkStringProperty(LinkProperty propertyID, int index);

    public static string aaApi_GetLinkStringProperty(LinkProperty propertyID, int index)
    {
        return Marshal.PtrToStringUni(unsafe_aaApi_GetLinkStringProperty(propertyID, index));
    }

    [DllImport("dmscli.dll", EntryPoint = "aaApi_SelectLinkData", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectLinkData
    (
       int lTableId,       /* i  Link table  number               */
       int lColumnId,      /* i  Link column number (0 for all)   */
       string lpctstrValue,   /* i  Link column value (NULL for all) */
       ref int lplColumns      /* o  Number of columns in link table  */
    );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaApi_AddViewColumnToDataBuffer(IntPtr hColumnBuffer, ref Guid fieldType, string sContext,
        int dataType, string fieldName, int alignment, int width);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_CreateView(ref Guid viewType, int iUserId, string sViewName, string sContext, IntPtr hColumnBuffer, ref int iViewId);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_DeleteView(int viewId);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern void aaApi_SetDefaultViewCacheRefresh();

    public static string aaApi_GetDocumentStringProperty(DocumentProperty PropertyId, int Index)
    {
        return Marshal.PtrToStringUni(unsafe_aaApi_GetDocumentStringProperty(PropertyId, Index));
    }

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_GetDocumentFileName(int iProjectId, int iDocumentId, StringBuilder sbLocalFilePath,
        int iBufferSize);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_GetDocumentFileSize64(int iProjectId, int iDocumentId, ref UInt64 ulFileSize);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_IsDocumentCheckedOutToMe(int lProjectId,
        int lDocumentId, ref bool bIsOutToMe);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_RefreshDocumentServerCopy(int lProjectId,
        int lDocumentId);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_CheckInDocument(int lProjectId,
        int lDocumentId);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectDocument(int ProjectId, int lDocumentId);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectDocumentsByNameProp(int vaultID, string fileName,
        string name, string description, string version);

    // aaApi_GetProjectNamePath has been depreciated, use aaApi_GetProjectNamePath2() instead
    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_GetProjectNamePath(int ProjectId, bool UseDesc, char tchSeparator,
        StringBuilder StringBuffer, int BufferSize);

    // dww - for data migration tools
    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_GetProjectNamePath2(int ProjectId, bool UseDesc, char tchSeparator,
        StringBuilder StringBuffer, int BufferSize);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_GetDocumentNamePath(int ProjectId, int DocId, bool UseDesc, char tchSeparator,
        StringBuilder StringBuffer, int BufferSize);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectProject(int lProjectId);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectTopLevelProjects();


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectChildProjects(int iParentId);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectAllProjects();

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GetEnvIdByProject(int lProjectId);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectProjectChain(int lProjectFrom, int lProjectTo);

    // dww 2013-09-30
    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_SelectSubProjectChainByNameDataBuffer(
        uint ulFlags,
        string lpctstrPath,
        string lpctstrDoc,
        string lpcstrVersion,
        ref int lplProjId,
        ref int lplDocId,
        ref IntPtr lphResultSet
        );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectAllApplications();

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaApi_SelectActionForApplication(int iApplicationId, ref Guid guidActionType);

    public static Guid[] ApplicationActionTypes = {new Guid("AC08DF6B-F420-44e8-95E0-8142CF2288C5"),
                                        new Guid("BAB502F2-85B0-403c-BFF5-C314E0783818"),
                                        new Guid("DBB7BD29-F1DC-4a7b-8B12-F585C4299D7E"),
                                        new Guid("FAF4FB46-1871-4834-9C6D-C991428DB202"),
                                        new Guid("ED52A96D-8D2E-4c54-A54A-8BDA68A8B4A6")};

    public enum ApplicationActionTypesIndices : int
    {
        DMS_APPLACTION_OPEN = 0,
        DMS_APPLACTION_VIEW = 1,
        DMS_APPLACTION_REDLINE = 2,
        DMS_APPLACTION_PRINT = 3,
        DMS_APPLACTION_SCANREFS = 4
        //{
        //    Guid DMS_APPLACTION_OPEN = new Guid("AC08DF6B-F420-44e8-95E0-8142CF2288C5");
        //    Guid DMS_APPLACTION_VIEW = new Guid("BAB502F2-85B0-403c-BFF5-C314E0783818");
        //    Guid DMS_APPLACTION_REDLINE = new Guid("DBB7BD29-F1DC-4a7b-8B12-F585C4299D7E");
        //    Guid DMS_APPLACTION_PRINT = new Guid("FAF4FB46-1871-4834-9C6D-C991428DB202");
        //    Guid DMS_APPLACTION_SCANREFS = new Guid("ED52A96D-8D2E-4c54-A54A-8BDA68A8B4A6");
        //}
    }

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaApi_SelectApplicationActions();

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaApi_SelectUserApplicationActions(int iApplicationId, int iUserId);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_StartPartnerApplication(string sApplCmd,
        string sApplArgs,
        string sFileName,
        int lProjectId,
        int lDocumentId,
        int lSetId,
        bool bAskCheckIn
        );

    [DllImport("dmscli.dll", EntryPoint = "aaApi_GetProjectStringProperty", CharSet = CharSet.Unicode)]
    private static extern IntPtr unsafe_aaApi_GetProjectStringProperty(ProjectProperty PropertyId, int lIndex);


    public static string aaApi_GetProjectStringProperty(ProjectProperty PropertyId, int lIndex)
    {
        return Marshal.PtrToStringUni(unsafe_aaApi_GetProjectStringProperty(PropertyId, lIndex));
    }

    [DllImport("dmscli.dll", EntryPoint = "aaApi_GetLinkDataColumnStringProperty", CharSet = CharSet.Unicode)]
    private static extern IntPtr unsafe_aaApi_GetLinkDataColumnStringProperty(LinkDataProperty propertyID, int index);


    public static string aaApi_GetLinkDataColumnStringProperty(LinkDataProperty propertyID, int index)
    {
        return Marshal.PtrToStringUni(unsafe_aaApi_GetLinkDataColumnStringProperty(propertyID, index));
    }


    [DllImport("dmscli.dll", EntryPoint = "aaApi_GetLinkDataColumnValue", CharSet = CharSet.Unicode)]
    private static extern IntPtr unsafe_aaApi_GetLinkDataColumnValue(int lRowIndex, int lColumnIndex);


    public static string aaApi_GetLinkDataColumnValue(int lRowIndex, int lColumnIndex)
    {
        return Marshal.PtrToStringUni(unsafe_aaApi_GetLinkDataColumnValue(lRowIndex, lColumnIndex));
    }


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GetProjectNumericProperty(ProjectProperty PropertyId, int lIndex);

    [DllImport("dmawin.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaApi_ViewGetLastViewName();

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_GiveOutDocument
    (
       int lProjectNo,        /* i  Project number              */
       int lDocumentId,       /* i  Document number             */
       String lpctstrWorkdir,    /* i  Working directory           */
       String lptstrFileName,    /* o  File name with full path    */
       int lBufferSize        /* i  Buffer size for file name   */
    );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GetActiveDatasourceNativeType();

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GetActiveInterface();

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_ExecuteSqlStatement(string description);


    // dww - added so that I can check if a table or view exists in a datasource

    public enum DatabaseTableType : int
    {
        Table = 0,
        View = 1
    };

    [DllImport("dmscli.dll", EntryPoint = "aaApi_DoesTableExist", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_DoesTableExist
    (
        string tableName,
        DatabaseTableType tableType,
        ref bool tableExists
        );

    [DllImport("dmscli.dll", EntryPoint = "aaApi_SqlSelectGetData", CharSet = CharSet.Unicode)]
    private static extern IntPtr unsafe_aaApi_SqlSelectGetData(int iRow, int iColumn);


    public static string aaApi_SqlSelectGetData(int iRow, int iColumn)
    {
        return Marshal.PtrToStringUni(unsafe_aaApi_SqlSelectGetData(iRow, iColumn));
    }


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectDatasources();


    [DllImport("dmscli.dll", EntryPoint = "aaApi_GetDatasourceFullName", CharSet = CharSet.Unicode)]
    private static extern IntPtr unsafe_aaApi_GetDatasourceFullName(int index);

    public static string aaApi_GetDatasourceFullName(int index)
    {
        return Marshal.PtrToStringUni(unsafe_aaApi_GetDatasourceFullName(index));
    }


    [DllImport("dmscli.dll", EntryPoint = "aaApi_GetDatasourceName", CharSet = CharSet.Unicode)]
    private static extern IntPtr unsafe_aaApi_GetDatasourceName(int index);

    public static string aaApi_GetDatasourceName(int index)
    {
        return Marshal.PtrToStringUni(unsafe_aaApi_GetDatasourceName(index));
    }


    [DllImport("dmscli.dll", EntryPoint = "aaApi_GetDatasourceInternalName", CharSet = CharSet.Unicode)]
    private static extern IntPtr unsafe_aaApi_GetDatasourceInternalName(int index);

    public static string aaApi_GetDatasourceInternalName(int index)
    {
        return Marshal.PtrToStringUni(unsafe_aaApi_GetDatasourceInternalName(index));
    }

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_GUIDFetchDocumentFromServer(FetchDocumentFlags flags,
        ref Guid guid, string sWorkingDir,
        StringBuilder StringBuffer, int BufferSize);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_GUIDCopyOutDocument(ref Guid guid, string sWorkingDir,
        StringBuilder StringBuffer, int BufferSize);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_GUIDCheckOutDocument(ref Guid guid, string sWorkingDir,
        StringBuilder StringBuffer, int BufferSize);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_GUIDCheckInDocument(ref Guid guid);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_CopyOutDocument(int lProjectNo, int lDocumentId,
        string lpctstrWorkdir, StringBuilder lptstrFileName, int lBufferSize);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_CheckOutDocument(int lProjectNo, int lDocumentId,
        string lpctstrWorkdir, StringBuilder lptstrFileName, int lBufferSize);

    [DllImport("dmawin.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_OpenDocument(int lProjectNo, int lDocumentId, bool bReadOnly);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_FetchDocumentFromServer
(
   FetchDocumentFlags ulFlags,           /* i  Flags (AADMS_DOCFETCH_*)     */
   int lProjectId,        /* i  Project number               */
   int lDocumentId,       /* i  Document number              */
   string lpctstrWorkdir,    /* i  Working directory            */
   StringBuilder lptstrFileName,    /* o  File name with full path     */
   int lBufferSize        /* i  Buffer size for file name    */
);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_GUIDPurgeDocumentCopy(ref Guid guid, int iUserID);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SqlSelect(string sqlStatement, IntPtr columnBind, ref int numColumnsSelected);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GUIDSelectDocument(ref Guid guid);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaApi_GUIDSelectDocumentDataBuffer(ref Guid guid);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_ChklSetGetDocGuidFromFileName(string pFileName, ref Guid pDocGuid);

    // this does not seem to work right anymore
    // needs work on marshalling - use aaApi_ChklSetGetDocGuidFromFileName(sFileName, ref docGuid)
    [DllImport("dmscli.dll", CharSet = CharSet.Ansi)]
    private static extern int aaApi_GetGuidsFromFileName([In, Out] ref Guid[] docGuids, [In, Out]ref int iNumGuids,
        [In] string sFileName, [In] int iValidateWithChkl);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaApi_SelectDocumentDataBufferVersions
    (
       int lProjectId,          /* i  Project number (must exist) */
       int lDocumentId          /* i  Document number             */
    );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaApi_SelectDocumentDataBufferByNameProp
(
   int lProjectId,       /* i  Project number (must exist)    */
   string sFileName,  /* i  File name to search for        */
   string sName,      /* i  Document name to search        */
   string sDesc,      /* i  Document description to search */
   string sVersion    /* i  Document version to search     */
);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectProjectsFromBranch(int iParentId,
       string lpctstrCode,    /* i  Project code to search for        */
       string lpctstrName,    /* i  Project name to search for        */
       string lpctstrDesc,    /* i  Project description to search for */
       string lpctstrVersion  /* i  Project version to search for     */
    );


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_CreateProject(ref int createdVaultID, int parentID,
        int storageID, int managerID, VaultType type, int workflowID,
        int workspaceProfileID, int copyAccessFromProject, string name, string description);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_CreateStorage(int lStorageId, string lpctstrName,
        string lpctstrDesc, string lpctstrNode, string lpctstrPath, string lpctstrProto);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_NewDocumentVersion(NewVersionCreationFlags ulFlags,
        int vaultID, int documentID, string version,
        string comment, ref int versionDocId);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_GUIDNewDocumentVersion
(
NewVersionCreationFlags ulFlags,
ref Guid pDocGuid,
string docVersion,
string comment,
ref Guid pVersionDocGuid
);

    [DllImport("dmawin.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_CreateDocumentVersionsFromSource
(
   IntPtr hDSTarget,    /* i  Target datasource handle    */
   [In]AaDocItem[] arrTrgtDocs,  /* i  Array of target documents   */
   IntPtr hDSSource,    /* i  Source datasource handle    */
   [In]AaDocItem[] arrSrcDocs,   /* i  Array of source documents   */
   int lArrayLen,    /* i  Length of arrays            */
   string lpctstrFrmt,  /* i  Version string format       */
   CreateVersionsFromSourceFlags ulFlags,      /* i  Additional flags (AARULEO_*)*/
   IntPtr fnCBack,      /* i  Pointer to callback function*/
   int aaCBackData   /* i  Callback function parameter */
);

    [DllImport("dmscli.dll", EntryPoint = "aaApi_CopyDocuments", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_CopyDocuments(
        int documentCount,
        IntPtr hDSSource,    /* i  Source datasource handle    */
        [In]AaDocItem[] arrSrcDocs,   /* i  Array of source documents   */
        IntPtr hDSTarget,    /* i  Target datasource handle    */
        [In, Out]AaDocItem[] arrTrgtDocs,  /* i  Array of target documents   */
        string workdir,
        [In, Out] string[] fileNames,
        [In, Out] string[] newNames,
        [In, Out] string[] newDescriptions,
        DocumentCopyFlags flags,
        IntPtr callback,
        IntPtr userParam
        );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_CreateDocument(ref int documentID, int vaultID,
        int storageID, int fileType, DocumentType itemType, int applicationID,
        int departmentID, int workspaceProfileID, string sourceFilePath, string fileName, string name,
        string description, string version, bool leaveCheckedOut, DocumentCreationFlag creationFlags,
        StringBuilder workingFile, int workingFileBufferSize, ref int attributeID);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_DocumentGenerateName(long lProjectId, StringBuilder lptstrDocName, int iBufferSize);

    //[DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    // public static extern bool aaApi_DocumentGenNameWithPrefix(long lProjectId, string lpctstrPrefix, StringBuilder lptstrDocName, int iBufferSize);

    //[DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    //public static extern bool aaApi_DocumentGenNameWithPrefix
    //(
    //    int lProjectId,
    //    string lpctstrPrefix,
    //    [MarshalAs(UnmanagedType.LPWStr)]
    //        StringBuilder lptstrDocName,
    //    int iBufferSize
    //    ); 

    //public static extern bool aaApi_DocumentGenNameWithPrefix(long lProjectId, [MarshalAs(UnmanagedType.LPStr)] string lpctstrPrefix, [MarshalAs(UnmanagedType.LPStr)]ref string lptstrDocName, int iBufferSize);

    //[DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    ////public static extern bool aaApi_DocumentGenFileNameWithPrefix(long lProjectId, string lpctstrPrefix, [MarshalAs(UnmanagedType.LPWStr)] StringBuilder lptstrFileName, int iBufferSize);
    //public static extern bool aaApi_DocumentGenFileNameWithPrefix
    //    (
    //    int lProjectId,
    //    string lpctstrPrefix,
    //    [MarshalAs(UnmanagedType.LPWStr)]
    //        StringBuilder lptstrDocName,
    //    int iBufferSize
    //    ); 


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaApi_ActivateDatasourceByHandle(IntPtr dsHandle);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_LogoutByHandle(IntPtr dsHandle);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaApi_GetActiveDatasource();

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_AddDocumentFile(int vaultID, int documentID, string fileName);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_ChangeDocumentFile(int vaultID, int documentID, string fileName);

    // dww
    // changed opFlags type and fileMIME type to implement in/out for fileMIME
    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_ChangeDocumentFile4(int vaultID, int documentID, DocumentFileOp opFlags, string sourcefile, string newfilename, StringBuilder fileMIME);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_RefreshDatasourceStatistics();

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectUser(int lUserId);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectUsersByProp(string name, string description, string email);

    [Flags]
    public enum SetUserIdentityFlags : uint
    {
        DMS_USERIDENTITY_SETF_NONE = 0,
        DMS_USERIDENTITY_SETF_DISASSOCIATE = 1,
        DMS_USERIDENTITY_SETF_CLEAR_PROVIDER_DATA = 2
    }

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_SetUserIdentity(int iUserId, SetUserIdentityFlags userIdentityFlags, string sIdentity);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GetUserId(int index);

    public static string GetUserName(int iUserId)
    {
        if (1 == PWWrapper.aaApi_SelectUser(iUserId))
            return PWWrapper.aaApi_GetUserStringProperty(UserProperty.Name, 0);
        return string.Empty;
    }

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GetApplicationId(int index);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GetCurrentUserId();

    [DllImport("dmscli.dll", EntryPoint = "aaApi_GetWorkingDirectory", CharSet = CharSet.Unicode)]
    private static extern IntPtr unsafe_aaApi_GetWorkingDirectory();

    public static string aaApi_GetWorkingDirectory()
    {
        return Marshal.PtrToStringUni(unsafe_aaApi_GetWorkingDirectory());
    }

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_CreateDepartment(ref int lDepartmentId, string name, string description);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_DeleteDepartmentById(int lDepartmentId);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_DeleteDepartmentByName(string name);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GetDepartmentCount();

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GetDepartmentId(int index);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GetDepartmentNumericProperty(int lPropertyId, int lIndex);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GetDepartmentPropertyLength(int lPropertyId);

    [DllImport("dmscli.dll", EntryPoint = "aaApi_GetDepartmentStringProperty", CharSet = CharSet.Unicode)]
    private static extern IntPtr unsafe_aaApi_GetDepartmentStringProperty(DepartmentProperty PropertyId, int Index);

    public static string aaApi_GetDepartmentStringProperty(DepartmentProperty PropertyId, int Index)
    {
        return Marshal.PtrToStringUni(unsafe_aaApi_GetDepartmentStringProperty(PropertyId, Index));
    }

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_ModifyDepartment(int lDepartmentId, string name, string description);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectAllDepartments();

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectDepartment(int lDepartmentId);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectDepartmentsForProject(int lDepartmentId);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectEnv(int lEnvironmentId);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_DeleteEnv(int lEnvironmentId);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_DeleteEnvAttrDefs(int lEnvironmentId, int iTableId, int iColumnId);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_DeleteEnvCodeDef(int lEnvironmentId);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_InitializeEnvCodeDefUpdate();

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_UpdateEnvCodeDef(int iEnvironmentID);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_AddEnvCodeDefUpdateField(int iTableId, int iColumnId, int iCodeType, int iSerialType,
        int iParams, int iOrderNo, string sConnector);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_DeleteEnvTriggerDefs(int lEnvironmentId, int iTableId, int iColumnId, int iTriggerColumn);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GetEnvId(int lIndex);

    [DllImport("dmsgen.dll", CharSet = CharSet.Unicode)]
    public static extern Boolean aaApi_RemoveAllErrors();

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GetEnvNumericProperty
        (
        int lPropertyId, /* i  Property id              */
        int lIndex       /* i  Index of selected enviroment  */
        );


    [DllImport("dmscli.dll", EntryPoint = "aaApi_GetEnvStringProperty", CharSet = CharSet.Unicode)]
    private static extern IntPtr unsafe_aaApi_GetEnvStringProperty
        (
        int lPropertyId,   /* i  Property id              */
        int lIndex         /* i  Index of selected environment  */
        );

    public static string aaApi_GetEnvStringProperty
        (int lPropertyId,   /* i  Property id              */
        int lIndex         /* i  Index of selected environment  */
        )
    {
        return Marshal.PtrToStringUni(unsafe_aaApi_GetEnvStringProperty(lPropertyId, lIndex));
    }

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GetEnvNumericProperty
        (
        EnvironmentProperty lPropertyId, /* i  Property id              */
        int lIndex       /* i  Index of selected enviroment  */
        );


    [DllImport("dmscli.dll", EntryPoint = "aaApi_GetEnvStringProperty", CharSet = CharSet.Unicode)]
    private static extern IntPtr unsafe_aaApi_GetEnvStringProperty
        (
        EnvironmentProperty lPropertyId,   /* i  Property id              */
        int lIndex         /* i  Index of selected environment  */
        );

    public static string aaApi_GetEnvStringProperty
        (EnvironmentProperty lPropertyId,   /* i  Property id              */
        int lIndex         /* i  Index of selected environment  */
        )
    {
        return Marshal.PtrToStringUni(unsafe_aaApi_GetEnvStringProperty(lPropertyId, lIndex));
    }

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaOApi_CopyClass
    (
        IntPtr pClassToCopy,                /* i Class to Copy  */
        ref IntPtr ppCopiedClass                /* i  Interface id    */
    );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaOApi_SetClassLabel
    (
        IntPtr lpClass,                /* i/o Class to set label  */
        int lIntfId,                /* i  Interface id    */
        string pszValue                /* i  Value of property    */
    );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectAllEnvs(bool showSystem);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectEnvByProjectId(int lProjectId);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectEnvByTableId(int iTableId);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectApplication(int lApplId);


    [DllImport("dmscli.dll", EntryPoint = "aaApi_GetApplicationStringProperty", CharSet = CharSet.Unicode)]
    private static extern IntPtr unsafe_aaApi_GetApplicationStringProperty(ApplicationProperty lPropertyId, int lIndex);


    public static string aaApi_GetApplicationStringProperty(ApplicationProperty lPropertyId, int lIndex)
    {
        return Marshal.PtrToStringUni(unsafe_aaApi_GetApplicationStringProperty(lPropertyId, lIndex));
    }


    public enum StorageProperty : int
    {
        ID = 1,
        Name = 2,
        Desc = 3,
        Node = 4,
        Path = 5,
        Protocol = 6,
        DisplayName = 7
    }


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectAllStorages();


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GetStorageId(int lIndex);


    [DllImport("dmscli.dll", EntryPoint = "aaApi_GetStorageStringProperty", CharSet = CharSet.Unicode)]
    private static extern IntPtr unsafe_aaApi_GetStorageStringProperty(StorageProperty PropertyId, int lIndex);


    public static string aaApi_GetStorageStringProperty(StorageProperty PropertyId, int lIndex)
    {
        return Marshal.PtrToStringUni(unsafe_aaApi_GetStorageStringProperty(PropertyId, lIndex));
    }


    public static int BuildPWPath(string pwPath, int targetVaultID, string userName, int environmentID)
    {
        bool bDoCaselessCompare = true;

        // oracle
        if (PWWrapper.aaApi_GetActiveDatasourceNativeType() == (int)PWWrapper.DataSourceType.Oracle ||
            PWWrapper.aaApi_GetActiveDatasourceNativeType() == (int)PWWrapper.DataSourceType.ODBC_Oracle ||
            PWWrapper.aaApi_GetActiveDatasourceNativeType() == 1)
        {
            bDoCaselessCompare = false;
        }

        // get the userID (should just be one user with that user name)
        int numUsers = aaApi_SelectUsersByProp(userName, null, null);
        int userID = aaApi_GetUserId(0);
        int storageID = 1;
        bool bUseEnvId = false;
        if (environmentID > 0)
            bUseEnvId = true;

        int iEnvironmentId = 0;

        char[] delimiters = { '\\', '/' };
        string[] pathSteps = pwPath.Split(delimiters, StringSplitOptions.RemoveEmptyEntries);

        int parentVaultID = targetVaultID;
        int childVaultID = 0;  // 0 is an invalid vault ID

        for (int i = 0; i < pathSteps.Length; i++)
        {
            string vaultName = pathSteps[i].Trim(" ".ToCharArray());

            if (string.IsNullOrEmpty(vaultName))
                continue;

            if (vaultName.Length > 63)
                vaultName = vaultName.Substring(0, 63);

            // search for the vault to see if it already exists
            int numChildren = -1;
            if (-1 == parentVaultID)
            {
                numChildren = aaApi_SelectTopLevelProjects();
            }
            else
            {
                numChildren = aaApi_SelectChildProjects(parentVaultID);
            }
            if (numChildren == -1)
            {
                // string message = "Error selecting child folders";
                return 0;
                // throw new ApplicationException(message);
            }

            bool childFound = false;
            for (int j = 0; j < numChildren; j++)
            {
                string childVaultName = aaApi_GetProjectStringProperty(ProjectProperty.Name, j);

                if ((childVaultName == vaultName) ||
                    (bDoCaselessCompare && childVaultName.ToLower() == vaultName.ToLower()))
                {
                    childFound = true;
                    childVaultID = aaApi_GetProjectNumericProperty(ProjectProperty.ID, j);
                    storageID = aaApi_GetProjectNumericProperty(ProjectProperty.StorageID, j);
                    iEnvironmentId = aaApi_GetProjectNumericProperty(ProjectProperty.EnvironmentID, j);
                    break;
                }
            }

            // if the child vault was not found, create it
            if (childFound == false)
            {
                bool success = aaApi_CreateProject(ref childVaultID,
                    parentVaultID, storageID, userID, VaultType.Normal,
                    0, 0, 0, vaultName, "");

                if (success == true)
                {
                    VaultDescriptor vaultDescriptor = new VaultDescriptor();
                    vaultDescriptor.Flags = (uint)(
                        VaultDescriptorFlags.EnvironmentID |
                        VaultDescriptorFlags.VaultID);

                    if (bUseEnvId)
                    {
                        vaultDescriptor.EnvironmentID = environmentID;
                    }
                    else
                    {
                        vaultDescriptor.EnvironmentID = iEnvironmentId;
                    }

                    vaultDescriptor.VaultID = childVaultID;

                    success = aaApi_ModifyProject2(ref vaultDescriptor);
                }

                if (success == false)
                {
                    // string message = String.Format(pwResources.GetString("ErrorCreatingVault"),
                    // vaultName, GetLastPWError());
                    // throw new ApplicationException(message);
                    return 0;
                }
            }

            parentVaultID = childVaultID;
        }

        return childVaultID;
    }


    public static int GetEnvironmentId(string sEnvironmentName)
    {
        if (!string.IsNullOrEmpty(sEnvironmentName))
        {
            for (int i = 0; i < PWWrapper.aaApi_SelectAllEnvs(false); i++)
            {
                string sEnvNameTest = PWWrapper.aaApi_GetEnvStringProperty(
                    (int)PWWrapper.EnvironmentProperty.Name, i);

                if (sEnvNameTest.ToLower() == sEnvironmentName.ToLower())
                {
                    return PWWrapper.aaApi_GetEnvId(i);
                }
            }
        }

        return 0;
    }

    public static int GetDepartmentId(string sDepartmentName)
    {
        if (!string.IsNullOrEmpty(sDepartmentName))
        {
            for (int i = 0; i < PWWrapper.aaApi_SelectAllDepartments(); i++)
            {
                string sDeptNameTest = PWWrapper.aaApi_GetDepartmentStringProperty(DepartmentProperty.Name, i);

                if (sDeptNameTest.ToLower() == sDepartmentName.ToLower())
                {
                    return PWWrapper.aaApi_GetDepartmentId(i);
                }
            }
        }

        return 0;
    }


    public static int GetStorageAreaId(string sStorageAreaName)
    {
        int iNumStorages = PWWrapper.aaApi_SelectAllStorages();

        for (int i = 0; i < iNumStorages; i++)
        {
            string sName =
                PWWrapper.aaApi_GetStorageStringProperty(PWWrapper.StorageProperty.Name, i);

            if (sStorageAreaName.ToLower() == sName.ToLower())
                return PWWrapper.aaApi_GetStorageId(i);
        }

        return 0;
    }


    public static int BuildPath(string pwPath, int targetVaultID,
        int environmentID, int storageID)
    {
        bool bDoCaselessCompare = true;

        // oracle
        if (PWWrapper.aaApi_GetActiveDatasourceNativeType() == (int)PWWrapper.DataSourceType.Oracle ||
            PWWrapper.aaApi_GetActiveDatasourceNativeType() == (int)PWWrapper.DataSourceType.ODBC_Oracle ||
            PWWrapper.aaApi_GetActiveDatasourceNativeType() == 1)
        {
            bDoCaselessCompare = false;
        }

        char[] delimiters = { '\\', '/' };
        string[] pathSteps = pwPath.Split(delimiters, StringSplitOptions.RemoveEmptyEntries);

        int parentVaultID = targetVaultID;
        int childVaultID = 0;  // 0 is an invalid vault ID

        if (pathSteps.Length == 0 && targetVaultID > 0)
            return targetVaultID;

        for (int i = 0; i < pathSteps.Length; i++)
        {
            string vaultName = pathSteps[i].Trim(" ".ToCharArray());

            if (string.IsNullOrEmpty(vaultName))
                continue;

            if (vaultName.Length > 63)
                vaultName = vaultName.Substring(0, 63);

            // search for the vault to see if it already exists
            int numChildren = -1;
            if (-1 == parentVaultID)
            {
                numChildren = PWWrapper.aaApi_SelectTopLevelProjects();
            }
            else
            {
                numChildren = PWWrapper.aaApi_SelectChildProjects(parentVaultID);
            }

            if (numChildren == -1)
            {
                return 0;
            }

            bool childFound = false;
            for (int j = 0; j < numChildren; j++)
            {
                string childVaultName =
                    PWWrapper.aaApi_GetProjectStringProperty(PWWrapper.ProjectProperty.Name, j);

                if ((childVaultName == vaultName) ||
                    (bDoCaselessCompare && (childVaultName.ToLower() == vaultName.ToLower())))
                {
                    childFound = true;
                    childVaultID =
                        PWWrapper.aaApi_GetProjectNumericProperty(PWWrapper.ProjectProperty.ID, j);
                    break;
                }
            }

            // if the child vault was not found, create it
            if (childFound == false)
            {
                if (-1 != parentVaultID && storageID <= 0)
                {
                    if (1 == PWWrapper.aaApi_SelectProject(parentVaultID))
                    {
                        storageID =
                            PWWrapper.aaApi_GetProjectNumericProperty(
                                PWWrapper.ProjectProperty.StorageID, 0);
                    }
                }

                if (storageID <= 0)
                {
                    int iNumStores = PWWrapper.aaApi_SelectAllStorages();
                    if (iNumStores > 0)
                        storageID = PWWrapper.aaApi_GetStorageId(0);
                }

                bool success = PWWrapper.aaApi_CreateProject(ref childVaultID,
                    parentVaultID, storageID, PWWrapper.aaApi_GetCurrentUserId(),
                    PWWrapper.VaultType.Normal,
                    0, 0, 0, vaultName, "");

                if (success == true)
                {
                    PWWrapper.VaultDescriptor vaultDescriptor = new PWWrapper.VaultDescriptor();
                    vaultDescriptor.Flags = (uint)(
                        PWWrapper.VaultDescriptorFlags.EnvironmentID |
                        PWWrapper.VaultDescriptorFlags.VaultID);

                    vaultDescriptor.EnvironmentID = environmentID;
                    vaultDescriptor.VaultID = childVaultID;

                    success = PWWrapper.aaApi_ModifyProject2(ref vaultDescriptor);
                }
                else
                {
                    return 0;
                }
            }

            parentVaultID = childVaultID;
        }

        return childVaultID;
    }

    public static int BuildPathWithBackslashesOnly(string pwPath, int targetVaultID,
        int environmentID, int storageID, int iWorkflowId)
    {
        bool bDoCaselessCompare = true;

        // oracle
        if (PWWrapper.aaApi_GetActiveDatasourceNativeType() == (int)PWWrapper.DataSourceType.Oracle ||
            PWWrapper.aaApi_GetActiveDatasourceNativeType() == (int)PWWrapper.DataSourceType.ODBC_Oracle ||
            PWWrapper.aaApi_GetActiveDatasourceNativeType() == 1)
        {
            bDoCaselessCompare = false;
        }

        char[] delimiters = { '\\' };
        string[] pathSteps = pwPath.Split(delimiters, StringSplitOptions.RemoveEmptyEntries);

        int parentVaultID = targetVaultID;
        int childVaultID = 0;  // 0 is an invalid vault ID

        if (pathSteps.Length == 0 && targetVaultID > 0)
            return targetVaultID;

        for (int i = 0; i < pathSteps.Length; i++)
        {
            string vaultName = pathSteps[i].Trim(" ".ToCharArray());

            if (string.IsNullOrEmpty(vaultName))
                continue;

            if (vaultName.Length > 63)
                vaultName = vaultName.Substring(0, 63);

            // search for the vault to see if it already exists
            int numChildren = -1;
            if (-1 == parentVaultID)
            {
                numChildren = PWWrapper.aaApi_SelectTopLevelProjects();
            }
            else
            {
                numChildren = PWWrapper.aaApi_SelectChildProjects(parentVaultID);
            }

            if (numChildren == -1)
            {
                return 0;
            }

            bool childFound = false;
            for (int j = 0; j < numChildren; j++)
            {
                string childVaultName =
                    PWWrapper.aaApi_GetProjectStringProperty(PWWrapper.ProjectProperty.Name, j);

                if ((childVaultName == vaultName) ||
                    (bDoCaselessCompare && (childVaultName.ToLower() == vaultName.ToLower())))
                {
                    childFound = true;
                    childVaultID =
                        PWWrapper.aaApi_GetProjectNumericProperty(PWWrapper.ProjectProperty.ID, j);
                    break;
                }
            }

            // if the child vault was not found, create it
            if (childFound == false)
            {
                if (-1 != parentVaultID && storageID <= 0)
                {
                    if (1 == PWWrapper.aaApi_SelectProject(parentVaultID))
                    {
                        storageID =
                            PWWrapper.aaApi_GetProjectNumericProperty(
                                PWWrapper.ProjectProperty.StorageID, 0);
                    }
                }

                if (storageID <= 0)
                {
                    int iNumStores = PWWrapper.aaApi_SelectAllStorages();
                    if (iNumStores > 0)
                        storageID = PWWrapper.aaApi_GetStorageId(0);
                }

                bool success = PWWrapper.aaApi_CreateProject(ref childVaultID,
                    parentVaultID, storageID, PWWrapper.aaApi_GetCurrentUserId(),
                    PWWrapper.VaultType.Normal,
                    0, 0, 0, vaultName, "");

                if (success == true)
                {
                    PWWrapper.VaultDescriptor vaultDescriptor = new PWWrapper.VaultDescriptor();
                    vaultDescriptor.Flags = (uint)(
                        PWWrapper.VaultDescriptorFlags.EnvironmentID |
                        PWWrapper.VaultDescriptorFlags.VaultID |
                        PWWrapper.VaultDescriptorFlags.Workflow);

                    vaultDescriptor.EnvironmentID = environmentID;
                    vaultDescriptor.VaultID = childVaultID;
                    vaultDescriptor.WorkflowID = iWorkflowId;

                    // doesn't work to set workflow
                    success = PWWrapper.aaApi_ModifyProject2(ref vaultDescriptor);

                    if (!PWWrapper.aaApi_SetProjectWorkflow(childVaultID, iWorkflowId))
                    {
                        System.Diagnostics.Debug.WriteLine(PWWrapper.aaApi_GetLastErrorMessage());
                    }
                }
                else
                {
                    return 0;
                }
            }

            parentVaultID = childVaultID;
        }

        return childVaultID;
    }


    public static int BuildPathWithBackslashesOnly(string pwPath, int targetVaultID,
        int environmentID, int storageID)
    {
        bool bDoCaselessCompare = true;

        // oracle
        if (PWWrapper.aaApi_GetActiveDatasourceNativeType() == (int)PWWrapper.DataSourceType.Oracle ||
            PWWrapper.aaApi_GetActiveDatasourceNativeType() == (int)PWWrapper.DataSourceType.ODBC_Oracle ||
            PWWrapper.aaApi_GetActiveDatasourceNativeType() == 1)
        {
            bDoCaselessCompare = false;
        }

        char[] delimiters = { '\\' };
        string[] pathSteps = pwPath.Split(delimiters, StringSplitOptions.RemoveEmptyEntries);

        int parentVaultID = targetVaultID;
        int childVaultID = 0;  // 0 is an invalid vault ID

        if (pathSteps.Length == 0 && targetVaultID > 0)
            return targetVaultID;

        for (int i = 0; i < pathSteps.Length; i++)
        {
            string vaultName = pathSteps[i].Trim(" ".ToCharArray());

            if (string.IsNullOrEmpty(vaultName))
                continue;

            if (vaultName.Length > 63)
                vaultName = vaultName.Substring(0, 63);

            // search for the vault to see if it already exists
            int numChildren = -1;
            if (-1 == parentVaultID)
            {
                numChildren = PWWrapper.aaApi_SelectTopLevelProjects();
            }
            else
            {
                numChildren = PWWrapper.aaApi_SelectChildProjects(parentVaultID);
            }

            if (numChildren == -1)
            {
                return 0;
            }

            bool childFound = false;
            for (int j = 0; j < numChildren; j++)
            {
                string childVaultName =
                    PWWrapper.aaApi_GetProjectStringProperty(PWWrapper.ProjectProperty.Name, j);

                if ((childVaultName == vaultName) ||
                    (bDoCaselessCompare && (childVaultName.ToLower() == vaultName.ToLower())))
                {
                    childFound = true;
                    childVaultID =
                        PWWrapper.aaApi_GetProjectNumericProperty(PWWrapper.ProjectProperty.ID, j);
                    break;
                }
            }

            // if the child vault was not found, create it
            if (childFound == false)
            {
                if (-1 != parentVaultID && storageID <= 0)
                {
                    if (1 == PWWrapper.aaApi_SelectProject(parentVaultID))
                    {
                        storageID =
                            PWWrapper.aaApi_GetProjectNumericProperty(
                                PWWrapper.ProjectProperty.StorageID, 0);
                    }
                }

                if (storageID <= 0)
                {
                    int iNumStores = PWWrapper.aaApi_SelectAllStorages();
                    if (iNumStores > 0)
                        storageID = PWWrapper.aaApi_GetStorageId(0);
                }

                bool success = PWWrapper.aaApi_CreateProject(ref childVaultID,
                    parentVaultID, storageID, PWWrapper.aaApi_GetCurrentUserId(),
                    PWWrapper.VaultType.Normal,
                    0, 0, 0, vaultName, "");

                if (success == true)
                {
                    PWWrapper.VaultDescriptor vaultDescriptor = new PWWrapper.VaultDescriptor();
                    vaultDescriptor.Flags = (uint)(
                        PWWrapper.VaultDescriptorFlags.EnvironmentID |
                        PWWrapper.VaultDescriptorFlags.VaultID);

                    vaultDescriptor.EnvironmentID = environmentID;
                    vaultDescriptor.VaultID = childVaultID;

                    success = PWWrapper.aaApi_ModifyProject2(ref vaultDescriptor);
                }
                else
                {
                    return 0;
                }
            }

            parentVaultID = childVaultID;
        }

        return childVaultID;
    }

    public static int ProjectNoFromPath(string pwPath)
    {
        bool bDoCaselessCompare = true;

        // oracle
        if (PWWrapper.aaApi_GetActiveDatasourceNativeType() == (int)PWWrapper.DataSourceType.Oracle ||
            PWWrapper.aaApi_GetActiveDatasourceNativeType() == (int)PWWrapper.DataSourceType.ODBC_Oracle ||
            PWWrapper.aaApi_GetActiveDatasourceNativeType() == 1)
        {
            bDoCaselessCompare = false;
        }

        return ProjectNoFromPath(pwPath, bDoCaselessCompare);
    }

    private static int ProjectNoFromPath(string pwPath, bool bDoCaselessCompare)
    {
        // char[] delimiters = { '\\', '/' };
        char[] delimiters = { '\\' }; // changed DAB 2012-09-07
        string[] pathSteps = pwPath.Split(delimiters, StringSplitOptions.RemoveEmptyEntries);

        int parentVaultID = -1;
        int childVaultID = 0;  // 0 is an invalid vault ID

        if (pathSteps.Length == 0)
            return 0;

        for (int i = 0; i < pathSteps.Length; i++)
        {
            if (string.IsNullOrEmpty(pathSteps[i]))
                continue;

            string vaultName = pathSteps[i].Trim(" ".ToCharArray());

            if (string.IsNullOrEmpty(vaultName))
                continue;

            if (vaultName.Length > 63)
                vaultName = vaultName.Substring(0, 63);

            // search for the vault to see if it already exists
            int numChildren = -1;
            if (-1 == parentVaultID)
            {
                numChildren = PWWrapper.aaApi_SelectTopLevelProjects();
            }
            else
            {
                numChildren = PWWrapper.aaApi_SelectChildProjects(parentVaultID);
            }

            if (numChildren == -1)
            {
                return 0;
            }

            bool bFoundChildVault = false;

            for (int j = 0; j < numChildren; j++)
            {
                string childVaultName =
                    PWWrapper.aaApi_GetProjectStringProperty(PWWrapper.ProjectProperty.Name, j);

                if ((bDoCaselessCompare && childVaultName.ToLower() == vaultName.ToLower()) ||
                    (childVaultName == vaultName))
                {
                    childVaultID =
                        PWWrapper.aaApi_GetProjectNumericProperty(PWWrapper.ProjectProperty.ID, j);
                    bFoundChildVault = true;
                    break;
                }
            }

            if (!bFoundChildVault)
                return 0;

            parentVaultID = childVaultID;
        }

        return childVaultID;
    }

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectProjectsByProp
(
string lpctstrCode,    /* i  Project code to search for        */
string lpctstrName,    /* i  Project name to search for        */
string lpctstrDesc,    /* i  Project description to search for */
string lpctstrVersion  /* i  Project version to search for     */
);


    [DllImport("dmsgen.dll", EntryPoint = "aaApi_GetLastErrorMessage", CharSet = CharSet.Unicode)]
    public static extern IntPtr unsafe_aaApi_GetLastErrorMessage();


    public static string aaApi_GetLastErrorMessage()
    {
        return Marshal.PtrToStringUni(unsafe_aaApi_GetLastErrorMessage());
    }

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern System.IntPtr aaApi_SelectProjectDataBufferChilds(
        int lProjectId             /* i  Number of parent project  */
        );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern System.IntPtr aaApi_SelectProjectDataBufferChilds2(
        int lProjectId,             /* i  Number of parent project  */
        bool bWithRichProjectsOnly
        );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern System.IntPtr aaApi_SelectProjectDataBuffer(
        int lProjectId             /* i  Number of project  */
        );


    public enum DmsDataBufferStringPropertyEnum : int
    {
        //project buffer properties
        PROJ_PROP_NAME = 12,
        PROJ_PROP_DESC = 13,
        PROJ_PROP_CODE = 14,
        PROJ_PROP_VERSION = 15,
        PROJ_PROP_CREATE_TIME = 16,
        PROJ_PROP_UPDATE_TIME = 17,

        //document buffer properties
        DOC_PROP_NAME = 20,
        DOC_PROP_FILENAME = 21
    }

    public enum ProjectNumericPropertyEnum : int
    {
        PROJ_PROP_ID = 1,
        PROJ_PROP_VERSIONNO = 2,
        PROJ_PROP_MANAGERID = 3,
        PROJ_PROP_STORAGEID = 4,
        PROJ_PROP_CREATORID = 5,
        PROJ_PROP_UPDATERID = 6,
        PROJ_PROP_WORKFLOWID = 7,
        PROJ_PROP_STATEID = 8,
        PROJ_PROP_TYPE = 9,
        PROJ_PROP_ARCHIVEID = 10,
        PROJ_PROP_ISPARENT = 11,

        PROJ_PROP_ENVIRONMENTID = 21,
        PROJ_PROP_PARENTID = 22,
        PROJ_PROP_MGRTYPE = 23
    }

    public enum DataSourceType
    {
        Unknown = 0,
        RIS = 1,
        ODBC = 2,
        Informix = 3,
        Ingres = 4,
        Oracle = 5,
        SqlAnywhere = 6,
        SqlServer = 7,
        Sybase = 8,
        DB2 = 9,
        Optinet = 10,
        Solid = 11,

        ODBC_Informix = 12,
        ODBC_Ingres = 16,
        ODBC_Oracle = 20,
        ODBC_SqlAnywhere = 24,
        ODBC_SqlServer = 28,
        ODBC_Sybase = 32,
        ODBC_DB2 = 36,
        ODBC_Solid = 44
    }


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_AdminLogin(DataSourceType lDSType, string sDatasourceName,
        string lpctstrUsername, string lpctstrPassword);

    [DllImport("dmscli.dll", EntryPoint = "aaApi_GetDocumentStringProperty", CharSet = CharSet.Unicode)]
    private static extern IntPtr unsafe_aaApi_GetDocumentStringProperty(DocumentProperty PropertyId, int Index);


    public enum GroupProperty : int
    {
        ID = 1,
        Name = 2,
        Desc = 3,
        Type = 4,
        SecProvider = 5
    }


    public enum ODSNativeIdType : int
    {
        Undefined = 0,
        None = 1,
        DgnElementId = 2,
        DgnModelId = 3,
        DgnLevelId = 4,
        XGLPath = 5,
        XPath = 6,
        JSpaceId = 7,
        SheetName = 8,
        DgnCustomLStyleId = 9,
        SheetView = 10,
        DgnCellId = 11,
        DgnSavedViewId = 12,
        JSpaceIdLink = 13
    }


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_IsCurrentUserAdmin();

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_IsUserRestrictedAdmin(int iUserId);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectAllGroups();


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_SetCurrentSession(IntPtr handle);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_GetCurrentSession(ref IntPtr handle);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectGroupsByUser(int lUserId);


    [DllImport("dmscli.dll", EntryPoint = "aaApi_GetGroupStringProperty", CharSet = CharSet.Unicode)]
    private static extern IntPtr unsafe_aaApi_GetGroupStringProperty(GroupProperty PropertyId, int lIndex);


    public static string aaApi_GetGroupStringProperty(GroupProperty PropertyId, int lIndex)
    {
        return Marshal.PtrToStringUni(unsafe_aaApi_GetGroupStringProperty(PropertyId, lIndex));
    }

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GetGroupNumericProperty(GroupProperty PropertyId, int lIndex);

    [Flags]
    public enum AccessControlSelectionFlags : uint
    {
        IgnoreParents = 0x00000001,
        IgnoreEnvironment = 0x00000002,
        IgnoreDefault = 0x00000004,
        ExactMatch = 0x00000008,
        IgnoreUserSettings = 0x00000010,
        AllWorkflowStates = 0x00000020,
        IgnoreObjAcce = 0x00010000
    }



    public enum AccessObjectType : int
    {
        UserIgnoresAccessCtrl = 0, /* only for return values */
        EnvironmentProject = 1,
        Project = 2,
        EnvironmentDocument = 3,
        Document = 4,
        ODSComponent = 5,
        AdminControl = 6,
        Group = 7,
        UserList = 8,
        Transmittal = 9
    }


    public enum AccessObjectProperty : int
    {
        ObjType = 1,
        ObjID1 = 2,
        ObjID2 = 3,
        Workflow = 4,
        State = 5,
        MemberType = 6,
        MemberId = 7,
        AccessMask = 8,
        ObjGUID = 9
    }




    public enum UserListProperty : int
    {
        ID = 1,
        Name = 2,
        Description = 3,
        Type = 4,
        Owner = 5
    }


    public enum UserListMemberProperty : int
    {
        ListID = 1,
        MemberType = 2,
        MemberID = 3
    }


    public enum ManagerTypes : int
    {
        User = 1,
        Group = 2,
        UserList = 3,
        AllUsers = 4
    }

    public enum MemberTypes : int
    {
        User = 1,
        Group = 2,
        UserList = 3,
        AllUsers = 4
    }

    public enum UserListTypes : int
    {
        UserList = 1,
        AddressBook = 2
    }

    public enum ProjectResourceTypes : int
    {
        Application = 1,
        Department = 2,
        Environment = 3,
        StorageArea = 4,
        View = 5,
        Workflow = 6,
        WorkspaceProfile = 7
    }


    [Flags]
    public enum AccessMaskFlags : uint
    {
        None = 0x00000000,
        Control = 0x00000001,
        Write = 0x00000002,
        Read = 0x00000004,
        FileWrite = 0x00000008,
        FileRead = 0x00000010,
        Create = 0x00000020,
        Delete = 0x00000040,
        Full = 0x0000FFFF
    }

    [Flags]
    public enum ReferenceListFlags : uint
    {
        FromReferenceInfo = 0x00000001,
        FromSetInfo = 0x00000002,
        AllowSelfReferences = 0x00000004
    }

    [Flags]
    public enum ProjectCopyDeleteAndExportFlags : uint
    {
        ExcludeParent = 0x00000001,                     /* copy and delete */
        NoDocuments = 0x00000002,                       /* copy and delete */
        NoSets = 0x00000004,                            /* copy and delete */
        NoRecursion = 0x00000008,                       /* copy, delete and export */
        Attributes = 0x00000010,                        /* copy and delete */
        NoProjects = 0x00000020,                        /* copy and delete */
        TakeOwnership = 0x00000040,                     /* copy only */
        AllowCopyAll = 0x00000080,                      /* copy only */
        NoCheckedOut = 0x00000100,                      /* copy and delete */
        SetReferences = 0x00000200,                     /* delete */
        OwnCheckOuts = 0x00000400,                      /* copy and delete */
        NoActiveVersion = 0x00000800,                   /* delete */
        DelManagedWorkspaceVars = 0x00001000,           /* delete */
        CopyWorkflow = 0x00001000,                      /* copy only   */
        CopyAccess = 0x00002000,                        /* copy only   */
        CopyManager = 0x00004000,                       /* copy only   */
        CopyStorage = 0x00008000,                       /* copy only   */
        CopyEnvironment = 0x00010000,                   /* copy only   */
        CopyVersions = 0x00020000,                      /* copy only   */
        Components = 0x00040000,                        /* copy only   */
        CopySavedSearch = 0x00080000,                   /* copy only   */
        CopyResources = 0x00100000,                     /* copy only   */
        CopyConfigurationBlocks = 0x00200000,           /* copy only   */
        CopyContents = 0x00400000,                      /* copy contents of project to existing project */
        CopyWorkspaceProfile = 0x00800000,              /* copy only   */
        ExportEmptyProjects = 0x00100000,               /* export only (export subprojects even if they are empty)  */
        ExportRootProject = 0x00200000,                 /* export only (create folder for exported root project)    */
        ExportUsingProjectDescriptions = 0x00400000,    /* export only (use project descriptions as folder names)   */
        ExportRefsToMaster = 0x00800000,                /* export only (export references to master document folder)*/
        ExportGiveOut = 0x01000000,                     /* export only (perform giveout for all documents) */
        ExportOuterRefs = 0x02000000,                   /* export only (references which are not in the export hierarchy will be exported to the specified folder) */
        ExportRewriteRefs = 0x04000000,                 /* export only (rewrite the reference attachments to reflect new hierarchy) */
        ExportShared = 0x08000000,                      /* export only (export shareable documents as shared) */
        ForCopy = 0x10000000,                           /* Fills attr for copy only */
        NonDMSExport = 0x20000000                       /* Handled by callback      */
    };

    /* Flags for project copy, delete and export */
    //#define AAPRO_ARRAY_EXCLUDE_PARENT  0x00000001L /* copy and delete */
    //#define AAPRO_ARRAY_NO_DOCUMENTS    0x00000002L /* copy and delete */
    //#define AAPRO_ARRAY_NO_SETS         0x00000004L /* copy and delete */
    //#define AAPRO_ARRAY_NO_RECURSIO     0x00000008L /* copy, delete and export */
    //#define AAPRO_ARRAY_ATTRIBUTES      0x00000010L /* copy and delete */
    //#define AAPRO_ARRAY_NO_PROJECTS     0x00000020L /* copy and delete */
    //#define AAPRO_ARRAY_TAKE_OWNERSHIP  0x00000040L /* copy only */
    //#define AAPRO_ARRAY_ALLOW_COPY_ALL  0x00000080L /* copy only */
    //#define AAPRO_ARRAY_NO_CHECKED_OUT  0x00000100L /* copy and delete */
    //#define AAPRO_ARRAY_SET_REFERENCES  0x00000200L /* delete only */
    //#define AAPRO_ARRAY_OWN_CHECK_OUTS  0x00000400L /* copy and delete */
    //#define AAPRO_ARRAY_NO_ACTIVE_VER   0x00000800L /* delete only */
    //#define AAPRO_ARRAY_DEL_MWP_VARS    0x00001000L /* delete only */
    //#define AAPRO_ARRAY_COPY_WORKFLOW   0x00001000L /* copy only   */
    //#define AAPRO_ARRAY_COPY_ACCESS     0x00002000L /* copy only   */
    //#define AAPRO_ARRAY_COPY_MANAGER    0x00004000L /* copy only   */
    //#define AAPRO_ARRAY_COPY_STORAGE    0x00008000L /* copy only   */
    //#define AAPRO_ARRAY_COPY_ENV        0x00010000L /* copy only   */
    //#define AAPRO_ARRAY_COPY_VERSIONS   0x00020000L /* copy only   */
    //#define AAPRO_ARRAY_COMPONENTS      0x00040000L /* copy only   */
    //#define AAPRO_ARRAY_COPY_SAVED_SRC  0x00080000L /* copy only   */
    //#define AAPRO_ARRAY_COPY_RESOURCES  0x00100000L /* copy only   */
    //#define AAPRO_ARRAY_COPY_CONFBLOCKS 0x00200000L /* copy only   */
    //#define AAPRO_ARRAY_COPY_CONTENTS   0x00400000L /* copy contents of project to existing project */
    //#define AAPRO_ARRAY_COPY_WS_PROFL   0x00800000L /* copy only   */

    //#define AAPRO_ARRAY_EXP_EMPTY_PRO   0x00100000L /* export only (export subprojects even if they are empty)  */
    //#define AAPRO_ARRAY_EXP_SUBF_ROOT   0x00200000L /* export only (create folder for exported root project)    */
    //#define AAPRO_ARRAY_EXP_PRJ_DESCR   0x00400000L /* export only (use project descriptions as folder names)   */
    //#define AAPRO_ARRAY_EXP_REF_2_MST   0x00800000L /* export only (export references to master document folder)*/
    //#define AAPRO_ARRAY_EXP_GIVE_OUT    0x01000000L /* export only (perform giveout for all documents) */
    //#define AAPRO_ARRAY_EXP_OUTER_REFS  0x02000000L /* export only (references which are not in the export hierarchy will be exported to the specified folder) */
    //#define AAPRO_ARRAY_EXP_REWRITE_REF 0x04000000L /* export only (rewrite the reference attachments to reflect new hierarchy) */
    //#define AAPRO_ARRAY_EXP_SHARED      0x08000000L /* export only (export shareable documents as shared) */

    //#define AAPRO_ARRAY_FOR_COPY        0x10000000L /* Fills attr for copy only */
    //#define AAPRO_ARRAY_NONDMS_EXPORT   0x20000000L /* Handled by callback      */



    public enum UserProperty : int
    {
        ID = 1,
        Name = 2,
        Desc = 3,
        Password = 4,
        Email = 5,
        Type = 6,
        SecProvider = 7,
        Flags = 8,
        CreateDate = 9
    }

    public enum ODSAttributeProperty : int
    {
        ID = 1,
        Name = 2,
        Desc = 3,
        Visibility = 4,
        DataType = 5,
        DataLength = 6,
        Control = 7,
        InstanceColumn = 8,
        Label = 9,
        FunctionId = 12,
        MirrorId = 13,
        LinkId = 14,
        LinkAttrId = 15,
        Direction = 16,
        UIType = 32
    }

    public enum ReferenceInformationProperty : int
    {
        ElementIDUint64 = 1, // uint64
        MasterGUID = 2,
        MasterModelID = 3,
        ReferenceGUID = 4,
        ReferenceModelID = 5,
        NestDepth = 6,
        ReferenceType = 7,
        Flags = 8
    }

    public enum ODSClassProperty : int
    {
        ID = 1,
        Name = 2,
        Desc = 3,
        SystemClass = 4,
        KeyId = 5,
        ClassIdAttr = 6,
        TableName = 7,
        SequenceName = 8,
        CatalogName = 9,
        CatalogKeyId = 10,
        ModTime = 11,
        IsVersion = 12,
        CurrentId = 13,
        FutureId = 14,
        HistoryId = 15,
        ClassType = 20,
        Label = 21
    }

    public enum ODSAttributeTypes : int
    {
        AAODS_ATTRTYPE_DATABASE = 1,
        AAODS_ATTRTYPE_USERDATA = 2,
        AAODS_ATTRTYPE_CONSTANT = 3,
        AAODS_ATTRTYPE_LINKAGE = 4,
        AAODS_ATTRTYPE_LAST = 4
    }

    public enum ODSClassTypes : uint
    {
        AAODS_CLASS_NORMAL = 0x00000001,
        AAODS_CLASS_LINK = 0x00000002,
        AAODS_CLASS_SYSTEM = 0x00000004,
        AAODS_CLASS_FUTURE = 0x00000010,
        AAODS_CLASS_HISTORY = 0x00000020,
    }

    public enum ODSAttributeDataType : int
    {
        Int16 = 1,
        Long32 = 2,
        Float32 = 3,
        Double64 = 4,
        String = 5,
        Timestamp = 6,
        Raw = 7,
        LongRaw = 8,
        DateTime = 9
    }


    public enum AccessUserProperty : int
    {
        UserID = 1,
        AccessMask = 2
    }


    [StructLayout(LayoutKind.Sequential)]
    public struct AaDocItem
    {
        public int lProjectId;
        public int lDocumentId;
    };

    //[StructLayout(LayoutKind.Explicit)]
    //public struct Rect
    //{
    //    [FieldOffset(0)]
    //    public int left;
    //    [FieldOffset(4)]
    //    public int top;
    //    [FieldOffset(8)]
    //    public int right;
    //    [FieldOffset(12)]
    //    public int bottom;
    //}


    [StructLayout(LayoutKind.Sequential)]
    public struct AaDocumentsParam
    {
        //  Specifies valid fields mask, specifying whether fields ulFlags and lParam2 are valid or not. 
        // [FieldOffset(0)]
        public uint uiMask;
        //  Specifies count of elements in the array specified by lpDocuments. 
        //[FieldOffset(4)]
        public int iCount;
        //  A Pointer to the array of structures specifying the documents to be processed. 
        //[FieldOffset(8)]
        public AaDocItem lpDocuments;
        //[MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)]
        //public int[] lpDocuments;
        //  Specifies first operation specific parameter. 
        // [FieldOffset(12)]
        public int iParam1;
        //  Specifies operation specific mask. 
        //[FieldOffset(16)]
        public uint uiFlags;
        //  Specifies second operation specific parameter. 
        //[FieldOffset(20)]
        public int iParam2;
        //  Operation comment for audit trail. 
        public string sComment;
        //  Buffer of processed documents (type AADMSBUFFER_DOCUMENT). 
        IntPtr hProcessedDocuments;
    };



    [Flags]
    public enum AaProjItemFlags : uint
    {
        AADMSPROJF_PROJECTID = 0x00000001,
        AADMSPROJF_ENVID = 0x00000002,
        AADMSPROJF_PARENTID = 0x00000004,
        AADMSPROJF_STORAGEID = 0x00000008,
        AADMSPROJF_MANAGERID = 0x00000010,
        AADMSPROJF_TYPEID = 0x00000020,
        AADMSPROJF_WORKFLOW = 0x00000040,
        AADMSPROJF_NAME = 0x00000080,
        AADMSPROJF_DESC = 0x00000100,
        AADMSPROJF_MGRTYPE = 0x00000400,
        AADMSPROJF_WSPACEPROFID = 0x00000800,
        AADMSPROJF_GUID = 0x00001000,
        AADMSPROJF_COMPONENT_CLASSID = 0x00002000,
        AADMSPROJF_PROJFLAGS = 0x00004000,
        AADMSPROJF_COMPONENT_INSTANCEID = 0x00008000,
        AADMSPROJF_REQUIREDONCREATE = (AADMSPROJF_STORAGEID | AADMSPROJF_NAME),
        AADMSPROJF_ALL = 0x0000FFFF
    }
    public enum AAODSHIERARCHY : int
    {
        AAODS_HIERARCHY_ID = 1
    }
    //public enum   AAODSHIERARCHY_: string
    //{
    //    AAODS_HIERARCHY_NAME             = "ActiveAsset",
    //    AAODS_HIERARCHY_DESC             = "ActiveAsset Hierarchy"
    //}

    [StructLayout(LayoutKind.Sequential)]
    public struct AaProjItem
    {
        public uint ulFlags;             /* specifies valid fields  AADMSPROJF_XXX */
        public int lProjectId;
        public int lEnvironmentId;
        public int lParentId;
        public int lStorageId;
        public int lManagerId;
        public int lTypeId;
        public int lWorkflowId;
        public string lptstrName;
        public string lptstrDesc;
        public int lManagerType;
        public int lWorkspaceProfileId;
        public Guid guidVault;
        public int lComponentClassId;
        public int lComponentInstanceId;
        public uint projFlagMask;        /* specifies valid bits in projFlags */
        public uint projFlags;           /* project flags AADMS_PROJF_XXX     */
    }


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaApi_SelectProjectDataBufferByStruct
(
int lProjectId,    /* i  Parent project id (branch) */
ref AaProjItem lpCriteria     /* i  Project select criteria    */
);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaApi_SelectProjectResourcesDataBuffer(int iprojId,
      int iResType);

    // aaApi_SelectProjectResourcesDataBuffer 
    // aaApi_CopyProjectResources 

    //#define  PROJECTRESOURCES_PROP_PROJECTID   1 
    //  Numeric property. 

    //#define  PROJECTRESOURCES_PROP_RESTYPE   2 
    //  Numeric property. 

    //#define  PROJECTRESOURCES_PROP_RESID   3 
    //  Numeric property. 

    //#define  PROJECTRESOURCES_PROP_FLAGS   4 
    //  Numeric property. 



    public static bool mcmMain_GetDocumentIdByFilePath(string sFileName, int iValidateWithChkl,
        ref int iProjectNo, ref int iDocumentNo)
    {
        bool bRetVal = false;

        Guid docGuid = new Guid();

        if (PWWrapper.aaApi_ChklSetGetDocGuidFromFileName(sFileName,
            ref docGuid))
        {
            if (1 == PWWrapper.aaApi_GUIDSelectDocument(ref docGuid))
            {
                bRetVal = true;

                iProjectNo = PWWrapper.aaApi_GetDocumentNumericProperty(PWWrapper.DocumentProperty.ProjectID, 0);
                iDocumentNo = PWWrapper.aaApi_GetDocumentNumericProperty(PWWrapper.DocumentProperty.ID, 0);
            }
        }

        return bRetVal;
    }

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_CreateRichProject
    (
        ref AaProjItem pProject,               /* io folder data             */
        IntPtr projectInstance,        /* i  optional... not freed by this function */
        bool cloneProjectInstance,   /* i  treat projectInstance as a template, make a replica */
        bool ensureFullAccess,       /* i  creates access controls overriding any inherited access of known types */
                                     /*     for the current user (to permit initial content population) */
        int copyAccessFrom          /* i  project to copy access from or -1 [ignored if (ensureFullAccess == TRUE)] */
    );


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaApi_SelectRichProjectOfFolder(int iProjectId);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_SystemVariableGet(string sVariableName,
        StringBuilder sbValue, int iSbValueSize);

    [DllImport("dmactrl.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_ShowInfoMessage(string sMessage);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_GetDocumentGUIDsByIds([In]int lCount, [In]ref AaDocItem pDocuments,
        [Out] Guid[] docGuids);

    [DllImport("dmawin.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_ImportDocuments([In]uint ulFlags, [In]int lCount, [In]AaDocItem[] pDocuments);

    [DllImport("dmawin.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_ExecuteDocumentCommand([In]uint ulCommandId, [In]int lCount, [In]AaDocItem[] pDocuments);

    [DllImport("dmawin.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectSetMembersDlgExt(IntPtr hWndParent,
      string lpctstrTitle,
      uint ulFlags,
      int lProjectId,
      int lDocumentId,
      int lSetId,
      ref uint lpulOptions,
      ref int lpDocCount,
      [Out] AaDocItem[] ppDocuments
     );

    [DllImport("dmawin.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectSetDlg(IntPtr hWndParent,
      string lpctstrTitle,
      uint ulSetFlags,
      ref int lProjectId,
      ref int lDocumentId
     );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_GetDocumentGUIDsByIds([In]int lCount, [In]AaDocItem[] pDocuments,
        [Out]Guid[] docGuids);



    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaApi_GUIDSelectNestedReferencesDataBuffer(ref Guid masterGuidP,
        ReferenceListFlags flags,
        int maxDepth);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaApi_SelectViewColumns
        (
        int lViewId     /* i  view id    */
        );


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaApi_SelectDocumentDataBuffer
(
int lProjectId,    /* i  project id */
int lDocumentId     /* i  document id    */
);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaApi_GUIDSelectProjectDataBuffer(ref Guid projGuid);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_CreateReferenceInformation2
        (
        UInt64 ui64ElementId,
        ref Guid masterGuid,
        int iMasterModelId,
        ref Guid referenceGuid,
        int iReferenceModelId,
        int iReferenceType, // use 177 for DGN Reference or 145 for Raster
        int iNestDepth, // use 0
        int iFlags // use 2
        );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_DeleteReferenceInformation2
        (
        UInt64 ui64ElementId,
        ref Guid masterGuid,
        int iMasterModelId
        );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_ModifyReferenceInformation2
        (
        UInt64 ui64ElementId,
        ref Guid masterGuid,
        int iMasterModelId,
        ref Guid referenceGuid,
        int iReferenceModelId,
        int iReferenceType, // use 177 for DGN Reference or 145 for Raster
        int iNestDepth, // use 0
        int iFlags // use 2
        );



    /// <summary>
    /// Selects reference information records into buffer id AADMSBUFFER_REFINFO (68) (Same as aaApi_GUIDSelectNestedReferencesDataBuffer()).
    /// Properties are retrieved using PWWrapper.ReferenceInformationProperty...
    /// </summary>
    /// <param name="ui64ElementId">Element ID of attachment</param>
    /// <param name="masterGuid">GUID of master file</param>
    /// <param name="iMasterModelId">Model ID for attachment</param>
    /// <param name="referenceGuid">GUID of reference file</param>
    /// <param name="iReferenceModelId">Model ID attached from reference</param>
    /// <param name="iReferenceType">Reference type is document file type (177 for DGN, 145 for raster ref, 114 for DWG)</param>
    /// <param name="iNestDepth">Use 0 for Next Depth (doesn't seem to have an effect)</param>
    /// <param name="iFlags">2 for raster references, or 4 for reference has moved</param>
    /// <returns></returns>
    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaApi_SelectReferenceInformation2
        (
        UInt64 ui64ElementId,
        ref Guid masterGuid,
        int iMasterModelId,
        ref Guid referenceGuid,
        int iReferenceModelId,
        int iReferenceType, // use 177 for DGN Reference or 145 for Raster
        int iNestDepth, // use 0
        int iFlags // use 2
        );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaApi_SelectUserDataBufferById(int iUserId);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaApi_SelectGroupDataBufferById(int iGroupId);

    [DllImport("dmscli.dll", EntryPoint = "aaApi_GetLinkDataDataBufferColumnValue", CharSet = CharSet.Unicode)]
    private static extern IntPtr unsafe_aaApi_GetLinkDataDataBufferColumnValue(IntPtr hBuf, int iRowIndex, int iColumnIndex);

    public static string aaApi_GetLinkDataDataBufferColumnValue(IntPtr hBuf, int iRowIndex, int iColumnIndex)
    {
        return Marshal.PtrToStringUni(unsafe_aaApi_GetLinkDataDataBufferColumnValue(hBuf, iRowIndex, iColumnIndex));
    }

    [DllImport("dmscli.dll", EntryPoint = "aaApi_GetLinkDataDataBufferStringProperty", CharSet = CharSet.Unicode)]
    private static extern IntPtr unsafe_aaApi_GetLinkDataDataBufferStringProperty(IntPtr hBuf, LinkDataProperty propertyID, int iColumnIndex);

    public static string aaApi_GetLinkDataDataBufferStringProperty(IntPtr hBuf, LinkDataProperty propertyID, int iColumnIndex)
    {
        return Marshal.PtrToStringUni(unsafe_aaApi_GetLinkDataDataBufferStringProperty(hBuf, propertyID, iColumnIndex));
    }

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_AttrFormatParse
    (
       string lpctstrFormat, /* i  Format string                       */
       ref int plDataType,    /* o  Data type(AADMS_ATTRFORM_DATATYPE_*)*/
       uint pulFlags,      /* o  Flags (AADMS_ATTRFORM_FLAG_*)       */
       ref int plWidth,       /* o  Width                               */
       ref int plPrecision,   /* o  Precision                           */
       ref int plAction       /* o  Action (AADMS_ATTRFORM_ACTION_*)    */
    );


    [DllImport("dmscli.dll", EntryPoint = "aaApi_GetLinkDataDataBufferNumericProperty", CharSet = CharSet.Unicode)]
    private static extern int unsafe_aaApi_GetLinkDataDataBufferNumericProperty(IntPtr hBuf, LinkDataProperty propertyID, int iColumnIndex);

    public static int aaApi_GetLinkDataDataBufferNumericProperty(IntPtr hBuf, LinkDataProperty propertyID, int iColumnIndex)
    {
        return (unsafe_aaApi_GetLinkDataDataBufferNumericProperty(hBuf, propertyID, iColumnIndex));
    }

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaApi_SelectLinkDataDataBuffer
(
   int lTableId,               /* i  Table id             */
   ObjectTypeForLinkData lItemType,              /* i  Reference item type  */
   int lItemId1,               /* i  First item id        */
   int lItemId2,               /* i  Second item id       */
   string sWhere,           /* i  Where statement      */
   int lColumnCount,           /* i  Column count         */
   int[] lplColumnIds,           /* i  Columns to fetch     */
   LinkDataSelectFlags ulFlags                 /* i  Flags (AADMSLDSF_*)  */
);
    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GetLinkDataDataBufferColumnCount(IntPtr hBuf);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectAccessUsers
(
uint ulFlags,            /* i  Operation flags          */
int lObjectType,        /* i  Access object type       */
int lObjectId1,         /* i  Access object id 1       */
int lObjectId2,         /* i  Access object id 2       */
int lWorkflowId,        /* i  Workflow id              */
int lStateId,           /* i  State id                 */
uint ulRequiredMask      /* i  Required bits for users  */
);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GetAccessUserNumericProperty
(
AccessUserProperty lPropertyId,   /* i  Property id (ACCESSUSR_PROP_*)  */
int lIndex         /* i  Index of selected item          */
);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_RemoveUserListMember
(
int lUserListId,      /* i  User list id                   */
int lMemberType,      /* i  Member type (AADMS_MGRTYPE_*)  */
int lMemberId         /* i  Member id                      */
);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaApi_SelectUserListMemberDataBufferByProp
(
int lUsrLstId,     /* i  Access user list number      */
int lMemType,      /* i  Member type (-1 for all)      */
int lMemberId      /* i  Member id (-1 for all)       */
);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaApi_SelectUserListDataBufferByProp
(
int lListId,         /* i  List number (-1 for all)  */
int lListType,       /* i  List type (-1 for all)    */
int lOwner           /* i  Owner id (-1 for all)     */
);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaApi_SelectUserListDataBuffer
    (
    );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaApi_SelectUserDataBufferByGroup
    (
        int lGroupId    /* i  Group number     */
    );

    // this is supposed to be obsolete

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_AssignAccessList
(
int lObjectType,      /* i  Object type (AADMSAOTYPE_*)    */
int lObjectId1,       /* i  Object identifier 1            */
int lObjectId2,       /* i  Object identifier 2            */
int lWorkflowId,      /* i  Workflow identifier (0 - none) */
int lStateId,         /* i  State identifier (0 - none)    */
int lMemberType,      /* i  Member type (AADMS_MGRTYPE_*)  */
int lMemberId,        /* i  Member identifier              */
uint lAccessMask       /* i  Access Mask                    */
);

    //added by MDS 20140127
    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_CopyInheritedAccessControl
(
int lObjectType,      /* i  Object type (AADMSAOTYPE_*)    */
int lObjectId1,       /* i  Object identifier 1            */
int lObjectId2,       /* i  Object identifier 2            */
int lWorkflowId,      /* i  Workflow identifier (0 - none) */
int lStateId         /* i  State identifier (0 - none)    */
);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_ModifyAccessItemMask
(
int lObjectType,      /* i  Object type (AADMSAOTYPE_*)    */
int lObjectId1,       /* i  Object identifier 1            */
int lObjectId2,       /* i  Object identifier 2            */
int lWorkflowId,      /* i  Workflow identifier (0 - none) */
int lStateId,         /* i  State identifier (0 - none)    */
int lMemberType,      /* i  Member type (AADMS_MGRTYPE_*)  */
int lMemberId,        /* i  Member identifier              */
uint lAccessMask       /* i  Access Mask                    */
);

    // this is supposed to be obsolete

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_RemoveAccessList
    (
    int lObjectType,     /* i  Object type (AADMSAOTYPE_*)    */
    int lObjectId1,      /* i  Object identifier 1            */
    int lObjectId2,      /* i  Object identifier 2            */
    int lWorkflowId,     /* i  Workflow identifier (0 - none) */
    int lStateId,        /* i  State identifier (0 - none)    */
    int lMemberType,     /* i  Member type (AADMS_MGRTYPE_*)  */
    int lMemberId        /* i  Member identifier              */
    );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_RemoveAccessList
    (
    AccessObjectType lObjectType,     /* i  Object type (AADMSAOTYPE_*)    */
    int lObjectId1,      /* i  Object identifier 1            */
    int lObjectId2,      /* i  Object identifier 2            */
    int lWorkflowId,     /* i  Workflow identifier (0 - none) */
    int lStateId,        /* i  State identifier (0 - none)    */
    int lMemberType,     /* i  Member type (AADMS_MGRTYPE_*)  */
    int lMemberId        /* i  Member identifier              */
    );


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaApi_SelectAccessControlDataBuffer
    (
    AccessControlSelectionFlags ulFlags,           /* i  Flags (AADMSFOAF_*)            */
    AccessObjectType lObjectType,       /* i  Object type (AADMSAOTYPE_*)    */
    int lObjectId1,        /* i  Object identifier 1            */
    int lObjectId2,        /* i  Object identifier 2            */
    int lWorkflowId,       /* i  Workflow identifier (0 - none) */
    int lStateId,          /* i  State identifier (0 - none)    */
    AccessMaskFlags ulRequiredMask     /* i  Required bits for users        */
    );

    public enum AccessObjectTypes : int
    {
        AADMSAOTYPE_ENV_PROJ = 1,      /*< Access control applies on folders or projects. Can be set on an environment.*/
        AADMSAOTYPE_PROJECT = 2,      /*< Access control applies on folders or projects. Can be set on a folder, datasource, workflow or workflow state. */
        AADMSAOTYPE_ENV_DOC = 3,       /*< Access control applies on documents. Can be set on the environment. */
        AADMSAOTYPE_DOCUMENT = 4,      /*< Access control applies on documents. Can be set directly on a document, folder, datasource, workflow or workflow state. */
        AADMSAOTYPE_ODS_COMPONENT = 5  /*< Access control applies engineering components associated with a project. Can be set on a project. */
    }

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern uint aaApi_GetAccessMaskForUser(AccessObjectTypes lObjectType, int lObject1, int lObject2, int lWorkflowId, int lStateId, int lUserId);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaApi_SelectProjectDataBufferFromBranch
    (
    int lProjectId,            /* i  Parent project number  */
    string lpctstrCode,           /* i  Project code           */
    string lpctstrName,           /* i  Project name           */
    string lpctstrDesc,           /* i  Project description    */
    string lpctstrVersion         /* i  Project version        */
    );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaApi_SelectProjectChainDataBuffer
    (
    int lProjectFrom,            /* i  Parent project number  */
    int lProjectTo
    );


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_CreateUserList
    (
        ref int lplUserListId,        /* o  User List Identifier          */
        int lListType,            /* i  User List Type                */
        int lOwnerId,             /* i  User List Owner               */
        string lpctstrName,          /* i  User list name                */
        string lpctstrDesc           /* i  User list description         */
    );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_GetStringSetting
    (
       DatasourceGenericSettings lSettingId,     /* i  Setting Identifier                */
       StringBuilder lptstrValue,    /* o  Value of the Setting              */
       ref int lplBufferLen    /* io Maximum number of chars in buffer */
    );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GetNumericSetting
    (
       DatasourceGenericSettings lSettingId     /* i  Setting Identifier  */
    );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_AddUserListMember
    (
    int lUserListId,      /* i  User list id                   */
    int lMemberType,      /* i  Member type (AADMS_MGRTYPE_*)  */
    int lMemberId         /* i  Member id                      */
    );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectAccessControlItems
    (
    int lObjectType,        /* i  Object type (AADMSAOTYPE_*)    */
    int lObjectId1,         /* i  Object identifier 1            */
    int lObjectId2,         /* i  Object identifier 2            */
    int lWorkflowId,        /* i  Workflow identifier (0 - none) */
    int lStateId,           /* i  State identifier (0 - none)    */
    int lMemberType,        /* i  Member type (AADMS_MGRTYPE_*)  */
    int lMemberId           /* i  Member identifier              */
    );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectAccessControlItems
    (
    AccessObjectType lObjectType,        /* i  Object type (AADMSAOTYPE_*)    */
    int lObjectId1,         /* i  Object identifier 1            */
    int lObjectId2,         /* i  Object identifier 2            */
    int lWorkflowId,        /* i  Workflow identifier (0 - none) */
    int lStateId,           /* i  State identifier (0 - none)    */
    int lMemberType,        /* i  Member type (AADMS_MGRTYPE_*)  */
    int lMemberId           /* i  Member identifier              */
    );


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GetAccessControlItemNumericProperty
    (
       int lPropertyId,     /* i  Property id (ACCE_PROP_*)     */
       int lIndex           /* i  Index of selected user list   */
    );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GetAccessControlItemNumericProperty
    (
       AccessObjectProperty lPropertyId,     /* i  Property id (ACCE_PROP_*)     */
       int lIndex           /* i  Index of selected user list   */
    );


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectGroupMembers
    (
       int lGroupId,   /* i  Group number (-1 for all) */
       int lUserId     /* i  User number               */
    );

    /// <summary>
    /// Used to get member type name (User/Group/AccessList) from member type ID
    /// </summary>
    /// <param name="piType"></param>
    /// <returns></returns>
    public static string GetUserListOrGroupMemberTypeName(int piType)
    {
        string sMemberTypeName = string.Empty;

        if (piType == 1) // User
        {
            sMemberTypeName = "User";
        }
        else if (piType == 2) // Group
        {
            sMemberTypeName = "Group";
        }
        else if (piType == 3) // Access List
        {
            sMemberTypeName = "UserList";
        }
        else if (piType == 4) // All_Users (*Everyone)
        {
            sMemberTypeName = "All_Users (*Everyone)";
        }

        return sMemberTypeName;
    }

    /// <summary>
    /// Used to get the User/Group/AccessList Name from Type and ID
    /// </summary>
    /// <param name="piMemberID"></param>
    /// <param name="piType"></param>
    /// <returns></returns>
    public static string GetUserListOrGroupMemberName(int piMemberID, int piType)
    {
        string sMemberName = string.Empty;

        // PopulateMemberTypeSortedLists();

        if (piType == 1) // User
        {
            if (1 == PWWrapper.aaApi_SelectUser(piMemberID))
                sMemberName = PWWrapper.aaApi_GetUserStringProperty(UserProperty.Name, 0);
        }
        else if (piType == 2) // Group
        {
            if (1 == PWWrapper.aaApi_SelectGroup(piMemberID))
                sMemberName = PWWrapper.aaApi_GetGroupStringProperty(GroupProperty.Name, 0);
        }
        else if (piType == 3) // Access List
        {
            if (1 == PWWrapper.aaApi_SelectUserList(piMemberID))
                sMemberName = PWWrapper.aaApi_GetUserListStringProperty(UserListProperty.Name, 0);
        }
        else if (piType == 4) // All_Users (*Everyone)
        {
            sMemberName = "All_Users (*Everyone)";
        }

        return sMemberName;
    }


    // HAADMSBUFFER aaApi_SelectGroupMemberDataBufferById  ( LONG  lGroupId, LONG lUserId ) 
    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaApi_SelectGroupMemberDataBufferById(int iGroupId, int iUserId);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GetGroupMemberNumericProperty
    (
        int lPropertyIndex,  /* i  Property id                    */
        int lIndex        /* i  Index of selected group member */
    );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GetUserListNumericProperty
    (
        int lPropertyIndex,  /* i  Property id                    */
        int lIndex        /* i  Index of selected user list  */
    );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GetGroupNumericProperty
    (
        int lPropertyIndex,  /* i  Property id                    */
        int lIndex        /* i  Index of selected Group    */
    );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_GetDocumentIdsByGUIDs([In]int lCount, [In] Guid[] docGuids,
        [Out][In] ref AaDocItem pDocuments);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_GetDocumentIdsByGUIDsTest([In]int lCount, [In] Guid[] docGuids,
        [Out][In] ref AaDocItem[] pDocuments);

    // dww - for COT project
    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_GetFileMime(string pFilePath, StringBuilder pMimeType, uint mimeTypeLen);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaOApi_NewClass
    (
        IntPtr lpBase,           /* i  Base class handle  */
        ref IntPtr lppNew           /* o  New class handle   */
    );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaOApi_NewClassPtr
    (
        IntPtr lpBase           /* i  Base class handle  */
    );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaOApi_SetClassName
    (
        string lpctstrClassName,       /* i  New class name  */
        IntPtr lpClass                 /* o  Class handle    */
    );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaOApi_CreateAttribute
    (
       IntPtr lpBase,            /* i  Base attribute handle  */
       ref IntPtr lppNew            /* o  New attribute handle   */
    );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaOApi_SetAttributeName
    (
       IntPtr lpAttr,                /* io Attribute handle    */
       string lpctstrAttrName        /* i  New attribute name  */
    );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaOApi_SetAttributeType
    (
       IntPtr lpAttr,             /* io Attribute handle  */
       ODSAttributeTypes lType               /* i  Type to set       */
    );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaOApi_GetAttributePicklistIds
    (
       int lAttrId,             /* i  Attribute ID                 */
       ref int lplPicklistClassId,  /* o  Picklist class ID            */
       ref int lplPicklistCodeId,   /* o  Picklist code attribute ID   */
       ref int lplPicklistValueId   /* o  Picklist value attribute ID  */
    );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaOApi_SetAttributeCmnProps
    (
       IntPtr lpAttr,             /* io Attribute handle     */
       string lpctstrDesc,        /* i  New description      */
       ref int lplVisibility,      /* i  New visibility       */
       ref int lplDataType,        /* i  New data type        */
       ref int lplDataLen          /* i  New max data length  */
    );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaOApi_SaveAttribute
    (
       ref IntPtr lppAttr       /* i  Attribute handle to save  */
    );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaOApi_SetAttributeDbProps
    (
       IntPtr lpAttr,              /* io Attribute handle     */
       ref int lplControl,          /* i  Attribute control    */
       string lpctstrInstCol       /* i  Column name          */
    );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaOApi_FindAttributeByName
    (
       ref IntPtr lppAttr,            /* o  Attribute handle        */
       string lpctstrAttrName      /* i  Attribute name to find  */
    );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaOApi_AddClassAttribute
    (
       IntPtr lpAttr,            /* i  Attribute handle    */
       int lAttrId,           /* i  Attribute ID        */
       int lIndex,            /* i  Position to add to  */
       IntPtr lpClass            /* o  Class handle        */
    );


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaOApi_LinkInstToDoc
    (
        IntPtr lpInstance,     /* i  Instance           */
        int lProjNo,        /* i  Project number     */
        int lDocNo          /* i  Document number    */
    );




    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaOApi_RemoveLink
    (
        int lClassId,           /* i  Link class ID    */
        IntPtr lpFromInst,         /* i  'From' instance  */
        IntPtr lpToInst            /* i  'To' instance    */
    );


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaOApi_SetLink
    (
        int lClassId,           /* i  Link class ID    */
        IntPtr lpFromInst,         /* i  'From' instance  */
        IntPtr lpToInst            /* i  'To' instance    */
    );


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaOApi_GetLinks
    (
    int lClassId,         /* i  Link class ID          */
    IntPtr lpInst,           /* i  Instance handle        */
    bool bFromFlag,        /* i  'From' instance given  */
    ref IntPtr lpppLinks,      /* o  Links to/from given    */
                               /*    instance               */
    ref int lpLinkCount       /* o  Count of found links   */
    );


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaOApi_FindClassPtrByName(string sClassName);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaOApi_FindClassPtr(int iClassId);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaOApi_GetClassAttrCount(IntPtr classPtrP, bool bAll);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaOApi_GetClassAttrId(IntPtr classPtrP, int iAttrId, int iIndex, ref int iAddOnId, ref int iParentId);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaOApi_FindAttributePtr(int iAttrId);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaOApi_SetClassProps
    (
       IntPtr lpClass,           /* io Class handle              */
       string lpctstrClassDesc,  /* i  Class description         */
       ref int lplSystem,         /* i  Class type                */
       ref int lplKeyId,          /* i  Primary key attribute ID  */
       ref int lplClAttrId,       /* i  Primary key attribute ID  */
       string lpctstrTblName,    /* i  Class primary table name  */
       string lpctstrSeqName,    /* i  Sequence generator name   */
       string lpctstrCatName,    /* i  Catalog table name        */
       int lplcatKeyId        /* i  Catalog key attribute ID  */
    );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaOApi_FindQualifierPtrByName(string sQualName);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaOApi_GetQualifierId(IntPtr pQual);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaOApi_AddClassQualifier
    (
       IntPtr lpClass,             /* io Class handle     */
       int lQualId,             /* i  Qualifier ID     */
       ref int lpValBuf             /* i  Qualifier value  */
    );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaOApi_GetClassesByQualId(int iQualId, ref IntPtr iClassIdArrayP);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaOApi_FindQualifierByName
    (
       string lpctstrQualName,     /* i  Qualifier name    */
       ref IntPtr lppQual             /* o  Qualifier handle  */
    );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaOApi_GetAttributeStringProperty
    (
    IntPtr lpAttr,                   /* i  Attribute handle  */
    ODSAttributeProperty lAttrPropId,              /* i  Property ID       */
    StringBuilder lptstrReturnBuffer,       /* o  Buffer for value  */
    int lSize                     /* i  Size of buffer    */
    );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaOApi_ModifyAttributeQualifier
    (
        int iQualifierId, /* Qualifier ID*/
        ref int iQualifierVal, /* Qualifier value */
        IntPtr lpAttr      /* i  Attribute handle  */
    );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaOApi_GetAttributeQualifierProperties(IntPtr lpAttr,      /* i  Attribute handle  */
      int lQualId,
      int lIndex,
      ref int lplQualId,
      ref int lplDataType,
      ref int lppValBuf
     );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaOApi_ModifyAttributeQualifier
    (
        int iQualifierId, /* Qualifier ID*/
        string sQualifierVal, /* Qualifier value */
        IntPtr lpAttr      /* i  Attribute handle  */
    );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaOApi_RemoveAttributeQualifier
    (
        int iQualifierId, /* Qualifier ID*/
        IntPtr lpAttr      /* i  Attribute handle  */
    );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaOApi_GetClassStringProperty
    (
    IntPtr lpClass,                  /* i  Class handle      */
    ODSClassProperty lClassPropId,             /* i  Property ID       */
    StringBuilder lptstrReturnBuffer,       /* o  Buffer for value  */
    int lSize                     /* i  Size of buffer    */
    );


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaOApi_GetClassNumericProperty
    (
    IntPtr lpClass,                  /* i  Class handle      */
    ODSClassProperty lClassPropId             /* i  Property ID       */
    );


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaOApi_GetAttributeDbProps
    (
    IntPtr lpAttr,             /* i  Attribute handle   */
    ref int lpControl,          /* o  Attribute control  */
    StringBuilder lpInstCol           /* o  Column name        */
    );

    // dww - for new import tool
    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaOApi_GetAttributeLabel
    (
    IntPtr lpAttr,             /* i  Attribute handle   */
    int lIntfId,                /* i Id of the interface to use */
    StringBuilder lptstrReturnBuffer,       /* o  Buffer for value  */
    int lSize                     /* i  Size of buffer    */
    );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaOApi_GetAttributeNumericProperty
    (
    IntPtr lpAttr,              /* i  Attribute handle  */
    ODSAttributeProperty lAttrPropId          /* i  Property ID       */
    );


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaOApi_GetClassId(IntPtr pClass);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaOApi_GetLinkToInstance(IntPtr pClass);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaOApi_GetLinkFromInstance(IntPtr pClass);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaOApi_InstanceNewQuery();


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaOApi_InstanceSetQueryClass
    (
    IntPtr lpQuery,         /* io Query handle             */
    int lClassId,        /* i  Class ID                 */
    int lVersionStatus,  /* i  Version status           */
    string lpctstrWhere,    /* i  Instance 'Where' string  */
    IntPtr lpFromInst,      /* i  'From' instance for      */
                            /*    link class               */
    IntPtr lpToInst         /* i  'To' instance for link   */
                            /*    class                    */
    );


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaOApi_InstanceCloseQuery(IntPtr pQuery);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaOApi_InstanceDeleteQuery(IntPtr pQuery);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaOApi_InstanceStartQuery(IntPtr pQuery);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaOApi_InstanceFetchQuery(ref IntPtr pInst, IntPtr pQuery);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaOApi_FreeInstance(IntPtr pInst);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaOApi_GetInstanceId(IntPtr pInst);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaOApi_GetInstanceClassId(IntPtr pInst);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaOApi_FindClass(ref IntPtr pClass, int iClassId);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaOApi_GetClassProps
    (
    IntPtr lpClass,          /* i  Class handle              */
    ref int lplClassId,       /* o  Class ID                  */
    StringBuilder lptstrClassName,  /* o  Class name                */
    ref int lplIsVersion,     /* o  Is class versionized      */
    StringBuilder lptstrClassDesc,  /* o  Class description         */
    ref int lplSystem,        /* o  Class type                */
    ref int lplKeyId,         /* o  Primary key attribute ID  */
    ref int lplClAttrId,      /* o  Primary control           */
                              /*    attribute ID              */
    StringBuilder lptstrTblName,    /* o  Class primary table name  */
    StringBuilder lptstrSeqName,    /* o  Sequence generator name   */
    StringBuilder lptstrCatName,    /* o  Catalog table name        */
    ref int lplCatKeyId,      /* o  Catalog key attribute ID  */
    StringBuilder lptstrModTime     /* o  Modification time         */
    );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaOApi_VersionizeClass(ref IntPtr pClass);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaOApi_FindAttribute(ref IntPtr pAttr, int iAttrId);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaOApi_GetInstanceNumAttrs(IntPtr pInst, int iAttrType, bool bVisible, ref int lNumAttrs);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaOApi_GetInstanceAttrId
    (
    IntPtr lpaaOdsInst,   /* i  Instance handle           */
    int lIndex,        /* i  Attribute index           */
    int lAttrType,     /* i  Attribute type            */
    bool bVisible,      /* i  Check visible attributes  */
    ref int lplAttrId,     /* o  Attribute ID              */
    ref int lplAttrType,   /* o  Attribute type            */
    ref int lplAddonId,    /* o  Addon of attribute        */
    ref int lplParentId,   /* o  Parent of attribute       */
    ref int lplVisibility  /* o  Visibility flag           */
    );


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaOApi_GetAttributeCmnProps
    (
    IntPtr lpAttr,           /* i  Attribute handle           */
    ref int lpAttrId,         /* o  Attribute ID               */
    StringBuilder lpctstrName,      /* o  Attribute name             */
    StringBuilder lpctstrDesc,      /* o  Attribute description      */
    ref int lpVisibility,     /* o  Attribute visibility       */
    ref int lpDataType,       /* o  Attribute data type        */
    ref int lpDataLen         /* o  Max attribute data length  */
    );



    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaOApi_GetInstanceAttrStrValue
    (
    IntPtr lpAAOdsInstance,   /* i  Instance handle       */
    int lAttrId,           /* i  Attribute ID          */
    int lArrayIndex,       /* i  Value array index     */
    StringBuilder lptstrValue,       /* o  Value                 */
    int lSize              /* i  Size of value buffer  */
    );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaOApi_LoadAttributePickList
    (
       IntPtr lpInstance,           /* Specifies the instance containing the context sensitivity attribute. This value can be NULL, if context sensitivity is not wanted.*/
       int lAttributeId,            /* Specifies the attribute id, whose picklist class id to obtain.         */
       bool bSorted                 /* Specifies whether to sort the values or not. When set to TRUE the picklist values are sorted.  */
    );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaOApi_LoadAllAttributes();

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaOApi_LoadAttribute(ref IntPtr attrP, int iAttrId);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaOApi_GetAttributePickListItemCount
    (
       IntPtr hPickList          /* Specifies a handle to the loaded picklist. */
    );

    [DllImport("dmscli.dll", EntryPoint = "aaOApi_GetAttributePickListItemValue", CharSet = CharSet.Unicode)]
    private static extern IntPtr __aaOApi_GetAttributePickListItemValue
    (
       IntPtr hPickList,         /* Specifies a handle to the loaded picklist. */
       int lPickIndex            /* Specifies an index pointing to one of the items in the picklist. */
    );

    public static string aaOApi_GetAttributePickListItemValue
    (
       IntPtr hPickList,         /* Specifies a handle to the loaded picklist. */
       int lPickIndex            /* Specifies an index pointing to one of the items in the picklist. */
    )
    {
        return Marshal.PtrToStringUni(__aaOApi_GetAttributePickListItemValue(hPickList, lPickIndex));
    }

    [DllImport("dmscli.dll", EntryPoint = "aaOApi_GetAttributePickListItemCode", CharSet = CharSet.Unicode)]
    private static extern IntPtr __aaOApi_GetAttributePickListItemCode
    (
       IntPtr hPickList,         /* Specifies a handle to the loaded picklist. */
       int lPickIndex            /* Specifies an index pointing to one of the items in the picklist. */
    );

    public static string aaOApi_GetAttributePickListItemCode
    (
       IntPtr hPickList,         /* Specifies a handle to the loaded picklist. */
       int lPickIndex            /* Specifies an index pointing to one of the items in the picklist. */
    )
    {
        return Marshal.PtrToStringUni(__aaOApi_GetAttributePickListItemCode(hPickList, lPickIndex));
    }


    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct AAODSPickList
    {
        public int lPicklistType;                /**< picklist type (AAODS_PICKLIST_TYPE_*) */
        public int lPicklistClassId;             /**< class of picklist (AAODS_PICKLIST_TYPE_CLASS)                        */
        public int lPicklistCodeAttrId;          /**< attribute id of picklist code column (AAODS_PICKLIST_TYPE_CLASS)     */
        public int lPicklistValueAttrId;         /**< attribute id of picklist value column (AAODS_PICKLIST_TYPE_CLASS)    */
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 1025)]
        public string sSelect;           /**< SQL SELECT statement (AAODS_PICKLIST_TYPE_SELECT)                    */
        public bool bForceToList;                /**< force to list */
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 1025)]
        public string sFileName;         /**< DLL file name */
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 1025)]
        public string sFunction;         /**< function name */
        public int lUpdateOnEditSiblingsFlags;   /**< AAQUALID_PICKLIST_UPDATE_ON_EDIT_SIBLING flags*/
    }

    public enum AAODSPickListType : int
    {
        None = 0,
        Class = 1,
        Select = 2
    }

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaOApi_GetAttributePicklistDefinition
    (
        int lAttrId,
        ref AAODSPickList lpPickList
    );


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaOApi_ClassGetBusinessKeyAttrId(IntPtr lpClass);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaOApi_GetClassAttributes
    (
    IntPtr lpClass,      /* i  Class handle                  */
    int lAttrId,      /* i  Attribute ID to retrieve      */
    int lIndex,       /* i  Index of attribute in class   */
    ref int lpAttrId,     /* o  Attribute ID                  */
    ref int lpAddonId,    /* o  Addon class ID of attribute   */
    ref int lpParentId    /* o  Parent class ID of attribute  */
    );


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaOApi_NewInstance
    (
    int lClassId,         /* i  Class ID                        */
    int lVersionStatus,   /* i  Version status                  */
    int[] lplIncAddons,     /* i  Addons to add                   */
    int lNumIncAddons,    /* i  Number of addons to add         */
    int[] lplIncAttrs,      /* i  Attributes to add               */
    int lNumIncAttrs,     /* i  Number of attributes to add     */
    bool bDefaultFlag      /* i  Initialize with default values  */
    );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaOApi_SetInstanceAttrStrValue
    (
        IntPtr lpAAOdsInstance, //Specifies the instance containing the attribute. If the function call was successful the modified instance is stored in this parameter.
        String lpctstrValue, //Pointer to a null-terminated string containing the string value to set.
        int lAttrId, //Specifies the attribute id of the attribute to modify in the instance.
        int lArrayIndex //Specifies an index pointing to the attribute, whose value to set, if the attribute id is in an array.
    );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaOApi_SetInstanceAttrStrValueExt
    (
    IntPtr lpAAOdsInstance,   /* io Instance handle      */
    string lpctstrValue,      /* i  Value to set         */
    int lAttrId,           /* i  Attribute ID         */
    int lArrayIndex,        /* i  Value array index    */
    bool bValidate           /*i   Validate attribute  */
    );


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaOApi_SetInstanceAttrValue
    (
    IntPtr lpAAOdsInstance,     /* io Instance handle    */
    ref int lpVoid,              /* i  Value to set       */
    int lSize,               /* i  Size of value      */
    int lAttrId,             /* i  Attribute ID       */
    int lArrayIndex          /* i  Value array index  */
    );


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaOApi_SaveInstance
    (
        IntPtr lpAAOdsInstance   /* io Instance handle      */
    );


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaOApi_LoadInstanceByIds
    (
    int lClassId,                 /* i  Class ID     */
    int lInstId,                  /* i  Instance ID  */
    int lVerId                    /* i  Version ID   */
    );



    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaOApi_DeleteInstance(IntPtr instanceP);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaOApi_FindAttributePtrByName(string wcAttrName);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaOApi_GetAttributeId(IntPtr attrP);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaOApi_Initialize();


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaOApi_InitializeSession();


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaOApi_LoadAllClasses(int iHierarchyId);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaOApi_GetLoadedClassCount();


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaOApi_GetLoadedClassPtr(int iIndex);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaOApi_SaveClass
    (
       ref IntPtr lppClass          /* i  Class handle  */
    );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaOApi_FindClassByName
    (
        string lpctstrName,          /* i  Class name to find  */
        ref IntPtr lppClass             /* o  Class handle        */
    );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern string aaOApi_GetHrcyStringProperty
    (
        int lPropertyId,      /* i  Property ID             */
        int lIndex            /* i  Selected storage index  */
    );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaOApi_SelectHrchy
    (
       int lHrchyId,             /* i  Hierarchy ID           */
       string lpctstrName,          /* i  Hierarchy name         */
       string lpctstrDecr,          /* i  Hierarchy description  */
       string lpctstrModTime,       /* i  Modification time      */
       int lSelMask              /* i  Selection mask         */
    );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaOApi_GetHrcyId
     (
        int lIndex               /* i  List index  */
     );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaOApi_HrchyAddClass
    (
       IntPtr lpClass,         /* i  Class handle         */
       IntPtr lpParent,        /* i  Parent class handle  */
       int lHrchyId         /* i  Hierarchy ID         */
    );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaOApi_ClassSyncDataBase
    (
       IntPtr lpClass,           /* i  Class handle         */
       int iCreateIndex       /* i  Create index of key  */
                              /*    attributes           */
    );
    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaOApi_FreeClass
    (
       IntPtr lpClass           /* i  Class handle  */
    );

    //AAOAPI_Funtion end here

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_CopyProject
    (
        int lSourceProjectId, /* i  Project to be copied    */
        int lTargetProjectId, /* i  Target project number   */
        ProjectCopyDeleteAndExportFlags ulFlags,          /* i  Copy mask(AAPRO_ARRAY_*)*/
        IntPtr fpCallBack,       /* i  Callback func. address  */
        IntPtr aaUserParam,      /* i  User defined callback   function parameter      */
        ref int lplCount          /* o  Count of copied projects*/
    );

    // dww 2013-10-02
    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_CopyProjectWithHierarchy
    (
        int lSourceProjectId, /* i project to be copied                     */
        int lTargetProjectId, /* i target project number                    */
        ProjectCopyDeleteAndExportFlags ulRootFlags,      /* i copy mask for target folder              */
        ProjectCopyDeleteAndExportFlags ulHierarchyFlags, /* i copy mask for subfolders                 */
        IntPtr fpCallBack,       /* i address of the Call back function        */
        IntPtr aaUserParam,      /* i user defined callback function parameter */
        ref int lplCount          /* o Count of copied projects                 */
        );


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_CopyProjectResources
        (
        int sourceProjectId,
        ProjectResourceTypes resType,
        int resId,
        int targetProjectId,
        ref int plCount
        );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_SetDefaultProjectView(int iUserId /* 0 - global */,
        int iProjectId,
        int iViewId);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_SetDefaultView(int iUserId /* 0 - global */,
        ref Guid objectTypeGuid /* see dmspublic.h */,
        int iObjectId,
        ref Guid viewTypeGuid /* see dmspublic.h */,
        int iViewId
        );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_RetrieveDefaultView(int iUserId /* 0 - global */,
        ref Guid objectTypeGuid /* see dmspublic.h - 67536006-10e4-4428-86b1-bbcc11699ded for project */,
        int iObjectId,
        ref Guid viewTypeGuid /* see dmspublic.h - 2073fbc1-2fbf-496f-9de0-015ab4010e94 for document list view | b7312de2-d64b-4c40-9369-a14c2f46d2ef for preview pane*/
        );

    public static int GetDefaultProjectView(int iProjectId, bool bGetPreviewPaneView)
    {
        Guid viewTypeGuid = new Guid("2073fbc1-2fbf-496f-9de0-015ab4010e94");
        Guid projectObjectTypeGuid = new Guid("67536006-10e4-4428-86b1-bbcc11699ded");

        if (bGetPreviewPaneView)
            viewTypeGuid = new Guid("b7312de2-d64b-4c40-9369-a14c2f46d2ef");

        return aaApi_RetrieveDefaultView(0, ref projectObjectTypeGuid, iProjectId, ref viewTypeGuid);
    }

    public static bool SetDefaultProjectView(int iProjectId, int iViewId, bool bSetPreviewPaneView)
    {
        Guid viewTypeGuid = new Guid("2073fbc1-2fbf-496f-9de0-015ab4010e94");
        Guid projectObjectTypeGuid = new Guid("67536006-10e4-4428-86b1-bbcc11699ded");

        if (bSetPreviewPaneView)
            viewTypeGuid = new Guid("b7312de2-d64b-4c40-9369-a14c2f46d2ef");

        return aaApi_SetDefaultView(0, ref projectObjectTypeGuid, iProjectId, ref viewTypeGuid, iViewId);
    }

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GetProjectDefaultPreviewPaneView(int iProjectId);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GetProjectDefaultView(IntPtr projectBuffer,
        int rowIndex);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_ModifyProject
(
int lProjectId,     /* i  Project number to modify       */
int lStorageId,     /* i  Storage number                 */
int lManagerId,     /* i  Project manager number         */
int lType,          /* i  Project type                   */
string lpctstrName,    /* i  Project name                   */
string lpctstrDesc    /* i  Project description            */
);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_IsConnectionLost();


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_HasAdminSetup();


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_CreateUser(ref int iUserId, string sUserType,
        string sSecurityProvider,
        string sUserName, string sNewPassword, string sNewDescription,
        string sEMail);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_CreateUser2(ref int iUserId, 
        string sUserType,
        string sSecurityProvider,
        string sUserName, 
        string sNewPassword, 
        string sNewDescription,
        string sEMail,
        string sIdentity);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GetLinkDataColumnCount();


    [DllImport("dmscli.dll", EntryPoint = "aaApi_GetLinkNumericProperty", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GetLinkNumericProperty(LinkProperty propertyID, int index);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectUserList(int iUserListId);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_DeleteUserById(int iUserId);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_DeleteUserByName(string sUserName);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaApi_DmsDataBufferSelect(int DsmBufferType);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaApi_SQueryDataBufferSelectAll();

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaApi_SQueryDataBufferSelect(int iQueryId);

    //HAADMSBUFFER aaApi_SQueryDataBufferSelectSubItems2(LONG lParQueryId, LONG lUserId, LONG lProjectId )
    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaApi_SQueryDataBufferSelectSubItems2(int lParQueryId, int lUserId, int lProjectId);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaApi_SQueryCriDataBufferSelect(int iQueryId);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_DmsDataBufferGetCount(IntPtr hDataBuffer);

    //#define SQRY_PROP_QUERYID                   1
    //#define SQRY_PROP_USERID                    2
    //#define SQRY_PROP_PQUERYID                  3
    //#define SQRY_PROP_HASCRITERIA               4
    //#define SQRY_PROP_HASSUBITEMS               5
    //#define SQRY_PROP_NAME                      6
    //#define SQRY_PROP_DESC                      7
    //#define SQRY_PROP_PROJECTID                 8
    //#define SQRY_PROP_FROMTYPE                  9

    //#define SQRYC_PROP_QUERYID                1     /**< \b Numeric property. Query identifier. */
    //#define SQRYC_PROP_CRITERIONID            2     /**< \b Numeric property. Criterion identifier. */
    //#define SQRYC_PROP_ORGROUPNUMBER          3     /**< \b Numeric property. OR group number. */
    //#define SQRYC_PROP_FLAGS                  4     /**< \b Numeric property. Flags from \ref aadmsdef_SearchFunctionalityDefinitions_QueryCriteriaconstants_Searchcriteriaflags. */
    //#define SQRYC_PROP_PROPERTYSET            5     /**< \b Guid property. Property set identifier. */
    //#define SQRYC_PROP_PROPERTYNAME           6     /**< \b String property. Property name. */
    //#define SQRYC_PROP_PROPERTYID             7     /**< \b Numeric property. Property identifier. */
    //#define SQRYC_PROP_RELATION               8     /**< \b Numeric property. Relation. Values from \ref aadmsdef_SearchFunctionalityDefinitions_QueryCriteriaconstants_PropertyRelationsValues. */
    //#define SQRYC_PROP_FIELDTYPE              9     /**< \b Numeric property. Value type from \ref aadmsdef_FormatUtilityDefinitions_AttributeTypes. */
    //#define SQRYC_PROP_FIELDVALUE            10     /**< \b String property. Value. */

    public enum SavedQueryProperty : int
    {
        SQRY_PROP_QUERYID = 1,
        SQRY_PROP_USERID = 2,
        SQRY_PROP_PQUERYID = 3,
        SQRY_PROP_HASCRITERIA = 4,
        SQRY_PROP_HASSUBITEMS = 5,
        SQRY_PROP_NAME = 6,
        SQRY_PROP_DESC = 7,
        SQRY_PROP_PROJECTID = 8,
        SQRY_PROP_FROMTYPE = 9
    }

    public enum SavedQueryCriterionProperty : int
    {
        SQRYC_PROP_QUERYID = 1,
        SQRYC_PROP_CRITERIONID = 2,
        SQRYC_PROP_ORGROUPNUMBER = 3,
        SQRYC_PROP_FLAGS = 4,
        SQRYC_PROP_PROPERTYSET = 5,
        SQRYC_PROP_PROPERTYNAME = 6,
        SQRYC_PROP_PROPERTYID = 7,
        SQRYC_PROP_RELATION = 8,
        SQRYC_PROP_FIELDTYPE = 9,
        SQRYC_PROP_FIELDVALUE = 10
    }


    [DllImport("dmscli.dll", EntryPoint = "aaApi_DmsDataBufferGetNumericProperty", CharSet = CharSet.Unicode)]
    public static extern int aaApi_DmsDataBufferGetNumericProperty(IntPtr hDataBuffer, int lPropertyId, int lIdxRow);

    [DllImport("dmscli.dll", EntryPoint = "aaApi_DmsDataBufferGetUint64Property", CharSet = CharSet.Unicode)]
    public static extern UInt64 aaApi_DmsDataBufferGetUint64Property(IntPtr hDataBuffer, int lPropertyId, int lIdxRow);


    [DllImport("dmscli.dll", EntryPoint = "aaApi_DmsDataBufferGetStringProperty", CharSet = CharSet.Unicode)]
    private static extern IntPtr unsafe_aaApi_DmsDataBufferGetStringProperty(IntPtr hDataBuffer, int lPropertyId, int lIdxRow);


    public static string aaApi_DmsDataBufferGetStringProperty(IntPtr hDataBuffer, int lPropertyId, int lIdxRow)
    {
        return Marshal.PtrToStringUni(unsafe_aaApi_DmsDataBufferGetStringProperty(hDataBuffer, lPropertyId, lIdxRow));
    }

    [DllImport("dmscli.dll", EntryPoint = "aaApi_DmsDataBufferGetGuidProperty", CharSet = CharSet.Unicode)]
    private static extern IntPtr unsafe_aaApi_DmsDataBufferGetGuidProperty(IntPtr hDataBuffer, int lPropertyId, int lIdxRow);


    public static Guid aaApi_DmsDataBufferGetGuidProperty(IntPtr hDataBuffer, int lPropertyId, int lIdxRow)
    {
        return (Guid)Marshal.PtrToStructure(unsafe_aaApi_DmsDataBufferGetGuidProperty(hDataBuffer, lPropertyId, lIdxRow), Type.GetType("System.Guid"));
    }

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_GUIDGetDocumentNamePath(ref Guid guid, bool UseDesc,
        char tchSeparator, StringBuilder StringBuffer, int BufferSize);


    [DllImport("dmscli.dll", EntryPoint = "aaApi_DmsThreadBufferGetStringProperty", CharSet = CharSet.Unicode)]
    private static extern IntPtr unsafe_aaApi_DmsThreadBufferGetStringProperty(int lBufferId, int lPropertyId, int lIdxRow);


    public static string aaApi_DmsThreadBufferGetStringProperty(int lBufferId, int lPropertyId, int lIdxRow)
    {
        return Marshal.PtrToStringUni(unsafe_aaApi_DmsThreadBufferGetStringProperty(lBufferId, lPropertyId, lIdxRow));
    }


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_ModifyUserExt(int iUserId, string sUserTypeDorW, string sSecProvider, string sName, string sPassword,
        string sDesc, string sEmail);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_ModifyUser(int iUserId, string sName, string sPassword,
        string sDesc, string sEmail);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_DeleteUserExt(int iUserId, int iUserIdForItems);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectAllUsers();

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaApi_SelectAllViews(ref Guid pViewTypeGuid);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_VerifyUser(string userName, string password);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GetUserNumericProperty(UserProperty lPropertyId, int lIndex);


    [DllImport("dmscli.dll", EntryPoint = "aaApi_GetUserStringProperty", CharSet = CharSet.Unicode)]
    private static extern IntPtr unsafe_aaApi_GetUserStringProperty(UserProperty lPropertyId, int lIndex);


    public static string aaApi_GetUserStringProperty(UserProperty lPropertyId, int lIndex)
    {
        return Marshal.PtrToStringUni(unsafe_aaApi_GetUserStringProperty(lPropertyId, lIndex));
    }

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_SetUserNumericSetting
(
   int lParam,           /* i  Parameter class to set  */
   int lParamValue       /* i  Parameter value to set  */
);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_SetUserNumericSettingByUser
(
   int lUserNo,          /* i  User number             */
   int lParam,           /* i  Parameter class to set  */
   int lParamValue       /* i  Parameter value to set  */
);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GetUserNumericSettingByUser
(
   int lUserNo,         /* i  User number              */
   int lParam           /* i  Parameter class to get   */
);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_GetUserStringSettingByUser(int lUserNo,
      int lParam,
      StringBuilder lptstrParam,
      int lParamLength
     );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_SetUserStringSettingByUser(int lUserNo,
      int lParam,
      string lpctstrParam
     );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GetUserNumericSetting
(
   int lParam           /* i  Parameter class to get   */
);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_SaveUserSettings();

    [DllImport("dmsgen.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GetLastErrorId();


    [DllImport("dmsgen.dll", EntryPoint = "aaApi_GetLastErrorDetail", CharSet = CharSet.Unicode)]
    public static extern IntPtr unsafe_aaApi_GetLastErrorDetail();


    public static string aaApi_GetLastErrorDetail()
    {
        return Marshal.PtrToStringUni(unsafe_aaApi_GetLastErrorDetail());
    }


    [DllImport("dmsgen.dll", EntryPoint = "aaApi_GetMessageByErrorId", CharSet = CharSet.Unicode)]
    private static extern IntPtr unsafe_aaApi_GetMessageByErrorId(int errorID);


    public static string aaApi_GetMessageByErrorId(int errorID)
    {
        return Marshal.PtrToStringUni(unsafe_aaApi_GetMessageByErrorId(errorID));
    }

    // dww - 2014-05-14 for COT project
    [DllImport("dmsgen.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_CleanFilePath([System.Runtime.InteropServices.MarshalAsAttribute(System.Runtime.InteropServices.UnmanagedType.LPWStr)] System.Text.StringBuilder lptstrFileName);

    [DllImport("dmsgen.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_Free(IntPtr pointer);

    public enum SetType : int
    {
        Unknown = 0,
        Flat = 2,
        Logical = 3
    }


    public enum SetRelationType : int
    {
        Sibling = 2,
        Redline = 3,
        Reference = 4
    }


    public enum TypeMask : int
    {
        Flat = 0x00010000,
        Logical = 0x00020000,
        Redline = 0x00080000,
        Ref = 0x00100000,
        All = 0x001F0000
    }


    public enum SetProperty : int
    {
        ID = 1,
        MemberId = 2,
        Type = 3,
        ParentProjectId = 4,
        ParentItemId = 5,
        ChildProjectId = 6,
        ChildItemId = 7,
        Relation = 8,
        Transfer = 9,
        SetProjectID = 10,
        SetItemId = 11,
        SDocGuid = 12,
        PDocGuid = 13,
        CDocGuid = 14
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct DocumentCreateParam
    {
        public uint ulMask;
        public int lProjectId;
        public int lDocumentId;
        public int lFileType;
        public int lItemType;
        public int lApplicationId;
        public int lDepartmentId;
        public string lpctstrFileName;
        public string lpctstrName;
        public StringBuilder lptstrWorkingFile;
        public int lBufferSize;
        public uint ulFlags;
        public int lWorkspaceProfileId;
        public bool bLeaveOut;
        public int lAttributeId; // supposed to out (might need to be ref?)
        public Guid guidProject;
        public Guid guidDocument; // supposed to out (might need to be ref?)
    }


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_CreateDocument2(ref DocumentCreateParam docCreateParam);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_GetConnectionInfo2(IntPtr hDataSource, ref bool lpbODBC,
                ref int lplNativeType, ref int lplLoginType, StringBuilder lptstrName,
                int lLenName, StringBuilder lptstrUser, int lLenUser, StringBuilder lptstrSchema, int lLenSchema);

    // still available in v8i
    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_GetConnectionInfo(IntPtr hDataSource, ref bool lpbODBC,
                ref int lplNativeType, ref int lplLoginType, StringBuilder lptstrName,
                int lLenName, StringBuilder lptstrUser, int lLenUser, StringBuilder lptstrPassword,
                int lLenPassword, StringBuilder lptstrSchema, int lLenSchema);


    [DllImport("dmscli.dll", EntryPoint = "aaApi_CopyDocumentAttributes", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_CopyDocumentAttributes
       (
       int lSourceProjectId,     /* i  Source project id    */
       int lSourceDocumentId,    /* i  Source document id   */
       int lTargetProjectId,     /* i  Target project id    */
       int lTargetDocumentId,    /* i  Target document id   */
       AttributeCopyFlags ulFlags               /* i  Operation flags      */
       );

    [DllImport("dmscli.dll", EntryPoint = "aaApi_GUIDCopyDocumentAttributes", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_GUIDCopyDocumentAttributes
       (
        ref Guid pSourceDocGuid,
        ref Guid pTargetDocGuid,
       AttributeCopyFlags ulFlags               /* i  Operation flags      */
       );


    [DllImport("dmscli.dll", EntryPoint = "aaApi_DeleteDocumentAttributes", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_DeleteDocumentAttributes
       (
       int lProjectId,     /* i  Source project id    */
       int lDocumentId
       );

    [DllImport("dmscli.dll", EntryPoint = "aaApi_GUIDDeleteDocumentAttributes", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_GUIDDeleteDocumentAttributes
    (
    ref Guid pDocGuid
    );


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectSetByTypeMask(TypeMask ulTypeMask, int lSetId,
        int lParentProjectId, int lParentDocumentId, int lChildProjectId, int lChildDocumentID);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GetSetNumericProperty(SetProperty lPropertyId, int lIndex);

    [DllImport("dmscli.dll", EntryPoint = "aaApi_GetSetStringProperty", CharSet = CharSet.Unicode)]
    public static extern IntPtr unsafe_aaApi_GetSetStringProperty(SetProperty iSetProp, int iIndex);

    public static string aaApi_GetSetStringProperty(SetProperty iSetProp, int iIndex)
    {
        return Marshal.PtrToStringUni(unsafe_aaApi_GetSetStringProperty(iSetProp, iIndex));
    }

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_UpgradeFolderToRichProject
    (
        ref AaProjItem projectItem,                /* i  folder data (same or modified, id is required) */
        IntPtr projectInstance,        /* i  optional... not freed by this function */
        bool cloneProjectInstance    /* i  treat projectInstance as a template, make a replica */
    );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_SetParentProject
    (
       int lChildId,     /* i  Child project number           */
       int lParentId     /* i  Parent project number (<0 top) */
    );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_GUIDSetParentProject
    (
        Guid pProjectGuid,
        Guid pParentProjectGuid
    );

    public static int GetProjectNoFromPath(string pwPath)
    {
        return ProjectNoFromPath(pwPath);
    }

    public static int GetProjectNoFromPathNoCase(string pwPath)
    {
        return ProjectNoFromPath(pwPath, true);
    }


    [Flags]
    public enum DocumentCopyFlags : uint
    {
        AADMS_DOCCOPY_CAN_OVERWRITE = 0x00000001,
        AADMS_DOCCOPY_ATTRS = 0x00000002,
        AADMS_DOCCOPY_NO_ATTRS = 0x00000004, /* attr copy denied      */
        AADMS_DOCCOPY_NO_SETITEM = 0x00000008, /* no set item for master*/
        AADMS_DOCCOPY_MOVE = 0x00000010,  /* move operation        */
        AADMS_DOCCOPY_NOFILE = 0x00000020, /* do not copy file      */
        AADMS_DOCCOPY_NO_HOOKS = 0x00010000, /* dont call pre/post hooks */
        AADMS_DOCCOPY_LOG_MOVE = 0x00020000, /* log audit trail move action */
        AADMS_DOCCOPY_INCLUDE_VERSIONS = 0x00040000, /* move versions if there are any */
        AADMS_DOCCOPY_COPY_ACCESS = 0x00080000, /* copy document access */
        AADMS_DOCCOPY_MOVE_NO_VERSIONS = 0x00100000, /* used with MOVE flag - versions will not be moved */
        AADMS_DOCCOPY_COPY_CONFBLOCKS = 0x00200000, /* copy Managed Workspace Profiles ConfBlock assignments */
        AADMS_DOCCOPY_COPYVERSIONSTR = 0x00400000 /* copy version string */
    }


    [DllImport("dmscli.dll", EntryPoint = "aaApi_DmsDataBufferGetBinaryProperty", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaApi_DmsDataBufferGetBinaryProperty
        (
        IntPtr hDataBuffer,    /* i  handle of the data buffer to free    */
        int lPropertyID,    /* i  Property identifier                  */
        int lIdxRow         /* i  Row Index                            */
        );



    [DllImport("dmscli.dll", EntryPoint = "aaApi_ThumbnailDataBufferSelectByDoc", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaApi_ThumbnailDataBufferSelectByDoc
        (
        ref Guid pDocGuid,               /* i  Document Id                           */
        string strThumbnailTimeStamp   /* i  Thumbnail timestamp. Can be NULL. If specified, */
                                       /* thumbnail will be returned only if newer exists. Else NOT_FOUND error will be returned */
        );


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_CopyDocument2
(
IntPtr hSrcDataSource,     /* i  Source datasource handle      */
int lSourceProjectNo,   /* i  Source project number         */
int lSourceDocumentId,  /* i  Source document number        */
IntPtr hTrgtDataSource,    /* i  Target datasource handle      */
int lTargetProjectNo,   /* i  Destination project number    */
ref int lpTargetDocumentId, /* io Target document number        */
string lpctstrWorkdir,     /* i  Working directory used in copy*/
string lpctstrFileName,    /* i  File name for the copy        */
string lpctstrName,        /* i  Name for the copy             */
string lpctstrDesc,        /* i  Description for the copy      */
DocumentCopyFlags ulFlags             /* i  Operation flags               */
);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_CopyDocument3
(
IntPtr hSrcDataSource,     /* i  Source datasource handle      */
int lSourceProjectNo,   /* i  Source project number         */
int lSourceDocumentId,  /* i  Source document number        */
IntPtr hTrgtDataSource,    /* i  Target datasource handle      */
int lTargetProjectNo,   /* i  Destination project number    */
ref int lpTargetDocumentId, /* io Target document number        */
string lpctstrWorkdir,     /* i  Working directory used in copy*/
string lpctstrFileName,    /* i  File name for the copy        */
string lpctstrName,        /* i  Name for the copy             */
string lpctstrDesc,        /* i  Description for the copy      */
DocumentCopyFlags ulFlags, /* i  Operation flags               */
[Out] out IntPtr ppVersionDocuments /* should be a structure that looks like { int count, 
                        [[int sourceProjId, int sourceDocId, int targetProjId, int targetDocId],[...]] */
);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_CopyDocument
(
   int lSourceProjectNo,    /* i  Source project identifier     */
   int lSourceDocumentId,   /* i  Source document identifier    */
   int lTargetProjectNo,    /* i  Destination project identifier*/
   ref int lplTargetDocumentId, /* io Target document identifier    */
   string lpctstrWorkdir,      /* i  Working directory used in copy*/
   string lpctstrFileName,     /* i  File name for the copy        */
   string lpctstrName,         /* i  Name for the copy             */
   string lpctstrDesc,         /* i  Description for the copy      */
   DocumentCopyFlags ulFlags              /* i  Flags fro the operation       */
);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectDocumentsByProjectId(int ProjectId);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_PurgeDocumentCopy
    (
       int lProjectNo,      /* i  Project number    */
       int lDocumentId,     /* i  Document number   */
       int lUserId          /* i  User number       */
    );



    public enum ManagerTypeProperty : int
    {
        //#define AADMS_MGRTYPE_USER                1
        //#define AADMS_MGRTYPE_GROUP               2
        //#define AADMS_MGRTYPE_USERLIST            3
        //#define AADMS_MGRTYPE_ALLUSERS            4   /* for access ctrl */
        None = 0,
        User = 1,
        Group = 2,
        UserList = 3,
        AllUsers = 4
    }


    public enum EnvAttributeGuiFlags : uint
    {
        ENVIRONMENTID = 0x00000001,
        TABLEID = 0x00000002,
        COLUMNID = 0x00000004,
        GUIID = 0x00000008,
        PAGENO = 0x00000010,
        TABORDER = 0x00000020,
        LABEL = 0x00000040,
        LABELFONT = 0x00000080,
        LABELFONT_HEIGHT = 0x00000100,
        LABEL_X = 0x00000200,
        LABEL_Y = 0x00000400,
        LABEL_WIDTH = 0x00000800,
        LABEL_HEIGHT = 0x00001000,
        EDIT_X = 0x00002000,
        EDIT_Y = 0x00004000,
        EDIT_WIDTH = 0x00008000,
        EDIT_HEIGHT = 0x00010000,
        PROMPT = 0x00020000,
        GUIFLAGS = 0x00040000,

        REQUIRED = (EnvAttributeGuiFlags.ENVIRONMENTID |
                                                   EnvAttributeGuiFlags.TABLEID |
                                                   EnvAttributeGuiFlags.COLUMNID |
                                                   EnvAttributeGuiFlags.GUIID)
    }

    //modified mds 20140127 -- added owner, [Flags] and changed to unit
    [Flags]
    public enum AccessMasks : uint
    {
        None = 0x00000000,
        Control = 0x00000001,
        Write = 0x00000002,
        Read = 0x00000004,
        FWrite = 0x00000008,
        FRead = 0x00000010,
        Create = 0x00000020,
        Delete = 0x00000040,
        Free = 0x00000080,
        ChangeWorkflowState = 0x00000100,
        //#define  AADMS_ACCESS_FREE      0x00000080      /**< Right to free document that is checked-out by other user. */
        //#define  AADMS_ACCESS_CWST      0x00000100      /**< Right to change workflow state. */
        Full = 0x0000FFFF,
        Owner = 0x00000200
    }

    //#define  AADMS_ACCESS_NONE      0x00000000      /**< Access is forbidden. */
    //#define  AADMS_ACCESS_CNTRL     0x00000001      /**< Right to change object permissions. */
    //#define  AADMS_ACCESS_WRITE     0x00000002      /**< Right to modify object attributes. */
    //#define  AADMS_ACCESS_READ      0x00000004      /**< The project is visible in Datasource Tree window, user can view the project's properties. */
    //#define  AADMS_ACCESS_FWRITE    0x00000008      /**< Right to modify file (used only for document access). */
    //#define  AADMS_ACCESS_FREAD     0x00000010      /**< Right to read file (used only for document access). */
    //#define  AADMS_ACCESS_CREATE    0x00000020      /**< Right to create. */
    //#define  AADMS_ACCESS_DELETE    0x00000040      /**< Right to delete. */
    //#define  AADMS_ACCESS_FREE      0x00000080      /**< Right to free document that is checked-out by other user. */
    //#define  AADMS_ACCESS_CWST      0x00000100      /**< Right to change workflow state. */
    //#define  AADMS_ACCESS_FULL      0x0000FFFF      /**< All rights. */




    public enum ObjectTypes : int
    {
        UserIgnoresAccessControl = 0,
        EnvProject = 1,
        Project = 2,
        EnvDoc = 3,
        Document = 4,
        Components = 5
    }

    public enum EnvAttrGuiProps : int
    {
        ENVIRONMENTID = 1,
        TABLEID = 2,
        COLUMNID = 3,
        GUIID = 4,
        PAGENO = 5,
        TABORDER = 6,
        LABELFONT_HEIGHT = 7,
        LABEL_X = 8,
        LABEL_Y = 9,
        LABEL_WIDTH = 10,
        LABEL_HEIGHT = 11,
        EDIT_X = 12,
        EDIT_Y = 13,
        EDIT_WIDTH = 14,
        EDIT_HEIGHT = 15,
        GUIFLAGS = 16,
        LABEL = 17,
        LABELFONT = 18,
        PROMPT = 19
    }



    [StructLayout(LayoutKind.Sequential)]
    public struct AADMSEATRGUIDEF
    {
        public uint ulFlags;            /* i Specifies valid fields       */
        public int lEnvironmentId;     /* i Environment ID               */
        public int lTableId;           /* i Attribute table ID           */
        public int lColumnId;          /* i Attribute column ID          */
        public int lGuiId;             /* i Interface ID                 */
        public int lPageNo;            /* i Property page ID             */
        public int lTabOrder;          /* i Tab order on the page        */
        public string lpctstrLabel;       /* i Text for the label           */
        public string lpctstrLabelFont;   /* i Label font name              */
        public int lLabelFontH;        /* i Label font height            */
        public int lLabelX;            /* i Label position (x-axis)      */
        public int lLabelY;            /* i Label position (y-axis)      */
        public int lLabelWidth;        /* i Label width                  */
        public int lLabelHeight;       /* i Label height                 */
        public int lEditX;             /* i Edit field position (x-axis) */
        public int lEditY;             /* i Edit field position (y-axis) */
        public int lEditWidth;         /* i Edit field width             */
        public int lEditHeight;        /* i Edit field height            */
        public string lpctstrPrompt;      /* i Message text                 */
        public int lGuiFlags;          /* i Interface parameter          */
    }


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_CreateEnvAttrGuiDef
    (
        [In] ref AADMSEATRGUIDEF lpAttrGuiDef   /* i Attribute Gui definition */
    );


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_CreateGui
    (
        ref int lGuiId,
        string sInterfaceName,
        string sInterfaceDescription
    );


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectStorage(int lStorageId);


    [DllImport("dmscli.dll", EntryPoint = "aaApi_GetStorageNumericProperty", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GetStorageNumericProperty(StorageProperty lPropertyId, int lIndex);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GUIDSelectProject(ref Guid guid);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_GUIDGetProjectNamePath(ref Guid guid, bool UseDesc,
        char tchSeparator, StringBuilder StringBuffer, int BufferSize);


    [DllImport("dmscli.dll", EntryPoint = "aaApi_GetEnvAttrGuiDefStringProperty", CharSet = CharSet.Unicode)]
    private static extern IntPtr unsafe_aaApi_GetEnvAttrGuiDefStringProperty(EnvAttrGuiProps PropertyId, int Index);


    public static string aaApi_GetEnvAttrGuiDefStringProperty(EnvAttrGuiProps PropertyId, int Index)
    {
        return Marshal.PtrToStringUni(unsafe_aaApi_GetEnvAttrGuiDefStringProperty(PropertyId, Index));
    }


    [DllImport("dmscli.dll", EntryPoint = "aaApi_GetEnvAttrGuiDefNumericProperty", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GetEnvAttrGuiDefNumericProperty(EnvAttrGuiProps PropertyId, int Index);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectAllGuis();


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectEnvAttrGuiDefs
    (
        int lEnvironmentId,   /* i  Environment number (-1 for all) */
        int lTableId,         /* i  Table number (-1 for all)       */
        int lColumnId,        /* i  Column number (-1 for all)      */
        int lGuiId,           /* i  Interface number (-1 for all)   */
        int lPageNo           /* i  Page number (-1 for all)        */
    );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectEnvCodeDefs(int environmentID, int tableID, int columnID, CodeDefinitionType type);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GetEnvCodeDefNumericProperty(DocumentCodeDefinitionProperty propertyID, int index);

    [DllImport("dmscli.dll", EntryPoint = "aaApi_GetEnvCodeDefStringProperty", CharSet = CharSet.Unicode)]
    private static extern IntPtr __aaApi_GetEnvCodeDefStringProperty(DocumentCodeDefinitionProperty property, int index);

    public static string aaApi_GetEnvCodeDefStringProperty(DocumentCodeDefinitionProperty propertyID, int index)
    {
        return Marshal.PtrToStringUni(__aaApi_GetEnvCodeDefStringProperty(propertyID, index));
    }

   [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectEnvAttrDefs(int environmentID, int tableID, int columnID);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GetEnvAttrDefNumericProperty(AttributeDefinitionProperty property, int index);

    [DllImport("dmscli.dll", EntryPoint = "aaApi_GetEnvAttrDefStringProperty", CharSet = CharSet.Unicode)]
    private static extern IntPtr unsafe_aaApi_GetEnvAttrDefStringProperty(AttributeDefinitionProperty property, int index);

    public static string aaApi_GetEnvAttrDefStringProperty(AttributeDefinitionProperty propertyID, int index)
    {
        return Marshal.PtrToStringUni(unsafe_aaApi_GetEnvAttrDefStringProperty(propertyID, index));
    }

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectEnvTriggerDefs(int iEnvironmentId,
          int iTableId,
          int iColumnId,
          int iTrigColumnId
        );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GetEnvTriggerDefNumericProperty(EnvTriggerProperty property, int index);

    [DllImport("dmscli.dll", EntryPoint = "aaApi_GetEnvTriggerDefStringProperty", CharSet = CharSet.Unicode)]
    private static extern IntPtr __aaApi_GetEnvTriggerDefStringProperty(EnvTriggerProperty property, int index);

    public static string aaApi_GetEnvTriggerDefStringProperty(EnvTriggerProperty propertyID, int index)
    {
        return Marshal.PtrToStringUni(__aaApi_GetEnvTriggerDefStringProperty(propertyID, index));
    }

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_CountEnvTriggerDefs();

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectAllEnvTriggerDefs();

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GetProjectId(int lIndex);


    public enum GuiProperty : int
    {
        Id = 1,
        Name = 2,
        Desc = 3
    }


    [DllImport("dmscli.dll", EntryPoint = "aaApi_GetGuiStringProperty", CharSet = CharSet.Unicode)]
    private static extern IntPtr unsafe_aaApi_GetGuiStringProperty(GuiProperty propertyID, int index);


    public static string aaApi_GetGuiStringProperty(GuiProperty propertyID, int index)
    {
        return Marshal.PtrToStringUni(unsafe_aaApi_GetGuiStringProperty(propertyID, index));
    }


    [DllImport("dmscli.dll", EntryPoint = "aaApi_SelectGui", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectGui(int iGuidId);


    [DllImport("dmscli.dll", EntryPoint = "aaApi_GetGuiId", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GetGuiId(int index);

    public enum TableProperty : int
    {
        Id = 1,
        Type = 2,
        Name = 3,
        Desc = 4
    }

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectAllTables();

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectTable(int iTableId);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GetTableId(int index);

    [DllImport("dmscli.dll", EntryPoint = "aaApi_GetTableStringProperty", CharSet = CharSet.Unicode)]
    private static extern IntPtr unsafe_aaApi_GetTableStringProperty(TableProperty propertyID, int index);

    public static string aaApi_GetTableStringProperty(TableProperty propertyID, int index)
    {
        return Marshal.PtrToStringUni(unsafe_aaApi_GetTableStringProperty(propertyID, index));
    }

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GetTableNumericProperty(TableProperty propertyID, int index);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectGroup(int lGroupId);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectUsersByGroup(int lGroupId);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectUsersByUserList(int lUserListId);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectUserListMembers(int lUsrLstId,  /* i  Access user list number   */
                ManagerTypeProperty lMemType,         /* i  Member type               */
                int lMemberId         /* i  Member id                 */);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_RemoveUserFromGroup(int iGroupId, int iUserId);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_AddUserToGroup(int iGroupId, int iUserId);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_DeleteUserListById(int iUserListId);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_DeleteGroupById(int iGroupId);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_DeleteGui(int iGuiId);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GetUserListMemNumericProperty
        (
           UserListMemberProperty lPropertyId,    /* i  Property id (USRLSTMEM_PROP_*)  */
           int lIndex          /* i  Index of selected user list     */
        );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GetUserListNumericProperty
        (
           UserListProperty lPropertyId,    /* i  Property id (USRLSTMEM_PROP_*)  */
           int lIndex          /* i  Index of selected user list     */
        );

    [DllImport("dmscli.dll", EntryPoint = "aaApi_SelectProjectsByType", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectProjectsByType(int iType);


    [DllImport("dmscli.dll", EntryPoint = "aaApi_SelectWorkspaceProfileByName", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaApi_SelectWorkspaceProfileByName(string sWorkspaceProfileName);


    [DllImport("dmscli.dll", EntryPoint = "aaApi_CreateGroup", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_CreateGroup(ref int iGroupNo,
           string lpctstrType,         /* i  Type name                     */
           string lpctstrSecProvider,  /* i  Security Provider             */
           string lpctstrName,         /* i  Group name                    */
           string lpctstrDesc          /* i  Group description             */
        );


    [DllImport("dmscli.dll", EntryPoint = "aaApi_GetUserListStringProperty", CharSet = CharSet.Unicode)]
    public static extern IntPtr unsafe_aaApi_GetUserListStringProperty(UserListProperty iUserListProp, int iIndex);


    public static string aaApi_GetUserListStringProperty(UserListProperty iUserListProp, int iIndex)
    {
        return Marshal.PtrToStringUni(unsafe_aaApi_GetUserListStringProperty(iUserListProp, iIndex));
    }


    [DllImport("dmscli.dll", EntryPoint = "aaApi_GetUserListId", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GetUserListId(int iIndex);


    [DllImport("dmscli.dll", EntryPoint = "aaApi_SelectUserLists", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectUserLists(int iListId, int iListType, int iOwner);

    [DllImport("dmscli.dll", EntryPoint = "aaApi_SelectUserListsByMembers", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectUserListsByMembers(
        UserListTypes lListType,
        int lOwnerId,
        MemberTypes lMemberType,
        int lMemberId
        );

    [DllImport("dmscli.dll", EntryPoint = "aaApi_GetGroupId", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GetGroupId(int iIndex);


    [DllImport("dmscli.dll", EntryPoint = "aaApi_SelectAllUserLists", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectAllUserLists();


    [DllImport("dmscli.dll", EntryPoint = "aaApi_SetDocumentFinalStatus", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_SetDocumentFinalStatus
(
int lProjectId,     /* i  Project number                     */
int lDocumentId,    /* i  Document number                    */
bool bAdd,            /* i  Add final status flag (0 - remove) */
string lpctstrComment  /* i  Operation comments for audit trail */
);

    [DllImport("dmscli.dll", EntryPoint = "aaApi_IsDocumentExported", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_IsDocumentExported
        (
        int lProjectId,         /* i  Project  id                    */
        int lDocumentId
        );


    [DllImport("dmscli.dll", EntryPoint = "aaApi_IsDocumentCheckedIn", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_IsDocumentCheckedIn
        (
        int lProjectId,         /* i  Project  id                    */
        int lDocumentId
        );


    [DllImport("dmscli.dll", EntryPoint = "aaApi_FreeDocument", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_FreeDocument
(
int lProjectNo,       /* i  Project number                */
int lDocumentId,      /* i  Document number               */
int lUserId           /* i  User number freeing document  */
);



    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GetFExtensionApplication(string sExtension);


    [DllImport("dmawin.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_LoginDlg(DataSourceType lDSType, StringBuilder lptstrDataSource, int lDSLength,
        string lpctstrUsername, string lpctstrPassword, string lpctstrSchema);


    public enum AuditTrailTypes : int
    {
        AADMSAT_TYPE_FIRST = 1,
        AADMSAT_TYPE_VAULT = (AADMSAT_TYPE_FIRST + 0),
        AADMSAT_TYPE_DOCUMENT = (AADMSAT_TYPE_FIRST + 1),
        AADMSAT_TYPE_DOCUMENT_SET = (AADMSAT_TYPE_FIRST + 2),
        AADMSAT_TYPE_WORKFLOW = (AADMSAT_TYPE_FIRST + 3),
        AADMSAT_TYPE_STATE = (AADMSAT_TYPE_FIRST + 4),
        AADMSAT_TYPE_USER = (AADMSAT_TYPE_FIRST + 5),
        AADMSAT_TYPE_GROUP = (AADMSAT_TYPE_FIRST + 6),
        AADMSAT_TYPE_USER_LIST = (AADMSAT_TYPE_FIRST + 7),
        AADMSAT_TYPE_LAST = AADMSAT_TYPE_USER_LIST,
        AADMSAT_TYPE_DS_REPORT = -55,
        AADMSAT_TYPE_USER_LILO = -56
    }

    // dww 2013-10-04 updated for SS4
    public enum AuditTrailActions : int
    {
        AADMSAT_ACT_DOC_FIRST = 1000,
        AADMSAT_ACT_DOC_UNKNOWN = (AADMSAT_ACT_DOC_FIRST),
        AADMSAT_ACT_DOC_CREATE = (AADMSAT_ACT_DOC_FIRST + 1),
        AADMSAT_ACT_DOC_MODIFY = (AADMSAT_ACT_DOC_FIRST + 2),
        AADMSAT_ACT_DOC_ATTR = (AADMSAT_ACT_DOC_FIRST + 3),
        AADMSAT_ACT_DOC_FILE_ADD = (AADMSAT_ACT_DOC_FIRST + 4),
        AADMSAT_ACT_DOC_FILE_REM = (AADMSAT_ACT_DOC_FIRST + 5),
        AADMSAT_ACT_DOC_FILE_REP = (AADMSAT_ACT_DOC_FIRST + 6),
        AADMSAT_ACT_DOC_CIN = (AADMSAT_ACT_DOC_FIRST + 7),
        AADMSAT_ACT_DOC_VIEW = (AADMSAT_ACT_DOC_FIRST + 8),
        AADMSAT_ACT_DOC_CHOUT = (AADMSAT_ACT_DOC_FIRST + 9),
        AADMSAT_ACT_DOC_CPOUT = (AADMSAT_ACT_DOC_FIRST + 10),
        AADMSAT_ACT_DOC_GOUT = (AADMSAT_ACT_DOC_FIRST + 11),
        AADMSAT_ACT_DOC_STATE = (AADMSAT_ACT_DOC_FIRST + 12),
        AADMSAT_ACT_DOC_FINAL_S = (AADMSAT_ACT_DOC_FIRST + 13),
        AADMSAT_ACT_DOC_FINAL_R = (AADMSAT_ACT_DOC_FIRST + 14),
        AADMSAT_ACT_DOC_VERSION = (AADMSAT_ACT_DOC_FIRST + 15),
        AADMSAT_ACT_DOC_MOVE = (AADMSAT_ACT_DOC_FIRST + 16),
        AADMSAT_ACT_DOC_COPY = (AADMSAT_ACT_DOC_FIRST + 17),
        AADMSAT_ACT_DOC_SECUR = (AADMSAT_ACT_DOC_FIRST + 18),
        AADMSAT_ACT_DOC_REDLINE = (AADMSAT_ACT_DOC_FIRST + 19),
        AADMSAT_ACT_DOC_DELETE = (AADMSAT_ACT_DOC_FIRST + 20),
        AADMSAT_ACT_DOC_EXPORT = (AADMSAT_ACT_DOC_FIRST + 21),
        AADMSAT_ACT_DOC_FREE = (AADMSAT_ACT_DOC_FIRST + 22),
        AADMSAT_ACT_DOC_EXTRACT = (AADMSAT_ACT_DOC_FIRST + 23),
        AADMSAT_ACT_DOC_DISTRIBUTE = (AADMSAT_ACT_DOC_FIRST + 24),
        AADMSAT_ACT_DOC_SEND_TO = (AADMSAT_ACT_DOC_FIRST + 25),
        AADMSAT_ACT_DOC_COMMENT = (AADMSAT_ACT_DOC_FIRST + 26),
        AADMSAT_ACT_DOC_IMPORT = (AADMSAT_ACT_DOC_FIRST + 27),
        AADMSAT_ACT_DOC_ACL_ASSIGN = (AADMSAT_ACT_DOC_FIRST + 28),
        AADMSAT_ACT_DOC_ACL_MODIFY = (AADMSAT_ACT_DOC_FIRST + 29),
        AADMSAT_ACT_DOC_ACL_REMOVE = (AADMSAT_ACT_DOC_FIRST + 30),
        AADMSAT_ACT_DOC_REVIT = (AADMSAT_ACT_DOC_FIRST + 31),
        AADMSAT_ACT_DOC_PACK = (AADMSAT_ACT_DOC_FIRST + 32),
        AADMSAT_ACT_DOC_UNPACK = (AADMSAT_ACT_DOC_FIRST + 33),
        AADMSAT_ACT_DOC_LAST = AADMSAT_ACT_DOC_UNPACK,
        AADMSAT_ACT_SET_FIRST = 2001,
        AADMSAT_ACT_SET_CREATE = (AADMSAT_ACT_SET_FIRST + 0),
        AADMSAT_ACT_SET_ADD = (AADMSAT_ACT_SET_FIRST + 1),
        AADMSAT_ACT_SET_REMOVE = (AADMSAT_ACT_SET_FIRST + 2),
        AADMSAT_ACT_SET_LAST = AADMSAT_ACT_SET_REMOVE,
        AADMSAT_ACT_NONE = -1,
        AADMSAT_ACT_DEFAULT = 0,
        AADMSAT_ACT_VLT_FIRST = 1,
        AADMSAT_ACT_VLT_CREATE = (AADMSAT_ACT_VLT_FIRST + 0),
        AADMSAT_ACT_VLT_MODIFY = (AADMSAT_ACT_VLT_FIRST + 1),
        AADMSAT_ACT_VLT_WFLOW = (AADMSAT_ACT_VLT_FIRST + 2),
        AADMSAT_ACT_VLT_DELETE = (AADMSAT_ACT_VLT_FIRST + 3),
        AADMSAT_ACT_VLT_STATE = (AADMSAT_ACT_VLT_FIRST + 4),
        AADMSAT_ACT_VLT_ACL_ASSIGN = (AADMSAT_ACT_VLT_FIRST + 5),
        AADMSAT_ACT_VLT_ACL_MODIFY = (AADMSAT_ACT_VLT_FIRST + 6),
        AADMSAT_ACT_VLT_ACL_REMOVE = (AADMSAT_ACT_VLT_FIRST + 7),
        AADMSAT_ACT_VLT_LAST = AADMSAT_ACT_VLT_ACL_REMOVE,
        AADMSAT_ACT_USER_FIRST = 3001,
        AADMSAT_ACT_USER_LOGIN = (AADMSAT_ACT_USER_FIRST + 0),
        AADMSAT_ACT_USER_LOGOUT = (AADMSAT_ACT_USER_FIRST + 1),
        AADMSAT_ACT_USER_CREATE = (AADMSAT_ACT_USER_FIRST + 2),
        AADMSAT_ACT_USER_LAST = AADMSAT_ACT_USER_CREATE,
        AADMSAT_ACT_GROUP_FIRST = 4001,
        AADMSAT_ACT_GROUP_CREATE = (AADMSAT_ACT_GROUP_FIRST + 0),
        AADMSAT_ACT_GROUP_MODIFY = (AADMSAT_ACT_GROUP_FIRST + 1),
        AADMSAT_ACT_GROUP_ADD = (AADMSAT_ACT_GROUP_FIRST + 2),
        AADMSAT_ACT_GROUP_REMOVE = (AADMSAT_ACT_GROUP_FIRST + 3),
        AADMSAT_ACT_GROUP_LAST = AADMSAT_ACT_GROUP_REMOVE,
        AADMSAT_ACT_USER_LIST_FIRST = 5001,
        AADMSAT_ACT_USER_LIST_CREATE = (AADMSAT_ACT_USER_LIST_FIRST + 0),
        AADMSAT_ACT_USER_LIST_MODIFY = (AADMSAT_ACT_USER_LIST_FIRST + 1),
        AADMSAT_ACT_USER_LIST_ADD = (AADMSAT_ACT_USER_LIST_FIRST + 2),
        AADMSAT_ACT_USER_LIST_REMOVE = (AADMSAT_ACT_USER_LIST_FIRST + 3),
        AADMSAT_ACT_USER_LIST_LAST = AADMSAT_ACT_USER_LIST_REMOVE,
        AADMSAT_ACT_VLT_CUSTOM_FIRST = 100,
        AADMSAT_ACT_VLT_CUSTOM_LAST = 999,
        AADMSAT_ACT_DOC_CUSTOM_FIRST = 1100,
        AADMSAT_ACT_DOC_CUSTOM_LAST = 1999,
        AADMSAT_ACT_SET_CUSTOM_FIRST = 2100,
        AADMSAT_ACT_SET_CUSTOM_LAST = 2999,
        AADMSAT_ACT_GEN_CUSTOM_FIRST = 10000,
        AADMSAT_ACT_GEN_CUSTOM_LAST = 19999,
        AADMSAT_ACT_DOC_ATTR_CREATE = 1,
        AADMSAT_ACT_DOC_ATTR_MODIFY = 2,
        AADMSAT_ACT_DOC_ATTR_DELETE = 3,
        // AADMSAT_ACT_USER_LOGIN = -669,
        // AADMSAT_ACT_USER_LOGOUT = -670,
        AADMSAT_ACT_USER_CONNCOUNT = -671
    }

    [DllImport("dmscli.dll", EntryPoint = "aaApi_GetAuditTrailActionTypeName", CharSet = CharSet.Unicode)]
    private static extern IntPtr unsafe_aaApi_GetAuditTrailActionTypeName(AuditTrailActions iActionTypeId, ref int pObjectTypeId);

    public static string aaApi_GetAuditTrailActionTypeName(AuditTrailActions iActionTypeId, ref int pObjectTypeId)
    {
        return Marshal.PtrToStringUni(unsafe_aaApi_GetAuditTrailActionTypeName(iActionTypeId, ref pObjectTypeId));
    }

    // DAB 2016-02-08
    [DllImport("dmscli.dll", EntryPoint = "aaApi_SetAuditLoggingSetting", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_SetAuditLoggingSetting(AuditTrailTypes auditTrailRecordType, AuditTrailActions auditTrailActionType, bool bEnabled);

    // DAB 2016-02-08
    [DllImport("dmscli.dll", EntryPoint = "aaApi_SetAuditLoggingSetting", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_SetAuditLoggingSetting(int iAuditTrailRecordType, int iAuditTrailActionType, bool bEnabled);

    // DAB 2016-02-08
    [DllImport("dmscli.dll", EntryPoint = "aaApi_GetAuditLoggingSetting", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_GetAuditLoggingSetting(AuditTrailTypes auditTrailRecordType, AuditTrailActions auditTrailActionType);

    // DAB 2016-02-08
    [DllImport("dmscli.dll", EntryPoint = "aaApi_GetAuditLoggingSetting", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_GetAuditLoggingSetting(int iAuditTrailRecordType, int iAuditTrailActionType);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GUIDSelectProjectsFromBranch(ref Guid guid,
        string codeString,
        string projectName,
        string projectDescription,
        string projectVersion
        );


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_GetProjectGUIDsByIds([In]int lCount, [In]ref int pProjectIds,
        [Out] Guid[] docGuids);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_DeleteDocument(DocumentDeleteMasks uiFlags, int iProjectId, int iDocumentId);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_GUIDDeleteDocument(DocumentDeleteMasks uiFlags, ref Guid pDocGuid);

    public static int GetFolderNoFromPath(string pwPath, int targetVaultID)
    {
        char[] delimiters = { '\\', '/' };
        string[] pathSteps = pwPath.Split(delimiters, StringSplitOptions.RemoveEmptyEntries);

        int parentVaultID = targetVaultID;
        int childVaultID = 0;  // 0 is an invalid vault ID

        for (int i = 0; i < pathSteps.Length; i++)
        {
            string vaultName = pathSteps[i].Trim(" ".ToCharArray());

            if (string.IsNullOrEmpty(vaultName))
                continue;

            // search for the vault to see if it already exists
            int numChildren = -1;
            if (-1 == parentVaultID)
            {
                numChildren = aaApi_SelectTopLevelProjects();
            }
            else
            {
                numChildren = aaApi_SelectChildProjects(parentVaultID);
            }
            if (numChildren == -1)
            {
                // string message = "Error selecting child folders";
                return 0;
                // throw new ApplicationException(message);
            }

            bool childFound = false;
            for (int j = 0; j < numChildren; j++)
            {
                string childVaultName = aaApi_GetProjectStringProperty(ProjectProperty.Name, j);

                if (childVaultName == vaultName)
                {
                    childFound = true;
                    childVaultID = aaApi_GetProjectNumericProperty(ProjectProperty.ID, j);
                    break;
                }
            }

            // if the child vault was not found, create it
            if (childFound == false)
            {
                return 0;
            }

            parentVaultID = childVaultID;
        } // for (int i = 0; i < pathSteps.Length; i++)

        return childVaultID;
    }

    [DllImport("dmawin.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectProjectDlg(IntPtr hWndParent,        /* i  Owner window handle  */
       string lpctstrTitle,      /* i  dialog title         */
       int lProjectId         /* i  Project identifier   */
    );

    [DllImport("dmawin.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectProjectDlg2(IntPtr hWndParent,
        string lpctstrTitle,
        string lpctstrRootText,
        uint ulFlags,
        IntPtr hIcon,
        ref int lplProjectId
    );

    [DllImport("dmawin.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SaveDocumentDlg(IntPtr hWndParent,
  string lpctstrTitle,
  uint ulDlgType,
  ref int lplProjectId,
  ref int lplDocumentId,
  StringBuilder lptstrFileName,
  int lLength
 );

    [DllImport("dmawin.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaApi_FindDocumentDlg();

    [DllImport("dmawin.dll", CharSet = CharSet.Unicode)]
    public static extern long aaApi_SelectSavedQueryDlg
    (
       IntPtr hWndParent,        /* i  Owner window handle      */
       string lpctstrTitle,       /* i  Dialog window caption    */
       ref int plSQueryNo        /* o  Select saved query ID    */
    );

    [Flags]
    public enum AttributeCopyFlags : uint
    {
        None = 0x00000000,
        IgnoreEnvironmentCopyOptions = 0x00000001,
        CopyCodeFields = 0x00000002,
        SkipCopyInSameEnvironment = 0x00000004
    }

    ///* Ignore attribute copy options defined by environment */
    //#define AADMS_ATTRCOPYF_IGNORE_ENVCOPYOPTS                0x00000001
    ///* Copy code fields (by default - code fields are not copied) */
    //#define AADMS_ATTRCOPYF_COPY_CODEFIELDS                   0x00000002
    ///* Skip attribute copying in the same environment */
    //#define AADMS_ATTRCOPYF_SKIP_COPY_IN_SAME_ENVIRONMENT     0x00000004



    //[StructLayout(LayoutKind.Sequential)]
    //public unsafe struct VaultDescriptorU
    //{
    //    public uint Flags;      /* specifies valid fields  AADMSPROJF_XXX */
    //    public int VaultID;
    //    public int EnvironmentID;
    //    public int ParentID;
    //    public int StorageID;
    //    public int ManagerID;
    //    public int TypeID;
    //    public int WorkflowID;
    //    public unsafe void* Name;
    //    public unsafe void* Description;
    //    public int ManagerType;
    //    public int WorkspaceProfileId;
    //    public Guid GuidVault;
    //    public int ClassId;
    //    public int InstanceId;
    //    public uint ProjectFlagsMask; /* specifies valid bits in projFlags */
    //    public uint ProjectFlags;   /* project flags AADMS_PROJF_XXX     */
    //}



    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SqlSelectDataBufGetColumnCount(IntPtr hDataBuffer);


    [DllImport("dmscli.dll", EntryPoint = "aaApi_SqlSelectDataBufGetValue", CharSet = CharSet.Unicode)]
    private static extern IntPtr unsafe_aaApi_SqlSelectDataBufGetValue(IntPtr hDataBuffer, int iRowIndex, int iColumnIndex);


    public static string aaApi_SqlSelectDataBufGetValue(IntPtr hDataBuffer, int iRowIndex, int iColumnIndex)
    {
        return Marshal.PtrToStringUni(unsafe_aaApi_SqlSelectDataBufGetValue(hDataBuffer, iRowIndex, iColumnIndex));
    }


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaApi_SqlSelectDataBuffer(string sqlStatement, IntPtr columnBind);

    public enum SqlSelectProperties : int
    {
        SQLSELECT_COLUMN_TYPE = 1,
        SQLSELECT_COLUMN_NATIVE_TYPE = 2,
        SQLSELECT_COLUMN_LENGTH = 3,
        SQLSELECT_COLUMN_NAME = 4
    }

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SqlSelectDataBufGetNumericProperty(IntPtr hDataBuffer, SqlSelectProperties lPropertyId, int lIdxCol);

    [DllImport("dmscli.dll", EntryPoint = "aaApi_SqlSelectDataBufGetStringProperty", CharSet = CharSet.Unicode)]
    private static extern IntPtr unsafe_aaApi_SqlSelectDataBufGetStringProperty(IntPtr hDataBuffer, SqlSelectProperties lPropertyId, int lIdxCol);


    public static string aaApi_SqlSelectDataBufGetStringProperty(IntPtr hDataBuffer, SqlSelectProperties lPropertyId, int lIdxCol)
    {
        return Marshal.PtrToStringUni(unsafe_aaApi_SqlSelectDataBufGetStringProperty(hDataBuffer, lPropertyId, lIdxCol));
    }


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SqlSelectGetNumericProperty(SqlSelectProperties lPropertyId, int lIdxCol);

    [DllImport("dmscli.dll", EntryPoint = "aaApi_SqlSelectGetStringProperty", CharSet = CharSet.Unicode)]
    private static extern IntPtr unsafe_aaApi_SqlSelectGetStringProperty(SqlSelectProperties lPropertyId, int lIdxCol);


    public static string aaApi_SqlSelectGetStringProperty(SqlSelectProperties lPropertyId, int lIdxCol)
    {
        return Marshal.PtrToStringUni(unsafe_aaApi_SqlSelectGetStringProperty(lPropertyId, lIdxCol));
    }

    [DllImport("dmscli.dll", EntryPoint = "aaApi_SqlSelectDataBufGetStringProperty", CharSet = CharSet.Unicode)]
    private static extern IntPtr unsafe_aaApi_SqlSelectDataBufGetStringProperty(IntPtr hDataBuffer, int lPropertyId, int lIdxCol);


    public static string aaApi_SqlSelectDataBufGetStringProperty(IntPtr hDataBuffer, int lPropertyId, int lIdxCol)
    {
        return Marshal.PtrToStringUni(unsafe_aaApi_DmsDataBufferGetStringProperty(hDataBuffer, lPropertyId, lIdxCol));
    }

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern void aaApi_DmsDataBufferFree(IntPtr hDataBuffer);

    public static string GetPWStringSetting(int iSetting)
    {
        string sSQL;

        sSQL = string.Format(
            "select o_textval from dms_gcfg where o_compguid = '4493b532-0cc3-45ee-be8c-1de7b9a7bad4' and o_paramno = {0}",
            iSetting);

        IntPtr hSqlBuf = aaApi_SqlSelectDataBuffer(sSQL, IntPtr.Zero);

        string sSetting = aaApi_SqlSelectDataBufGetValue(hSqlBuf, 0, 0);

        aaApi_DmsDataBufferFree(hSqlBuf);

        return sSetting;
    }


    public static int GetPWNumericSetting(int iSetting)
    {
        string sSQL;

        sSQL = string.Format(
            "select o_intval from dms_gcfg where o_compguid = '4493b532-0cc3-45ee-be8c-1de7b9a7bad4' and o_paramno = {0}",
            iSetting);

        IntPtr hSqlBuf = aaApi_SqlSelectDataBuffer(sSQL, IntPtr.Zero);

        string sSetting = aaApi_SqlSelectDataBufGetValue(hSqlBuf, 0, 0);

        aaApi_DmsDataBufferFree(hSqlBuf);

        int iSettingVal = 0;

        int.TryParse(sSetting, out iSettingVal);

        return iSettingVal;
    }

    public static bool SetPWStringSetting(int iSetting, string sValue)
    {
        aaApi_ExecuteSqlStatement("create view v_dms_gcfg as select * from dms_gcfg");

        string sSql = string.Format("delete from v_dms_gcfg where o_compguid = '4493b532-0cc3-45ee-be8c-1de7b9a7bad4' and o_paramno = {0}", iSetting);

        aaApi_ExecuteSqlStatement(sSql);

        sSql =
            string.Format("insert into v_dms_gcfg(o_compguid, o_paramno, o_intval, o_textval) values ('4493b532-0cc3-45ee-be8c-1de7b9a7bad4', {0}, 0, '{1}')",
            iSetting,
            sValue);

        return aaApi_ExecuteSqlStatement(sSql);
    }

    public static bool SetPWNumericSetting(int iSetting, int iValue)
    {
        aaApi_ExecuteSqlStatement("create view v_dms_gcfg as select * from dms_gcfg");

        string sSql = string.Format("delete from v_dms_gcfg where o_compguid = '4493b532-0cc3-45ee-be8c-1de7b9a7bad4' and o_paramno = {0}", iSetting);

        aaApi_ExecuteSqlStatement(sSql);

        sSql =
            string.Format("insert into v_dms_gcfg(o_compguid, o_paramno, o_intval) values ('4493b532-0cc3-45ee-be8c-1de7b9a7bad4', {0}, {1})",
            iSetting,
            iValue);

        return aaApi_ExecuteSqlStatement(sSql);
    }

    // Don't know why this is in here twice.  Removing 2011-01-12  DAB
    //[DllImport("dmscli.dll", EntryPoint = "aaApi_ModifyProject2", CharSet = CharSet.Unicode)]
    //public static extern bool aaApi_ModifyProject2U(ref VaultDescriptorU vaultDescriptorU);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_ModifyDocument
(
int lProjectId,          /* i  Project number                */
int lDocumentId,         /* i  Document number               */
int lFileType,           /* i  Modified file type            */
int lItemType,           /* i  Modified type                 */
int lApplicationId,      /* i  Modified application type     */
int lDepartmentId,       /* i  Modified department type      */
int lWorkspaceProfileId, /* i  Modified workspace profile id */
string lpctstrFileName,     /* i  Modified file name            */
string lpctstrName,         /* i  Modified document name        */
string lpctstrDesc          /* i  Modified document description */
);


    [Flags]
    public enum ModifyDocumentItemFlags : uint
    {
        Default = 0x00000000,        
        FileType = 0x00000001,     
        ItemType = 0x00000002,     
        Application = 0x00000004,
        Department = 0x00000008,
        FileName = 0x00000010,
        Description = 0x00000020,
        WorkspaceProfile = 0x00000400,
        ItemFlags = 0x00000800,
        MimeType = 0x00001000,
        FileRevision = 0x00002000,
        Name = 0x00004000,
        Version = 0x00008000,
        DocumentGuid = 0x80000000
    }



    // dww 2013-12-18 for COT project
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct DocumentParameters
    {
        public uint ulMask;
        public int ProjectId;
        public int DocumentId;
        public int FileType;
        public int ItemType;
        public int ApplicationId;
        public int DepartmentId;
        public string FileName;
        public string Name;
        public string Description;
        public int WorkspaceProfileId;
        public Guid GuidDoc;
        public ModifyDocumentItemFlags ItemFlagMask;
        public uint ItemFlags;
        public string MimeType;
        public string Revision;
        public string Version;
    }

    // dww 2013-12-18 for COT project
    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_ModifyDocument2(ref DocumentParameters documentParameters);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_ExportDocument(int lProjectNo, int lDocumentId,
        string lpctstrWorkdir, StringBuilder lptstrFileName, int lBufferSize);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_CreateFlatSet
(
int lProjectId,       /* i  Project number for set           */
ref int lplDocumentId,    /* o  Document number of set item      */
ref int lplSetId,         /* o  Set number of created set        */
int lDepartmentId,    /* i  Department number for set        */
string lpctstrName,      /* i  Set name                         */
string lpctstrDesc,      /* i  Set description                  */
int lChildProjectId,  /* i  Child project number of set      */
int lChildDocumentId, /* i  Child document number of set     */
bool bCheckOut         /* i  Transfer type for the set member */
);

    ///* Ignore attribute copy options defined by environment */
    //#define AADMS_ATTRCOPYF_IGNORE_ENVCOPYOPTS                0x00000001
    ///* Copy code fields (by default - code fields are not copied) */
    //#define AADMS_ATTRCOPYF_COPY_CODEFIELDS                   0x00000002
    ///* Skip attribute copying in the same environment */
    //#define AADMS_ATTRCOPYF_SKIP_COPY_IN_SAME_ENVIRONMENT     0x00000004


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_AddFlatSetMember
(
int lSetId,       /* i  Set number                             */
int lProjectId,   /* i  Member project number                  */
int lDocumentId,  /* i  Member document number                 */
bool bCheckOut     /* i  Check out when set is checked out flag */
);

    ///* Ignore attribute copy options defined by environment */
    //#define AADMS_ATTRCOPYF_IGNORE_ENVCOPYOPTS                0x00000001
    ///* Copy code fields (by default - code fields are not copied) */
    //#define AADMS_ATTRCOPYF_COPY_CODEFIELDS                   0x00000002
    ///* Skip attribute copying in the same environment */
    //#define AADMS_ATTRCOPYF_SKIP_COPY_IN_SAME_ENVIRONMENT     0x00000004


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectProjectsByEnvironment(int iEnvId);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_MoveDocument
(
int lSourceProjectNo,    /* i  Source project number         */
int lSourceDocumentId,   /* i  Source document number        */
int lTargetProjectNo,    /* i  Destination project number    */
ref int lplTargetDocumentId, /* io Target document number        */
string lpctstrWorkdir,      /* i  Working directory used in move*/
string lpctstrFileName,     /* i  File name for the copy        */
string lpctstrName,         /* i  Name for the copy             */
string lpctstrDesc,         /* i  Description for the copy      */
DocumentCopyFlags ulFlags              /* i  Flags for the operation       */
);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_MoveDocumentBindings
(
   ref AaDocItem pSourceDocs,     /* i  Source document IDs        */
   ref AaDocItem pTargetDocs,     /* i  Target document IDs        */
   int lDocumentCount,  /* i  Source document number     */
   uint ulFalgs          /* i  Operation flags (reserved) */
);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_DeleteDocumentFile(int vaultID, int documentID);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectDocumentVersions(int ProjectId, int lDocumentId);

    public static bool CopyOutDocument(string sDocumentGuid, bool bUseWorkingDirectory, FetchDocumentFlags fetchFlags, ref string sFileName)
    {
        int iProjectId = 0, iDocumentId = 0;

        if (GetIdsFromGuidString(sDocumentGuid, ref iProjectId, ref iDocumentId))
        {
            return CopyOutDocument(iProjectId, iDocumentId, bUseWorkingDirectory, fetchFlags, ref sFileName);
        }

        return false;
    }

    public static bool CopyOutDocument(Guid docGuid, bool bUseWorkingDirectory, FetchDocumentFlags fetchFlags, ref string sFileName)
    {
        PWWrapper.AaDocItem docItem = new PWWrapper.AaDocItem();
        Guid[] guids = new Guid[1];

        guids[0] = docGuid;

        if (PWWrapper.aaApi_GetDocumentIdsByGUIDs(1, guids, ref docItem))
        {
            return CopyOutDocument(docItem.lProjectId, docItem.lDocumentId, bUseWorkingDirectory, fetchFlags, ref sFileName);
        }

        return false;
    }

    private static object _lockObject = new object();

    public static bool CopyOutDocument(int iProjectId, int iDocumentId, bool bUseWorkingDirectory, FetchDocumentFlags fetchFlags, ref string sFileName)
    {
        string sWorkingDir = string.Empty;

        if (!bUseWorkingDirectory)
        {
            sWorkingDir = System.IO.Path.Combine(System.IO.Path.GetTempPath(), Guid.NewGuid().ToString());

            try
            {
                System.IO.Directory.CreateDirectory(sWorkingDir);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(string.Format("Error: {0}\n{1}", ex.Message, ex.StackTrace));
                return false;
            }
        }

        lock (_lockObject)
        {
            StringBuilder sbOutFile = new StringBuilder(1024);

            if (PWWrapper.aaApi_FetchDocumentFromServer(fetchFlags, iProjectId, iDocumentId,
                bUseWorkingDirectory ? null : sWorkingDir, sbOutFile, sbOutFile.Capacity))
            {
                if (System.IO.File.Exists(sbOutFile.ToString()))
                {
                    sFileName = sbOutFile.ToString();

                    return true;
                }
            }
            else
            {
                // fix for bug in native api which creates a zero byte file in the working directory if specified as "null"!!!
                return false;
            }
        }

        if (bUseWorkingDirectory)
        {
            StringBuilder sbCheckFile = new StringBuilder(1024);
            if (aaApi_GetDocumentFileName(iProjectId, iDocumentId, sbCheckFile, sbCheckFile.Capacity))
            {
                if (System.IO.File.Exists(sbCheckFile.ToString()))
                {
                    sFileName = sbCheckFile.ToString();
                    return true;
                }
            }
        }

        return false;
    }

    public static int GetAttrIDFromName
    (
        string wcAttrNameP	// i - attribute name, wide chars
    )
    {
        // find a pointer to an attribute using its name
        IntPtr pAttr = PWWrapper.aaOApi_FindAttributePtrByName(wcAttrNameP);

        if (pAttr == IntPtr.Zero || pAttr == null)
            return 0;
        // return the id of the attribute 
        return PWWrapper.aaOApi_GetAttributeId(pAttr);
    }


    public static bool SetRichProjectProperty(int iProjectId, string sAttributeName, string sAttributeValue)
    {
        bool bRetVal = false;

        IntPtr hProjBuf = PWWrapper.aaApi_SelectRichProjectOfFolder(iProjectId);

        if (hProjBuf != IntPtr.Zero)
        {
            if (1 == PWWrapper.aaApi_DmsDataBufferGetCount(hProjBuf))
            {
                int iClassId = PWWrapper.aaApi_DmsDataBufferGetNumericProperty(hProjBuf,
                    (int)PWWrapper.ProjectProperty.ComponentClassId, 0);
                int iInstanceId = PWWrapper.aaApi_DmsDataBufferGetNumericProperty(hProjBuf,
                    (int)PWWrapper.ProjectProperty.ComponentInstanceId, 0);

                IntPtr instP = PWWrapper.aaOApi_LoadInstanceByIds(iClassId, iInstanceId, 0);

                SortedList<string, int> slProperties = GetClassPropertyIdsInList(iClassId);

                if (instP != IntPtr.Zero)
                {
                    int lAttribId = 0;

                    if (!slProperties.ContainsKey(sAttributeName))
                    {
                        string sAttrName = sAttributeName.ToLower().Replace(" ", "_");

                        if (!slProperties.ContainsKey(sAttrName))
                        {
                            sAttrName = (string.Format("PROJECT_{0}", sAttributeName.ToLower())).Replace(" ", "_");

                            if (!slProperties.ContainsKey(sAttrName))
                            {
                            }
                            else
                            {
                                lAttribId = slProperties[sAttrName];
                            }
                        }
                        else
                        {
                            lAttribId = slProperties[sAttrName];
                        }
                    }
                    else
                    {
                        lAttribId = slProperties[sAttributeName.ToLower()];
                    }

                    // int lAttribId = PWWrapper.GetAttrIDFromName(sAttributeName);

                    //if (lAttribId == 0)
                    //{
                    //    string sAttrName = (string.Format("PROJECT_{0}", sAttributeName)).Replace(" ", "_");

                    //    lAttribId = PWWrapper.GetAttrIDFromName(sAttrName);
                    //}

                    if (lAttribId > 0)
                    {
                        if (!PWWrapper.aaOApi_SetInstanceAttrStrValueExt(instP,
                            sAttributeValue, lAttribId, 0, true))
                        {
                            System.Diagnostics.Debug.WriteLine(string.Format("Error setting {0} to {1}",
                                sAttributeName, sAttributeValue));
                        }

                        if (!PWWrapper.aaOApi_SaveInstance(instP))
                        {
                            System.Diagnostics.Debug.WriteLine("Error saving instance");
                        }
                        else
                        {
                            bRetVal = true;
                        }
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine(string.Format("Attribute {0} not found",
                            sAttributeName));
                    }
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine(string.Format("Error loading instance for {0},{1}",
                            iClassId, iInstanceId));
                }
            }

            PWWrapper.aaApi_DmsDataBufferFree(hProjBuf);
        }

        return bRetVal;
    }

    public static bool SetRichProjectProperties(int iProjectId, SortedList<string, string> slPropertyNamesPropertyValues)
    {
        bool bRetVal = false;

        IntPtr hProjBuf = PWWrapper.aaApi_SelectRichProjectOfFolder(iProjectId);

        if (hProjBuf != IntPtr.Zero)
        {
            if (1 == PWWrapper.aaApi_DmsDataBufferGetCount(hProjBuf))
            {
                int iClassId = PWWrapper.aaApi_DmsDataBufferGetNumericProperty(hProjBuf,
                    (int)PWWrapper.ProjectProperty.ComponentClassId, 0);
                int iInstanceId = PWWrapper.aaApi_DmsDataBufferGetNumericProperty(hProjBuf,
                    (int)PWWrapper.ProjectProperty.ComponentInstanceId, 0);

                IntPtr instP = PWWrapper.aaOApi_LoadInstanceByIds(iClassId, iInstanceId, 0);

                if (instP != IntPtr.Zero)
                {
                    SortedList<string, int> slProperties = GetProjectPropertyIdsInList(iProjectId);

                    foreach (string sAttributeName in slPropertyNamesPropertyValues.Keys)
                    {
                        if (string.IsNullOrEmpty(sAttributeName))
                            continue;

                        string sAttributeValue = slPropertyNamesPropertyValues[sAttributeName];

                        int lAttribId = 0;

                        if (!slProperties.ContainsKey(sAttributeName))
                        {
                            string sAttrName = sAttributeName.ToLower().Replace(" ", "_");

                            if (!slProperties.ContainsKey(sAttrName))
                            {
                                sAttrName = (string.Format("PROJECT_{0}", sAttributeName.ToLower())).Replace(" ", "_");

                                if (!slProperties.ContainsKey(sAttrName))
                                {
                                }
                                else
                                {
                                    lAttribId = slProperties[sAttrName];
                                }
                            }
                            else
                            {
                                lAttribId = slProperties[sAttrName];
                            }
                        }
                        else
                        {
                            lAttribId = slProperties[sAttributeName.ToLower()];
                        }

                        //int lAttribId = PWWrapper.GetAttrIDFromName(sAttributeName);

                        //if (lAttribId == 0)
                        //{
                        //    string sAttrName = (string.Format("PROJECT_{0}", sAttributeName)).Replace(" ", "_");

                        //    lAttribId = PWWrapper.GetAttrIDFromName(sAttrName);
                        //}

                        if (lAttribId > 0)
                        {
                            if (!PWWrapper.aaOApi_SetInstanceAttrStrValueExt(instP,
                                sAttributeValue, lAttribId, 0, true))
                            {
                                System.Diagnostics.Debug.WriteLine(string.Format("Error setting {0} to {1}",
                                    sAttributeName, sAttributeValue));
                            }
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine(string.Format("Attribute {0} not found",
                                sAttributeName));
                        }
                    }

                    if (!PWWrapper.aaOApi_SaveInstance(instP))
                    {
                        System.Diagnostics.Debug.WriteLine("Error saving instance");
                    }
                    else
                    {
                        bRetVal = true;
                    }
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine(string.Format("Error loading instance for {0},{1}",
                            iClassId, iInstanceId));
                }
            }

            PWWrapper.aaApi_DmsDataBufferFree(hProjBuf);
        }

        return bRetVal;
    }

    public static bool SetRichProjectPropertiesByIds(int iProjectId, SortedList<int, string> slPropertyIdsPropertyValues)
    {
        bool bRetVal = false;

        IntPtr hProjBuf = PWWrapper.aaApi_SelectRichProjectOfFolder(iProjectId);

        if (hProjBuf != IntPtr.Zero)
        {
            if (1 == PWWrapper.aaApi_DmsDataBufferGetCount(hProjBuf))
            {
                int iClassId = PWWrapper.aaApi_DmsDataBufferGetNumericProperty(hProjBuf,
                    (int)PWWrapper.ProjectProperty.ComponentClassId, 0);
                int iInstanceId = PWWrapper.aaApi_DmsDataBufferGetNumericProperty(hProjBuf,
                    (int)PWWrapper.ProjectProperty.ComponentInstanceId, 0);

                IntPtr instP = PWWrapper.aaOApi_LoadInstanceByIds(iClassId, iInstanceId, 0);

                if (instP != IntPtr.Zero)
                {
                    foreach (int lAttribId in slPropertyIdsPropertyValues.Keys)
                    {
                        if (lAttribId > 0)
                        {
                            string sAttributeValue = slPropertyIdsPropertyValues[lAttribId];

                            if (!PWWrapper.aaOApi_SetInstanceAttrStrValueExt(instP,
                                sAttributeValue, lAttribId, 0, true))
                            {
                                System.Diagnostics.Debug.WriteLine(string.Format("Error setting {0} to {1}",
                                    lAttribId, sAttributeValue));
                            }
                        }
                    }

                    if (!PWWrapper.aaOApi_SaveInstance(instP))
                    {
                        System.Diagnostics.Debug.WriteLine("Error saving instance");
                    }
                    else
                    {
                        bRetVal = true;
                    }
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine(string.Format("Error loading instance for {0},{1}",
                            iClassId, iInstanceId));
                }
            }

            PWWrapper.aaApi_DmsDataBufferFree(hProjBuf);
        }

        return bRetVal;
    }

    // to go in PWWrapper
    public static void SetAttributesToRichProjectProperties
    (
        int iProjectId //==>Selected folder 
    )
    {
        SortedList<string, string> slProjectPropValues =
            PWWrapper.GetProjectPropertyValuesInList(iProjectId);

        if (slProjectPropValues.Count > 0)
        {
            int iEnvId = 0, iTableId = 0, iColId = 0;

            if (PWWrapper.aaApi_GetEnvTableInfoByProject(iProjectId, ref iEnvId, ref iTableId, ref iColId))
            {
                int iAttrDefCount = PWWrapper.aaApi_SelectEnvAttrDefs(iEnvId, iTableId, -1);

                SortedList<int, string> slAttributeValues = new SortedList<int, string>();

                for (int i = 0; i < iAttrDefCount; i++)
                {
                    int iCurrColId =
                        PWWrapper.aaApi_GetEnvAttrDefNumericProperty(
                            PWWrapper.AttributeDefinitionProperty.ColumnID, i);
                    string sCurrDefVal =
                        PWWrapper.aaApi_GetEnvAttrDefStringProperty(
                            PWWrapper.AttributeDefinitionProperty.DefaultValue, i);

                    if (sCurrDefVal.StartsWith("$PROJECT#"))
                    {
                        string sPropertyName = (sCurrDefVal.Replace("$PROJECT#", "")).Replace("$", "");

                        if (slProjectPropValues.ContainsKey(sPropertyName))
                        {
                            if (!slAttributeValues.ContainsKey(iCurrColId))
                                slAttributeValues.Add(iCurrColId,
                                    slProjectPropValues[sPropertyName]);
                        }
                    }
                }

                if (slAttributeValues.Count > 0)
                {
                    ArrayList alBranchProjects = PWWrapper.GetBranchProjectNos(iProjectId, true);

                    foreach (int lProjectNo in alBranchProjects)
                    {
                        int iNumDocs = PWWrapper.aaApi_SelectDocumentsByProjectId(lProjectNo);

                        for (int j = 0; j < iNumDocs; j++)
                        {
                            PWWrapper.SetAttributesValuesFromColumnIds(lProjectNo, PWWrapper.aaApi_GetDocumentId(j),
                                slAttributeValues);
                        }
                    }
                }
            }
        }
    }

    //[DllImport("LANLUpdPropertiesExt.dll", EntryPoint = "LANLUpdateProjectPropertiesExt_UpdateDocumentProperties", CharSet = CharSet.Unicode)]
    //public static extern int SetDocPropsFromRichProj(int iProjectId);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode, EntryPoint = "aaApi_SetCurrentSession2")]
    public static extern bool SetCurrentSession2(
        System.Int64 hSession                  /* i  Session handle to set                 */
    );


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode, EntryPoint = "aaApi_GetCurrentSession2")]
    public static extern bool GetCurrentSession2(
        out System.Int64 pSession                 /* o  Current session handle                */
    );

    public static bool CheckConnectionAndWorkingFolder()
    {
        if (aaApi_GetCurrentUserId() < 1)
        {
            BPSUtilities.WriteLog("No valid ProjectWise connection.");
            return false;
        }

        // Verify that user has working directory on this machine.
        // if user doesn't have a working directory, the copy-out routine fails if called...
        string sWorkDir = aaApi_GetWorkingDirectory();
        if (string.IsNullOrEmpty(sWorkDir))
        {
            BPSUtilities.WriteLog("Error, no working directory defined for current user.");
            return false;
        }
        else if (!System.IO.Directory.Exists(sWorkDir))
        {
            try
            {
                System.IO.Directory.CreateDirectory(sWorkDir);
                // BPSUtilities.WriteLog("Created working directory '{0}' for current user.", sWorkDir);
            }
            catch
            {
                BPSUtilities.WriteLog("Error, could not create working directory '{0}' for current user.", sWorkDir);
                return false;
            }
        }

        return true;
    }


    [DllImport("dmawin.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_AdminLoginDlg(DataSourceType lDSType,
        StringBuilder lptstrDataSource, int lDSLength,
        string lpctstrUsername, string lpctstrPassword);


    [DllImport("dmscli.dll")]
    public static extern IntPtr aaApi_SelectDataSourceDataBufferByHandle(IntPtr hDatasource);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_GetConnectedUsers(bool bRefresh, ref int iUsersP, ref int iUserCountP);

    [DllImport("dmscli.dll", EntryPoint = "aaApi_SelectConnectedUser", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectConnectedUser(int iUserId); // -1 for all


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectDocumentsByProp
(
int lProjectId,      /* i  Project number (must exist)      */
int lStorageId,      /* i  Storage number                   */
int lFileType,       /* i  File type of items               */
int lItemType,       /* i  Type of items                    */
int lApplicationId,  /* i  Application type                 */
int lDepartmentId,   /* i  Department type                  */
string lpctstrFileName, /* i  File name to search              */
string lpctstrName,     /* i  Document name to search          */
string lpctstrDesc,     /* i  Document description to search   */
string lpctstrVersion,  /* i  Version to search                */
int lVersionNo,      /* i  Version sequence number          */
int lCreatorId,      /* i  Creator number                   */
int lUpdaterId,      /* i  Updater number                   */
int lLastUserId,     /* i  Last user number                 */
string lpctstrStatus,   /* i  Document status (in,out,exported)*/
int lWorkflowId,     /* i  Workflow number                  */
int lStateId         /* i  State number                     */
);


    [DllImport("dmsgen.dll", EntryPoint = "aaApi_GetProductVersionString", CharSet = CharSet.Unicode)]
    private static extern IntPtr unsafe_aaApi_GetProductVersionString();


    public static string aaApi_GetProductVersionString()
    {
        return Marshal.PtrToStringUni(unsafe_aaApi_GetProductVersionString());
    }

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_GetServerVersion(ref uint pRelease, ref uint pMajor, ref uint pMinor, ref uint pBuild);

    [DllImport("dmsgen.dll", EntryPoint = "aaApi_GetAPIVersion", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_GetAPIVersion(ref int iMajorVersionHi, ref int iMajorVersionLo,
        ref int iMinorVersion, ref int iBuildVersion);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectSet(int iSetId);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectSetReferences(int iMasterProjId, int iMasterDocId);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectSetMasters(int iChildProjId, int iChildDocId);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectSetByTypeMask(uint ulTypeMask,
      int lSetId,
      int lParentProjectId,
      int lParentDocumentId,
      int lChildProjectId,
      int lChildDocumentId
      );


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GetConnectedUserNumericProperty(UserProperty PropertyId, int lIndex);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_CreateAuditTrailRecordByGUID
    (
        AuditTrailTypes lObjectTypeId,    /* i  Object type id (AADMSAT_TYPE_*)      */
        ref Guid lpcguidObjGUID,   /* i  Affected object id                   */
        AuditTrailActions lActionTypeId,    /* i  Performed action id (AADMSAT_ACT_*)  */
        string lpctstrComment,   /* i  Operation comment                    */
        int lParam1,          /* i  Additional parameter                 */
        int lParam2,          /* i  Additional parameter                 */
        string lpctstrParam,     /* i  Additional string parameter          */
        ref Guid lpcguidGUIDParam  /* i  Additional GUID parameter            */
    );

    // dww 2013-10-04
    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_CreateAuditTrailRecordById
    (
        AuditTrailTypes lObjectTypeId,      /* i  Object type id (AADMSAT_TYPE_*)      */
        int lObjectId,          /* i  Affected object id                   */
        int lActionTypeId,      /* i  Performed action id (AADMSAT_ACT_*)  */
        string lpctstrComment,     /* i  Operation comment                    */
        int lParam1,            /* i  Additional parameter                 */
        int lParam2,            /* i  Additional parameter                 */
        string lpctstrParam,       /* i  Additional string parameter          */
        ref Guid lpcguidGUIDParam    /* i  Additional GUID parameter            */
    );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_IsUserConnected(int iUserId);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_RefreshConnectedUsers();


    public enum UserProperties : int
    {
        Name = 0,
        Desc = 1,
        Password = 2,
        Email = 3,
        Type = 4,
        SecProvider = 5
    }



    public enum UserType
    {
        DMSUser = 0,
        NTUser = 1
    }



    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_CopyAccessControl
(
uint ulFlags,           /* i  Operation flags (reserved)        */
AccessObjectType lObjectTypeFrom,   /* i  Source access object type         */
int lObjectId1From,    /* i  Source access object identifier 1 */
int lObjectId2From,    /* i  Source access object identifier 2 */
int lWorkflowIdFrom,   /* i  Workflow identifier (-1 - all)    */
int lStateIdFrom,      /* i  State identifier (-1 - all)       */
AccessObjectType lObjectTypeTo,     /* i  Target access object type         */
int lObjectId1To,      /* i  Target access identifier 1        */
int lObjectId2To       /* i  Target access identifier 2        */
);


    /// <summary>
    /// From AADMSBUFFER_COLUMN
    /// </summary>
    public enum ColumnProperty : int
    {
        ColumnID = 1,
        TableID = 2,
        SQLType = 3,
        Precision = 4,
        Scale = 5,
        Type = 6,
        Length = 7,
        Unique = 8,
        Name = 9,
        Desc = 10,
        Format = 11
    }


    [Flags]
    public enum QueryResultFlags : uint
    {
        Unfiltered = 0x00000000,        /**< all results that match the given query returned          */
        NoDocVersions = 0x00000001,     /**< document versions are not returned as results            */
        VersionsBySetting = 0x00000002, /**< versions are returned if "Show versions" setting is On   */
        VersionFieldMask = 0x00000003   /**< <mask to filter out from the flags */
    }

    public static Guid PSET_DOCUMENT_GENERIC = new Guid("4E43310A-4524-4417-AE6E-D73BD7796123");
    public static Guid PSET_ATTRIBUTE_GENERIC = new Guid("53DC3798-A8FC-415a-9EBF-2370DF192031");
    public static Guid PSET_DBTABLE_GENERIC = new Guid("84088B7B-F2E5-4bb1-BD80-8A1F6489B281");
    public static Guid PSET_DOCUMENT_ACCESS = new Guid("AD9E44D9-00A7-412c-98C5-75A561297B1C");
    public static Guid PSET_FREE_TEXT = new Guid("A45B753A-FE55-426c-B91E-366596E8CAED");
    public static Guid PSET_FILEPROPS = new Guid("3ACDE94E-CCFF-4f00-8440-152CEBF1E3C3");


    [StructLayout(LayoutKind.Sequential)]
    public struct CompactPropertyReference
    {
        /** Globally unique identifier specifying a set or class of properties. 
         * There are several property sets defined by the ProjectWise, other may be introduced by the 
         * third parties by extending the find documents functionality with registering so called 
         * 'query executors' and 'query compilers'. Property set must be recognized by the built-in or 
         * extended find documents functionality, or the query execution will fail. 
         * ProjectWise defines the following property sets: 
         * <table>
         * <tr><td>#PSET_DOCUMENT_GENERIC</td><td>{4E43310A-4524-4417-AE6E-D73BD7796123}</td></tr>
         * <tr><td>#PSET_ATTRIBUTE_GENERIC</td><td>{53DC3798-A8FC-415a-9EBF-2370DF192031}</td></tr>
         * <tr><td>#PSET_DBTABLE_GENERIC</td><td>{84088B7B-F2E5-4bb1-BD80-8A1F6489B281}</td></tr>
         * <tr><td>#PSET_DOCUMENT_ACCESS</td><td>{AD9E44D9-00A7-412c-98C5-75A561297B1C}</td></tr>
         * <tr><td>#PSET_FREE_TEXT</td><td>{A45B753A-FE55-426c-B91E-366596E8CAED}</td></tr>
         * <tr><td>#PSET_FILEPROPS</td><td>{3ACDE94E-CCFF-4f00-8440-152CEBF1E3C3}</td></tr>
         * </table> */
        public Guid guidPropertySet;
        /** A relative identifier identifying the property inside the property set. See below the name 
         * meaning for the built-in sets: 
         * <table>
         * <tr><td>#PSET_DOCUMENT_GENERIC</td><td>Not used, specify L"".</td></tr>
         * <tr><td>#PSET_ATTRIBUTE_GENERIC</td><td>Database column name in the environment table(s), e.g. L"employee_id".</td></tr>
         * <tr><td>#PSET_DBTABLE_GENERIC</td><td>Database table and column names, e.g. L"employees.employee_id".</td></tr>
         * <tr><td>#PSET_DOCUMENT_ACCESS</td><td>Not used, specify L"".</td></tr>
         * <tr><td>#PSET_FREE_TEXT</td><td>Not used, specify L"".</td></tr>
         * <tr><td>#PSET_FILEPROPS</td><td>A constructed string: file property set (guid) + DMS_SQRY_ARRAY_SEPARATOR + file property id, 
         * e.g. L"11111111-2222-3333-4444-555555555555|author".</td></tr>
         * </table> */
        public string sPropertyName;
        /** A relative identifier identifying the property inside the property set. See below the name 
         * meaning for the built-in sets: 
         * <table>
         * <tr><td>#PSET_DOCUMENT_GENERIC</td><td>One of QRY_DOC_PROP_* values.</td></tr>
         * <tr><td>#PSET_ATTRIBUTE_GENERIC</td><td>Not used, specify 0.</td></tr>
         * <tr><td>#PSET_DBTABLE_GENERIC</td><td>Not used, specify 0.</td></tr>
         * <tr><td>#PSET_DOCUMENT_ACCESS</td><td>QRY_DOCUMENT_ACCESS/QRY_PROJECT_ACCESS.</td></tr>
         * <tr><td>#PSET_FREE_TEXT</td><td>QRY_FTR_PROP_SEARCH_TEXT/QRY_FTR_PROP_SCOPE_ID.</td></tr>
         * <tr><td>#PSET_FILEPROPS</td><td>Not used, specify 0.</td></tr>
         * </table> */
        public uint uiPropertyID;
        /** A result type for the column. Normally you can set this to 0 for any ProjectWise built-in property 
         * (except #PSET_DBTABLE_GENERIC), as the executor will ignore this for the properties whose type is known to it. 
         * For other properties, like #PSET_DBTABLE_GENERIC, specify DMS_RESULT_TYPE_* value matching the real column type. 
         * Remarks Note that the properties identified by this structure for the result columns are the same kind of 
         * identifiers that are used when constructing the query criteria. But also note, that not every property 
         * identifier will be supported as both the property id in query criterion, and the property id in result column; 
         * some may be only valid for a single usage type, for example: #PSET_DOCUMENT_GENERIC/QRY_DOC_PROP_INCSUBVAULTS 
         * is not valid as a result column. */
        public uint uiResultType;
    }


    [StructLayout(LayoutKind.Sequential)]
    public struct DocumentProperties_stc
    {
        /** The count of values insPropertyarray. */
        public uint uiPropertyCount;
        /** Flags describing the expected result set. It must be one of defined QRY_RESULT_* values. */
        public QueryResultFlags eResultFlags;
        /** An array of columns requested as query results.  */
        public CompactPropertyReference[] sProperty;
    }


    [StructLayout(LayoutKind.Explicit)]
    public struct DocumentRequestColumns
    {
        [FieldOffset(0)]
        public DocumentProperties_stc properties;
        [FieldOffset(0)]
        public byte[] padding;
        [FieldOffset(0)]
        public CompactPropertyReference[] columns;
    };



    // v8i
    /*
    [StructLayout(LayoutKind.Explicit)]
    public struct _FINDDOC_RESULTCOL
    {
        [FieldOffset(0)]
        public uint dwType; // < DMS_RESULT_TYPE_
        [FieldOffset(4)]
        public int lValue;
        [FieldOffset(4)]
        public double[] lpcDoubleValue;
        [FieldOffset(4)]
        public string lpctstrValue;
        [FieldOffset(4)]
        public Guid[] lpcGuidValue;
        [FieldOffset(4)]
        public uint ulValue;
        [FieldOffset(4)]
        public ulong uint64Value;
    }; */
    [StructLayout(LayoutKind.Explicit)]
    public struct _FINDDOC_RESULTCOL
    {
        [FieldOffset(0)]
        public uint dwType; /**< DMS_RESULT_TYPE_* */
        [FieldOffset(4)]
        public uint __padding;
        [FieldOffset(8)]
        public int lValue;
        //[FieldOffset(8)]
        //public double[] lpcDoubleValue;
        //[FieldOffset(8)]
        //public string lpctstrValue;
        //[FieldOffset(8)]
        //public Guid[] lpcGuidValue;
        [FieldOffset(8)]
        public uint ulValue;
        [FieldOffset(8)]
        public ulong uint64Value;
        [FieldOffset(8)]
        ulong int64Value;
    };

    // XM

    //[StructLayout(LayoutKind.Explicit)]
    //public struct _FINDDOC_RESULTCOL_XM
    //{
    //    [FieldOffset(0)]
    //    uint dwType; /**< DMS_RESULT_TYPE_* */


    //    [FieldOffset(4)]
    //    int lValue;
    //    [FieldOffset(4)]
    //    double[] lpcDoubleValue;
    //    [FieldOffset(4)]
    //    string lpctstrValue;
    //    [FieldOffset(4)]
    //    Guid[] lpcGuidValue;
    //    [FieldOffset(4)]
    //    uint ulValue;
    //};



    [StructLayout(LayoutKind.Sequential)]
    public struct _FINDDOC_RESULT
    {
        public uint dwColumnCount;
        public _FINDDOC_RESULTCOL[] pCol;
    }


    [StructLayout(LayoutKind.Sequential)]
    public struct _FINDDOC_RESULTS
    {
        public uint dwRowCount;
        public _FINDDOC_RESULT[] pRow;
    }


    public enum QueryProperty : int
    {
        QRY_DOC_PROP_NONE = 0,
        QRY_DOC_PROP_ENVIRONMENT_ID = 1,
        QRY_DOC_PROP_PROJ_ID = 2,
        QRY_DOC_PROP_PROJ_NAME = 3,
        QRY_DOC_PROP_PROJ_DESC = 4,
        QRY_DOC_PROP_FILENAME = 5,
        QRY_DOC_PROP_NAME = 6,
        QRY_DOC_PROP_DESC = 7,
        QRY_DOC_PROP_VERSION = 8,
        QRY_DOC_PROP_VERSIONSEQ = 9,
        QRY_DOC_PROP_CREATORID = 10,
        QRY_DOC_PROP_UPDATERID = 11,
        QRY_DOC_PROP_DMSSTATUS = 12,
        QRY_DOC_PROP_LASTUSERID = 13,
        QRY_DOC_PROP_FILETYPE = 14,
        QRY_DOC_PROP_ITEMTYPE = 15,
        QRY_DOC_PROP_STORAGEID = 16,
        QRY_DOC_PROP_WORKFLOWID = 17,
        QRY_DOC_PROP_STATEID = 18,
        QRY_DOC_PROP_APPLICATIONID = 19,
        QRY_DOC_PROP_DEPARTMENTID = 20,
        QRY_DOC_PROP_INCSUBVAULTS = 21,
        QRY_DOC_PROP_FINAL_STATUS = 22,
        QRY_DOC_PROP_FINAL_USER = 23,
        QRY_DOC_PROP_FINAL_DATE = 24,
        QRY_DOC_PROP_LOCATIONID = 25,
        QRY_DOC_PROP_FILE_REVISION = 26,
        QRY_DOC_PROP_OVERLAPS = 27,
        QRY_DOC_PROP_MIMETYPE = 28,

        QRY_DOC_PROP_ID = 101,
        QRY_DOC_PROP_PROPOSALNO = 102,
        QRY_DOC_PROP_SIZE = 104,
        QRY_DOC_PROP_SETID = 105,
        QRY_DOC_PROP_SETTYPE = 106,
        QRY_DOC_PROP_ORIGINALNO = 107,
        QRY_DOC_PROP_IS_OUT_TO_ME = 108,
        QRY_DOC_PROP_CREATE_TIME = 109,
        QRY_DOC_PROP_UPDATE_TIME = 110,
        QRY_DOC_PROP_DMSDATE = 111,
        QRY_DOC_PROP_NODE = 112,
        QRY_DOC_PROP_ACCESS = 113,
        QRY_DOC_PROP_MANAGERID = 114,
        QRY_DOC_PROP_FILE_UPDATERID = 115,
        QRY_DOC_PROP_LAST_RT_LOCKERID = 116,
        QRY_DOC_PROP_ITEM_FLAGS = 117,
        QRY_DOC_PROP_FILE_UPDATE_TIME = 118,
        QRY_DOC_PROP_LAST_RT_LOCK_TIME = 119,
        QRY_DOC_PROP_MGRTYPE = 120,
        QRY_DOC_PROP_DOCGUID = 121,
        QRY_DOC_PROP_PROJGUID = 122,
        QRY_DOC_PROP_ORIGGUID = 123,

        QRY_DOC_PROP_PROJ_VERSIONNO = 124,
        QRY_DOC_PROP_PROJ_MANAGERID = 125,
        QRY_DOC_PROP_PROJ_STORAGEID = 126,
        QRY_DOC_PROP_PROJ_CREATORID = 127,
        QRY_DOC_PROP_PROJ_UPDATERID = 128,
        QRY_DOC_PROP_PROJ_WORKFLOWID = 129,
        QRY_DOC_PROP_PROJ_STATEID = 130,
        QRY_DOC_PROP_PROJ_TYPE = 131,
        QRY_DOC_PROP_PROJ_ARCHIVEID = 132,
        QRY_DOC_PROP_PROJ_ISPARENT = 133,
        QRY_DOC_PROP_PROJ_CODE = 134,
        QRY_DOC_PROP_PROJ_VERSION = 135,
        QRY_DOC_PROP_PROJ_CREATE_TIME = 136,
        QRY_DOC_PROP_PROJ_UPDATE_TIME = 137,
        QRY_DOC_PROP_PROJ_CONFIG = 138,
        QRY_DOC_PROP_PROJ_PARENTID = 140,
        QRY_DOC_PROP_PROJ_MGRTYPE = 141,
        QRY_DOC_PROP_PROJ_ACCESS = 142,
        QRY_DOC_PROP_PROJ_PROJGUID = 143,
        QRY_DOC_PROP_PROJ_PPRJGUID = 144,

        QRY_PROP_ACCUMULATED_TEXTS = 200,   /**< used for FTR queries only */
        QRY_PROP_DATASOURCE_GUID = 201,   /**< used for FTR internally */

        QRY_PROP_VIEW_ID = 250,   /**< used as UI hint only */

        QRY_DOC_PROP_CHECKOUT_USERID = 301,   /**< equivalent to QRY_CHKLOC_PROP_USERID + QRY_CHKLOC_PROP_TYPEFLAGS ('CO','CS','XS')   */
        QRY_DOC_PROP_CHECKOUT_NODE = 302,   /**< equivalent to QRY_CHKLOC_PROP_NODE + QRY_CHKLOC_PROP_TYPEFLAGS ('CO','CS','XS')     */
        QRY_DOC_PROP_CHECKOUT_COUTTIME = 303,    /**< equivalent to QRY_CHKLOC_PROP_COUTTIME + QRY_CHKLOC_PROP_TYPEFLAGS ('CO','CS','XS') */

        QRY_FTR_PROP_SEARCH_TEXT = 1,
        QRY_FTR_PROP_SCOPE_ID = 2
    }


    public enum RestrictionRelation : int
    {
        DMS_RELATION_NONE = 0,   /**< No relation specified */
        DMS_RELATION_EQUAL = 1,   /**<  =   */
        DMS_RELATION_NOTEQUAL = 2,   /**<  <>  */
        DMS_RELATION_LESSTHAN = 3,   /**<  <   */
        DMS_RELATION_GREATERTHAN = 4,   /**<  >   */
        DMS_RELATION_GREATEROREQUAL = 5,   /**<  >=  */
        DMS_RELATION_LESSOREQUAL = 6,   /**<  <=  */
        DMS_RELATION_BETWEEN = 7,   /**< BETWEEN     */
        DMS_RELATION_ISNULL = 8,   /**< IS NULL     */
        DMS_RELATION_ISNOTNULL = 9,   /**< IS NOT NULL */
        DMS_RELATION_ISLIKE = 10,   /**< LIKE        */
        DMS_RELATION_IN = 11,   /**< IN          */
        DMS_RELATION_NOTIN = 12,   /**< NOT IN      */
        DMS_RELATION_INNERJOIN = 13,   /**< INNER JOIN  */
        DMS_RELATION_LEFTOUTERJOIN = 14,   /**< LEFT OUTER JOIN */
        DMS_RELATION_RIGHTOUTERJOIN = 15,   /**< RIGHT OUTER JOIN */
        DMS_RELATION_ISNOTLIKE = 16,   /**< NOT LIKE    */
        DMS_RELATION_NOTBETWEEN = 17,   /**< NOT BETWEEN */
        DMS_RELATION_NODE_OR_SUBNODE = 18,   /**< folders & subfolders, etc. */
        DMS_RELATION_DERIVED_TYPE = 19,   /**< folders & subfolders, etc. */

        /** valid for FTR criteria only */
        DMS_RELATION_EXPRESSION = 1024,

        DMS_RELATION_INCL_PHRASE = 5000, /**< includes whole phrase */
        DMS_RELATION_INCL_ANYWORD = 5001, /**< includes any word */
        DMS_RELATION_INCL_ALLWORDS = 5002, /**< includes all words */
        DMS_RELATION_INCL_NONEOFWORDS = 5003, /**< does not include any of the words */

        DMS_RELATION_SUBQUERY_COLUMN = 6001, /**< subquery column (value == column id) [GT_SUBQUERY]*/
        DMS_RELATION_UNION_ALL = 6002, /**< used witn GT_UNION/PSET_DBSUBQUERY */

        DMS_RELATION_SUBORGROUP = 7003, /**< used to define sub group in query defined by UT_SUBGROUP_TAG. OR operand will be used to join criteria in sub group*/
        DMS_RELATION_SUBANDGROUP = 7004  /**< used to define sub group in query defined by UT_SUBGROUP_TAG. AND operand will be used to join criteria in sub group*/
    }


    [Flags]
    public enum CriteriaGroupType : uint
    {
        GT_RESTRICTION = 0x00000000,       /**< Group defines result restrictions */
        GT_JOIN = 0x00000001,       /**< Group defines relation conditions between tables */
        GT_ROWSET_ID = 0x00000002,       /**< Defines table columns to be unique per returned row set */
        GT_SUBQUERY = 0x00000003,       /**< Group defines SQL sub-select columns */
        GT_FREE_TEXT = 0x00000004,       /**< Group defines free text search restrictions */
        GT_UNION = 0x00000005,       /**< Group defines union, etc. of 2 or more SQL sub-selects */
        GT_QRY_SPLIT = 0x00000006,       /**< Defines psetSimpleSearch SQL query split group */
        GT_FIELD_MASK = 0x00000007        /**< Mask to filter out from the flags */
    }


    [Flags]
    public enum CriteriaValueType : uint
    {
        VT_SINGLE_VALUE = 0x00000000,       /**< Criterion contains a single value. Specify this flag if you define criterion like Document.Id = 123 */
        VT_NO_VALUE = 0x00000008,       /**< Criterion contains no value. Specify this flag with criterion like Table.Column IS NULL */
        VT_VALUE_ARRAY = 0x00000010,       /**< Criterion contains an array of values. Specify this flag with criterion like Document.VersionSeq BETWEEN 12 AND 13. In such case multiple values are put to the same value field and need to be setarated by DMS_SQRY_ARRAY_SEPARATOR, for example: "12|13".Criterion contains an array of values.*/
        VT_PROPERTY_ID = 0x00000018,       /**< Criterion contains a reference to another object property in the value field. Specify this flag with criterion defining join condition including properties from two tables. The property reference needs to be coverted to a specifically formatted string. */
        VT_VALUE_BLOB = 0x00001000,       /**< Criterion contains a binary data */
        VT_FIELD_MASK = 0x0000f018        /**< Mask to filter out from the flags */
    }


    [Flags]
    public enum CriteriaUsageType : uint
    {
        UT_REGULAR = 0x00000000,       /**< regular criterion usage */
        UT_UI_HINT_ONLY = 0x00000020,       /**< not a criterion - used only in user interface */
        UT_OVERLAY = 0x00000060,       /**< not a criterion - used only internally to extend other criteria */
        UT_SUBQUERY_TAG = 0x000000a0,       /**< not a criterion - specifies the sub-query (Union argument) the group applies to */
        UT_SUBGROUP_TAG = 0x000000c0,       /**< not a criterion - specifies the sub "or" group */
        UT_FIELD_MASK = 0x00000fe0        /**< <mask to filter out from the flags> */
    }


    [Flags]
    public enum CriteriaFlags : uint
    {
        GT_RESTRICTION = 0x00000000,       /**< Group defines result restrictions */
        GT_JOIN = 0x00000001,       /**< Group defines relation conditions between tables */
        GT_ROWSET_ID = 0x00000002,       /**< Defines table columns to be unique per returned row set */
        GT_SUBQUERY = 0x00000003,       /**< Group defines SQL sub-select columns */
        GT_FREE_TEXT = 0x00000004,       /**< Group defines free text search restrictions */
        GT_UNION = 0x00000005,       /**< Group defines union, etc. of 2 or more SQL sub-selects */
        GT_QRY_SPLIT = 0x00000006,       /**< Defines psetSimpleSearch SQL query split group */
        GT_FIELD_MASK = 0x00000007,       /**< Mask to filter out from the flags */
        VT_SINGLE_VALUE = 0x00000000,       /**< Criterion contains a single value. Specify this flag if you define criterion like Document.Id = 123 */
        VT_NO_VALUE = 0x00000008,       /**< Criterion contains no value. Specify this flag with criterion like Table.Column IS NULL */
        VT_VALUE_ARRAY = 0x00000010,       /**< Criterion contains an array of values. Specify this flag with criterion like Document.VersionSeq BETWEEN 12 AND 13. In such case multiple values are put to the same value field and need to be setarated by DMS_SQRY_ARRAY_SEPARATOR, for example: "12|13".Criterion contains an array of values.*/
        VT_PROPERTY_ID = 0x00000018,       /**< Criterion contains a reference to another object property in the value field. Specify this flag with criterion defining join condition including properties from two tables. The property reference needs to be coverted to a specifically formatted string. */
        VT_VALUE_BLOB = 0x00001000,       /**< Criterion contains a binary data */
        VT_FIELD_MASK = 0x0000f018,       /**< Mask to filter out from the flags */
        UT_REGULAR = 0x00000000,       /**< regular criterion usage */
        UT_UI_HINT_ONLY = 0x00000020,       /**< not a criterion - used only in user interface */
        UT_OVERLAY = 0x00000060,       /**< not a criterion - used only internally to extend other criteria */
        UT_SUBQUERY_TAG = 0x000000a0,       /**< not a criterion - specifies the sub-query (Union argument) the group applies to */
        UT_SUBGROUP_TAG = 0x000000c0,       /**< not a criterion - specifies the sub "or" group */
        UT_FIELD_MASK = 0x00000fe0        /**< <mask to filter out from the flags> */
    }


    public enum CriterionDataType : int
    {
        AADMS_ATTRFORM_DATATYPE_STRING = 1,
        AADMS_ATTRFORM_DATATYPE_INT = 2,
        AADMS_ATTRFORM_DATATYPE_UINT = 3,
        AADMS_ATTRFORM_DATATYPE_FLOAT = 4,
        AADMS_ATTRFORM_DATATYPE_DATE_TO_DAY = 5,
        AADMS_ATTRFORM_DATATYPE_DATE_TO_SEC = 6,
        AADMS_ATTRFORM_DATATYPE_STRING_AS_DATE_TO_DAY = 7,
        AADMS_ATTRFORM_DATATYPE_STRING_AS_DATE_TO_SEC = 8,
        AADMS_ATTRFORM_DATATYPE_GUID = 9,
        AADMS_ATTRFORM_DATATYPE_BIGINT = 10,
        AADMS_ATTRFORM_DATATYPE_UBIGINT = 11,
        AADMS_ATTRFORM_DATATYPE_TIMESPAN = 12
    }


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaApi_SQueryCriDataBufferAddCriterion(IntPtr hQueryCriBuffer,
        int iOrGroup,
        CriteriaFlags iFlags, // CriteriaGroupType | CriteriaUsageType | CriteriaValueType
        ref Guid pGuidPropertySet,
        string sPropertyName,
        QueryProperty iPropertyId,
        RestrictionRelation iRelationId,
        CriterionDataType iFieldType,
        string sFieldValue);


    /* [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
     public static extern bool aaApi_FindDocumentsToBuffer(IntPtr hCriteriaBuf,
         DocumentRequestColumns resultCols,
         // DocumentProperties_stc[] resultCols, 
         ref bool bCancel,
         // ref IntPtr );
         ref IntPtr findDocResults);
     // ref _FINDDOC_RESULTS[] findDocResults);
 */
    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_FindDocumentsToBuffer(IntPtr hCriteriaBuf,
        DocumentRequestColumns resultCols,
        // DocumentProperties_stc[] resultCols, 
        ref bool bCancel,
    // ref IntPtr );
    //ref IntPtr findDocResults);
    ref _FINDDOC_RESULTS findDocResults);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_FindDocumentsToBuffer(IntPtr hCriteriaBuf,
        DocumentRequestColumns resultCols,
        // DocumentProperties_stc[] resultCols, 
        ref bool bCancel,
    // ref IntPtr );
    //ref IntPtr findDocResults);
    ref IntPtr findDocResults);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_FindDocuments(IntPtr hCriteriaBuf,
        // DocumentProperties_stc[] resultCols, 
        DocumentRequestColumns resultCols,
        IntPtr pCallback,
        ref bool bCancel,
        uint uiTimeToWaitForChunk,
        uint uiItemCountInChunk);

    [DllImport("dmawin.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_UpdateDocumentWindows();

    [DllImport("dmawin.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_UpdateProjectWindows();

    [DllImport("dmawin.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_UpdateTreeItems(IntPtr handle);

    [Flags]
    public enum FindDscItemFlags : uint
    {
        AAFINDDSCITEM_VISIBLE = 0x00000001,
        AAFINDDSCITEM_ITEMTYPE = 0x00000002,
        AAFINDDSCITEM_ITEMID = 0x00000004,
        AAFINDDSCITEM_ITEMTEXT = 0x00000008,
        AAFINDDSCITEM_EXPAND = 0x00000010,
        AAFINDDSCITEM_ITEMDATA = 0x00000020,
        AAFINDDSCITEM_ALLVISIBLE = 0x00001000,
        AAFINDDSCITEM_AFFECTSPARENT = 0x00002000,
        AAFINDDSCITEM_KEEPSELECTION = 0x00004000,
        AAFINDDSCITEM_RECURSIVE = 0x00008000,
    }

    public enum DscItemTypes : int
    {
        DSCITYPE_NONE = 0,
        DSCITYPE_DEFAULT_ROOT = 1,
        DSCITYPE_DATASOURCE_ROOT = 2,
        DSCITYPE_DATASOURCE = 3,
        DSCITYPE_VAULTROOT = 4,
        DSCITYPE_DMSVAULT = 5,
        DSCITYPE_DMSDOCUMENT = 6,
        DSCITYPE_MSGFLDRROOT = 11,
        DSCITYPE_MSGFOLDER = 12,
        DSCITYPE_MSGSPECIALFOLDER = 13,
        DSCITYPE_UWSPROOT = 14,
        DSCITYPE_UWSPACE = 15,
        DSCITYPE_UWSPACE_FOLDER = 16,
        DSCITYPE_UWSPACE_USER = 17,
        DSCITYPE_DMSDOCUMENTSET = 18,
        DSCITYPE_HTMLPG = 19,
        DSCITYPE_SEARCH_RESULTS = 20,
        DSCITYPE_SAVED_QUERY_ROOT = 21,
        DSCITYPE_SAVED_QUERY_TYPE = 22,
        DSCITYPE_SAVED_QUERY_ITEM = 23,
        DSCITYPE_COMPONENT_MAIN = 24,
        DSCITYPE_COMPONENT = 25,
        DSCITYPE_LINKSET_ROOT = 26,
        DSCITYPE_ALL = 0x0FFFFFFF
    }

    /***********************************************************************
    /* Predefined ids for DSCITYPE_UWSPACE_FOLDER
    ***********************************************************************/
    public enum CustomFolderItemIds : int
    {
        DSCI_UWSPFLDR_PRIVATE = 1,
        DSCI_UWSPFLDR_GLOBAL = 2,
        DSCI_UWSPFLDR_OTHERS = 3
    }

    [DllImport("dmawin.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_RefreshDscTreeSubItemsByParams
(
   IntPtr hWndTree,      /* i  Tree window handle         */
   FindDscItemFlags ulFlags,       /* i  Search flags               */
   DscItemTypes lTypeId,       /* i  Item type to search for    */
   int lItemId,       /* i  Item ID to search for      */
   IntPtr lpbyItemData,  /* i  Item data to search for    */
   int lItemData,     /* i  Size of lpbyItemData       */
   IntPtr fpCompare      /* i  Compare callback function  */
);

    [DllImport("dmawin.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaApi_GetMainFrameWindow();

    [DllImport("dmawin.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaOApi_CreateProgressDlg(IntPtr hWndParent,
        string sTitle,
        IntPtr lpParam);

    [DllImport("dmawin.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaOApi_DestroyProgressDlg(IntPtr hWnd);

    [DllImport("dmawin.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaApi_GetMainDscTree();

    //[DllImport("dmawin.dll", CharSet = CharSet.Unicode)]
    //public static extern IntPtr aaApi_GetMainDocumentList();

    [DllImport("dmawin.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaApi_FindDscTreeDataSourceItem
(
   IntPtr hWndTree,       /* i  Tree window handle         */
   IntPtr hDataSource,    /* i  Datasource handle to find  */
   FindDscItemFlags ulFlags,        /* i  Find flags                 */
   IntPtr lphItem         /* o  Last found tree item       */
);

    [DllImport("dmawin.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_DscTreeEnsureVisibleHTreeItem
(
   IntPtr hWndTree,       /* i  Tree window handle         */
   IntPtr lphItem         /* i  tree item to ensure visible       */
);

    [DllImport("dmawin.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaApi_FindDscTreeItem
    (
   IntPtr hWndTree,       /* i  Tree window handle       */
   IntPtr hParent,        /* i  Parent item handle       */
   DscItemTypes lTypeId,        /* i  Item type to search for  */
   int lItemId,        /* i  Item ID to search for    */
   string lpctstrText,    /* i  Item text to search for  */
   FindDscItemFlags ulFlags         /* i  Search flags             */
        );

    [DllImport("dmawin.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaApi_FindDscTreeItemByName
    (
   IntPtr hWndTree,       /* i  Tree window handle       */
   string sDatasourceName    /* i  Datasource Name  */
    );


    [DllImport("dmawin.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_ProjectTreeSelectItem(IntPtr hDscTree, int iProjectId);

    [DllImport("dmawin.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_ProjectTreeSetDocList(IntPtr hDscTree, IntPtr hWndDocList);

    [DllImport("dmawin.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaApi_CreateDocumentList(IntPtr hWndParent, int iLeft, int iTop, int iWidth, int iHeight, DocumentListDefinitions ulFlags);

    [DllImport("dmawin.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_DocListUpdateItemStatus(IntPtr hWndDocList, int lProjectId, int lDocumentId, DocListUpdateTypeMasks ulFlags);

    [DllImport("dmawin.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_DocListSynchronizeAttributeSheet(IntPtr hWndDocList);

    [DllImport("dmawin.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_AttributeSheetModifyDocument(int lProjectId, int lDocumentId, int lAttributeId);

    [DllImport("dmawin.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_EnableMenuCommand(int iMenuType /* 0 for all */,
        MenuCommandIds menuCmdId, bool bEnable);

    [DllImport("dmawin.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_EnableMenuCommand(int iMenuType /* 0 for all */,
        int iMenuCmdId, bool bEnable);

    [DllImport("dmsgen.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_AddHook(int lHookId, int lHookType, HookFunction lpfnHook);

    [DllImport("dmsgen.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_AddHook(int lHookId, int lHookType, DoumentHookFunction lpfnHook);

    [DllImport("dmsgen.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_AddHook(int lHookId, PWWrapper.HookTypes lHookType, GenericHookFunction lpfnHook);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_CreateCustomHierarchy
(
   ref int lplCustHrchyId,  /* o  Custom hierarchy id              */
   int lUserId,         /* i  User id (0 - global, -1 - current user)   */
   int lParentId,       /* i  Parent workspace id  (top level 0)  */
   uint ulFlags,         /* i  Folder flags (ignored, pass 0)         */
   string lpctstrName,     /* i  Workspace name                   */
   string lpctstrDesc      /* i  Workspace description            */
);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectCustomHierarchiesByUserId
(
   int lUserId           /* i  User id (required)  */
);

    public enum CustomHierarchyProperties : int
    {
        UWSP_PROP_USER = 1,
        UWSP_PROP_CUSTHRCHY = 2,
        UWSP_PROP_PARENT = 3,
        UWSP_PROP_FLAGS = 4, // Specifies custom hierarchy parameter mask. This property is not used in the current version of ProjectWise. 
        UWSP_PROP_HASSUBITEMS = 5,
        UWSP_PROP_NAME = 6,
        UWSP_PROP_DESC = 7
    }

    [DllImport("dmscli.dll", EntryPoint = "aaApi_GetCustomHierarchyStringProperty", CharSet = CharSet.Unicode)]
    private static extern IntPtr unsafe_aaApi_GetCustomHierarchyStringProperty(CustomHierarchyProperties PropertyId, int lIndex);

    public static string aaApi_GetCustomHierarchyStringProperty(CustomHierarchyProperties PropertyId, int lIndex)
    {
        return Marshal.PtrToStringUni(unsafe_aaApi_GetCustomHierarchyStringProperty(PropertyId, lIndex));
    }

    [DllImport("dmscli.dll", EntryPoint = "aaApi_GetCustomHierarchyNumericProperty", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GetCustomHierarchyNumericProperty(CustomHierarchyProperties PropertyId, int lIndex);

    public enum CustomHierarchMemberItemType : int
    {
        AADMSUWITYPE_PROJECT = 1,
        AADMSUWITYPE_DOCUMENT = 2
    }

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_AddCustomHierarchyMember
(
   int iUserId,        /* i  User id (-1 - all or current user, 0 for global (admin only))   */
   int iCustomHierarcyId,        /* i  Item id              */
   CustomHierarchMemberItemType iMemberItemType,      /* i  Member Type          */
   int iMemberId1,      /* i  Member id            */
   int iMemberId2      /* i  Member id 2          */
);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_FreeLinkDataInsertDesc();


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_CreateLinkDataAndLink
(
int lTableId,                /* i  Link data table id      */
int lLinkType,               /* i  Link type (1 for document) */
int lObjectId1,              /* i  Object ID1 (project id)  */
int lObjectId2,              /* i  Object ID2 (document id) */
ref int lplColumnId,             /* o  Unique value column id  */
StringBuilder lptstrValueBuffer,       /* o  Unique value buffer     */
int lLenghtBuffer            /* i  Buffer length           */
);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_CreateLinkData(int lTableId,
  ref int lplColumnId,
  StringBuilder lptstrValueBuffer,
  int llengthBuffer
 );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_CreateLink(int lProjectId,
  int lDocumentId,
  int lTableId,
  int lColumnId,
  string lpctstrValue);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_SetLinkDataColumnValue(int tableID, int columnID, string columnValue);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectColumnsByTable(int iTableId);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectColumn(int iTableId, int iColumnId);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GetColumnNumericProperty(ColumnProperty PropertyId, int iIndex);


    [DllImport("dmscli.dll", EntryPoint = "aaApi_GetColumnStringProperty", CharSet = CharSet.Unicode)]
    private static extern IntPtr unsafe_aaApi_GetColumnStringProperty(ColumnProperty PropertyId, int lIndex);


    public static string aaApi_GetColumnStringProperty(ColumnProperty PropertyId, int lIndex)
    {
        return Marshal.PtrToStringUni(unsafe_aaApi_GetColumnStringProperty(PropertyId, lIndex));
    }


    public enum WorkflowProperty : int
    {
        ID = 1,
        Type = 2,
        Name = 3,
        Desc = 4
    }


    public enum StateProperty : int
    {
        ID = 1,
        Name = 2,
        Desc = 3
    }

    public enum ValueListProperty : int
    {
        EnvironmentID = 1,
        TableID = 2,
        ColumnID = 3,
        ValueID = 4,
        Value = 5,
        Description = 6
    }

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectEnvValListItems(int iEnvironmentId, int iTableId, int iColumnId, int iValueId);

    [DllImport("dmscli.dll", EntryPoint = "aaApi_GetEnvValListStringProperty", CharSet = CharSet.Unicode)]
    private static extern IntPtr _aaApi_GetEnvValListStringProperty(ValueListProperty PropertyId, int lIndex);

    public static string aaApi_GetEnvValListStringProperty(ValueListProperty PropertyId, int lIndex)
    {
        return Marshal.PtrToStringUni(_aaApi_GetEnvValListStringProperty(PropertyId, lIndex));
    }

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GetEnvValListNumericProperty(ValueListProperty PropertyId, int lIndex);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectAllWorkflows();

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectWorkflow(int iWorkflowId);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GetWorkflowId(int lIndex);

    [DllImport("dmscli.dll", EntryPoint = "aaApi_GetWorkflowStringProperty", CharSet = CharSet.Unicode)]
    private static extern IntPtr unsafe_aaApi_GetWorkflowStringProperty(WorkflowProperty PropertyId, int lIndex);


    public static string aaApi_GetWorkflowStringProperty(WorkflowProperty PropertyId, int lIndex)
    {
        return Marshal.PtrToStringUni(unsafe_aaApi_GetWorkflowStringProperty(PropertyId, lIndex));
    }

    public enum WorkflowStateProperty : int
    {
        WorkflowID = 1,
        StateID = 2,
        PreviousState = 3,
        NextState = 4
    }

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectWorkflowStateLinks(int iWorkflowId, int iState);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GetWorkflowStateLinkNumericProperty(WorkflowStateProperty iPropertyId, int iIndex);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_CreateWorkflow
    (
        ref int iWorkflowId,
        PWWrapper.WorkflowTypes lWorkflowType,
        string sWorkflowName,
        string sWorkflowDesc
    );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_CreateState
    (
        ref int iStateId,
        string sStateName,
        string sStateDesc
    );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_AddStateToWorkflow
    (
        int iWorkflowId,
        int iStateId,
        int iPrevStateId,
        int iNextStateId
    );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectAllStates();

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectState(int iStateId);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectStatesByWorkflow(int iWorkflowId);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_SelectStatesNotInWorkflow(int iWorkflowId);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GetStateId(int lIndex);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_DeleteEnvValListItems(int lEnvironmentId, int lTableId, int lColumnId, int lValueId);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_CreateEnvValListItem(int lEnvironmentId,
      int lTableId,
      int lColumnId,
      int lValueId,
      string sListValue,
      string sListValueDesc
     );

    [DllImport("dmscli.dll", EntryPoint = "aaApi_GetStateStringProperty", CharSet = CharSet.Unicode)]
    private static extern IntPtr unsafe_aaApi_GetStateStringProperty(StateProperty PropertyId, int lIndex);


    public static string aaApi_GetStateStringProperty(StateProperty PropertyId, int lIndex)
    {
        return Marshal.PtrToStringUni(unsafe_aaApi_GetStateStringProperty(PropertyId, lIndex));
    }

    public static int GetInterfaceId(string sInterfaceName)
    {
        if (!string.IsNullOrEmpty(sInterfaceName))
        {
            for (int i = 0; i < PWWrapper.aaApi_SelectAllGuis(); i++)
            {
                string sIntName = PWWrapper.aaApi_GetGuiStringProperty(GuiProperty.Name, i);

                if (sIntName.ToLower() == sInterfaceName.ToLower())
                {
                    return PWWrapper.aaApi_GetGuiId(i);
                }
            }
        }

        return 0;
    }

    public static bool IsOracle()
    {
        //#define AADMS_DBTYPE_UNKNOWN   0 
        //            Datasource database type is unknown.
        //#define AADMS_DBTYPE_ORACLE   1 
        //  Datasource database type is Oracle.
        //#define AADMS_DBTYPE_SQLSERVER   2 
        //  Datasource database type is Microsoft SQL Server.
        return (1 == aaApi_GetActiveDatasourceType());
    }

    public static int GetWorkflowId(string sWorkflowName)
    {
        if (!string.IsNullOrEmpty(sWorkflowName))
        {
            for (int i = 0; i < PWWrapper.aaApi_SelectAllWorkflows(); i++)
            {
                string sWfName = PWWrapper.aaApi_GetWorkflowStringProperty(
                    PWWrapper.WorkflowProperty.Name, i);

                if (sWfName.ToLower() == sWorkflowName.ToLower())
                {
                    return PWWrapper.aaApi_GetWorkflowId(i);
                }
            }
        }

        return 0;
    }

    public class ProjectWiseApplication
    {
        public int ID;// { get; set; }
        public string Name;// { get; set; }
        public int ViewerId;// { get; set; }

        public ProjectWiseApplication(int iID, string sName)
        {
            ID = iID;
            Name = sName;
        }

        public ProjectWiseApplication(int iID, string sName, int iViewerId)
        {
            ID = iID;
            Name = sName;
            ViewerId = iViewerId;
        }
    }

    public class ProjectWiseUser
    {
        public int ID;// { get; set; }
        public string Name;// { get; set; }
        public string Description;// { get; set; }
        public string SecurityProvider;// { get; set; }
        public string EMail;// { get; set; }
        public string UserType;// { get; set; }
        public bool Disabled;// { get; set; }
        public string Identity; // will populate in other function
        public string IdentityProvider; // will populate in other function and add here later

        public ProjectWiseUser(int iID, string sName)
        {
            ID = iID;
            Name = sName;
        }

        public ProjectWiseUser(int iID, string sName, string sDescription, string sSecProvider, string sEmail, string sUserType, bool bDisabled)
        {
            ID = iID;
            Name = sName;
            Description = sDescription;
            SecurityProvider = sSecProvider;
            EMail = sEmail;
            UserType = sUserType;
            Disabled = bDisabled;
        }

        public ProjectWiseUser (int iID)
        {
            if (1 == PWWrapper.aaApi_SelectUser(iID))
            {
                ID = iID;
                Name = PWWrapper.aaApi_GetUserStringProperty(UserProperty.Name, 0);
                Description = PWWrapper.aaApi_GetUserStringProperty(UserProperty.Desc, 0);
                SecurityProvider = PWWrapper.aaApi_GetUserStringProperty(UserProperty.SecProvider, 0);
                EMail = PWWrapper.aaApi_GetUserStringProperty(UserProperty.Email, 0);
                UserType = PWWrapper.aaApi_GetUserStringProperty(UserProperty.Type, 0);
                Disabled = (1 == PWWrapper.aaApi_GetUserNumericProperty(UserProperty.Flags, 0));
            }
        }
    }

    public class ProjectWiseUserList
    {
        public int ID;// { get; set; }
        public string Name;// { get; set; }
        public string Description;// { get; set; }
        public int ListType;// { get; set; }
        public int Owner;// { get; set; }

        // [BMF Added on 07/21/2019]
        public string ListTypeName;// { get; set; }
        //public string OwnerName;// { get; set; }

        public DataTable GetMembers()
        {
            DataTable dt = new DataTable();

            IntPtr iPtr = PWWrapper.aaApi_SelectUserListMemberDataBufferByProp(this.ID, -1, -1);
            if (iPtr == IntPtr.Zero)
            {
                return dt;
            }
            int iCount = PWWrapper.aaApi_DmsDataBufferGetCount(iPtr);

            dt.Columns.Add(new DataColumn("ID", typeof(int)));
            dt.Columns.Add(new DataColumn("MemberType", typeof(string)));
            dt.Columns.Add(new DataColumn("MemberName", typeof(string)));

            for (int i = 0; i < iCount; i++)
            {
                int iMemberID = PWWrapper.aaApi_DmsDataBufferGetNumericProperty(iPtr, (int)PWWrapper.UserListMemberProperty.MemberID, i);
                int iMemberType = PWWrapper.aaApi_DmsDataBufferGetNumericProperty(iPtr, (int)PWWrapper.UserListMemberProperty.MemberType, i);

                string sMemberType = PWWrapper.GetUserListOrGroupMemberTypeName(iMemberType);
                string sMemberName = PWWrapper.GetUserListOrGroupMemberName(iMemberID, iMemberType);

                DataRow dr = dt.NewRow();

                dr["ID"] = iMemberID;
                dr["MemberType"] = sMemberType;
                dr["MemberName"] = sMemberName;

                dt.Rows.Add(dr);
            }
            PWWrapper.aaApi_DmsDataBufferFree(iPtr);

            return dt;
        }

        public ProjectWiseUserList(int iID, string sName)
        {
            ID = iID;
            Name = sName;
        }

        public ProjectWiseUserList(int iID, string sName, string sDescription, int iOwner, int iType)
        {
            ID = iID;
            Name = sName;
            Description = sDescription;
            Owner = iOwner;
            ListType = iType;
        }

        // [BMF Added on 07/21/2019]
        public ProjectWiseUserList(int iID, string sName, string sDescription, int iOwner, int iType, string sType)
        {
            ID = iID;
            Name = sName;
            Description = sDescription;
            ListType = iType;
            ListTypeName = sType;
            Owner = iOwner;
        }

    }

    public class ProjectWiseGroup
    {
        public int ID { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
        public string GroupType { get; set; }
        public string SecurityProvider { get; set; }

        public DataTable GetMembers()
        {
            DataTable dt = new DataTable();

            IntPtr iPtr = PWWrapper.aaApi_SelectGroupMemberDataBufferById(this.ID, -1);
            if (iPtr == IntPtr.Zero)
            {
                return dt;
            }
            int iCount = PWWrapper.aaApi_DmsDataBufferGetCount(iPtr);

            dt.Columns.Add(new DataColumn("ID", typeof(int)));
            dt.Columns.Add(new DataColumn("MemberType", typeof(string)));
            dt.Columns.Add(new DataColumn("MemberName", typeof(string)));

            for (int i = 0; i < iCount; i++)
            {

                int iMemberType = 1;
                int iMemberID = PWWrapper.aaApi_DmsDataBufferGetNumericProperty(iPtr, 2, i);

                string sMemberType = PWWrapper.GetUserListOrGroupMemberTypeName(iMemberType);
                string sMemberName = PWWrapper.GetUserListOrGroupMemberName(iMemberID, iMemberType);

                DataRow dr = dt.NewRow();

                dr["ID"] = iMemberID;
                dr["MemberType"] = sMemberType;
                dr["MemberName"] = sMemberName;

                dt.Rows.Add(dr);
            }

            PWWrapper.aaApi_DmsDataBufferFree(iPtr);

            return dt;
        }

        public ProjectWiseGroup(int iID, string sName)
        {
            ID = iID;
            Name = sName;
        }

        public ProjectWiseGroup(int iID, string sName, string sDescription, string sSecurityProvider,
            string sType)
        {
            ID = iID;
            Name = sName;
            Description = sDescription;
            SecurityProvider = sSecurityProvider;
            GroupType = sType;
        }

        public ProjectWiseGroup()
        {

        }
    }

    public static SortedList<string, PWWrapper.ProjectWiseApplication> GetApplicationsByName()
    {
        SortedList<string, ProjectWiseApplication> slApplications = new SortedList<string, ProjectWiseApplication>(StringComparer.InvariantCultureIgnoreCase);

        int iNumApplications = PWWrapper.aaApi_SelectAllApplications();

        for (int i = 0; i < iNumApplications; i++)
        {
            string sName = PWWrapper.aaApi_GetApplicationStringProperty(ApplicationProperty.Name, i);
            int iId = PWWrapper.aaApi_GetApplicationNumericProperty(ApplicationProperty.ID, i);
            //int iViewerId = PWWrapper.aaApi_GetApplicationNumericProperty(ApplicationProperty.ViewerId, i);

            if (!slApplications.ContainsKey(sName))
            {
                //    slApplications.Add(sName, new ProjectWiseApplication(iId, sName, iViewerId));
                slApplications.Add(sName, new ProjectWiseApplication(iId, sName));
            }
        }

        return slApplications;
    }

    public static SortedList<int, PWWrapper.ProjectWiseApplication> GetApplicationsById()
    {
        SortedList<int, ProjectWiseApplication> slApplications = new SortedList<int, ProjectWiseApplication>();

        int iNumApplications = PWWrapper.aaApi_SelectAllApplications();

        for (int i = 0; i < iNumApplications; i++)
        {
            int iId = PWWrapper.aaApi_GetApplicationNumericProperty(ApplicationProperty.ID, i);
            slApplications.AddWithCheck(iId, new ProjectWiseApplication(iId, PWWrapper.aaApi_GetApplicationStringProperty(ApplicationProperty.Name, i)));
        }

        return slApplications;
    }

    public static SortedList<string, PWWrapper.ProjectWiseUser> GetUsersByName()
    {
        SortedList<string, ProjectWiseUser> slUsers = new SortedList<string, ProjectWiseUser>(StringComparer.InvariantCultureIgnoreCase);

        int iNumUsers = PWWrapper.aaApi_SelectAllUsers();

        for (int i = 0; i < iNumUsers; i++)
        {
            string sName = PWWrapper.aaApi_GetUserStringProperty(UserProperty.Name, i);
            int iID = PWWrapper.aaApi_GetUserNumericProperty(UserProperty.ID, i);

            bool bDisabled = (1 == PWWrapper.aaApi_GetUserNumericProperty(UserProperty.Flags, i));

            if (!slUsers.ContainsKey(sName))
            {
                slUsers.Add(sName, new ProjectWiseUser(iID, sName,
                    PWWrapper.aaApi_GetUserStringProperty(UserProperty.Desc, i),
                    PWWrapper.aaApi_GetUserStringProperty(UserProperty.SecProvider, i),
                    PWWrapper.aaApi_GetUserStringProperty(UserProperty.Email, i),
                    PWWrapper.aaApi_GetUserStringProperty(UserProperty.Type, i),
                    bDisabled));
            }
        }

        return slUsers;
    }

    public static SortedList<string, PWWrapper.ProjectWiseUser> GetUsersByIdentity()
    {
        SortedList<string, PWWrapper.ProjectWiseUser> slUsers = new SortedList<string, PWWrapper.ProjectWiseUser>();

        int iNumUsers = PWWrapper.aaApi_SelectAllUsers();

        DataTable dt = PWWrapper.CreateDataTableFromSQLSelect("select o_userno, o_idpno, o_idname from dms_identity", "Identities");

        SortedList<int, string> slUserIdsToIdentities = new SortedList<int, string>();

        foreach (DataRow dr in dt.Rows)
        {
            if (!slUserIdsToIdentities.ContainsKey((int)dr["o_userno"]))
                slUserIdsToIdentities.Add((int)dr["o_userno"], (string)dr["o_idname"]);
        }

        for (int i = 0; i < iNumUsers; i++)
        {
            string sName = PWWrapper.aaApi_GetUserStringProperty(PWWrapper.UserProperty.Name, i);
            int iID = PWWrapper.aaApi_GetUserNumericProperty(PWWrapper.UserProperty.ID, i);
            bool bDisabled = (1 == PWWrapper.aaApi_GetUserNumericProperty(PWWrapper.UserProperty.Flags, i));

            if (slUserIdsToIdentities.ContainsKey(iID))
            {
                string sIdentity = slUserIdsToIdentities[iID];

                if (!string.IsNullOrEmpty(sIdentity))
                {
                    slUsers.AddWithCheck(sIdentity, new PWWrapper.ProjectWiseUser(iID, sName,
                                                        PWWrapper.aaApi_GetUserStringProperty(PWWrapper.UserProperty.Desc, i),
                                                        PWWrapper.aaApi_GetUserStringProperty(PWWrapper.UserProperty.SecProvider, i),
                                                        PWWrapper.aaApi_GetUserStringProperty(PWWrapper.UserProperty.Email, i),
                                                        PWWrapper.aaApi_GetUserStringProperty(PWWrapper.UserProperty.Type, i),
                                                        bDisabled)
                    {
                        Identity = sIdentity
                    });
                }
            }
        }

        return slUsers;
    }


    public static SortedList<string, PWWrapper.ProjectWiseUser> GetUsersByEmail()
    {
        SortedList<string, ProjectWiseUser> slUsers = new SortedList<string, ProjectWiseUser>(StringComparer.InvariantCultureIgnoreCase);

        int iNumUsers = PWWrapper.aaApi_SelectAllUsers();

        for (int i = 0; i < iNumUsers; i++)
        {
            string sName = PWWrapper.aaApi_GetUserStringProperty(UserProperty.Name, i);
            int iID = PWWrapper.aaApi_GetUserNumericProperty(UserProperty.ID, i);

            string sEmail = PWWrapper.aaApi_GetUserStringProperty(UserProperty.Email, i);

            bool bDisabled = (1 == PWWrapper.aaApi_GetUserNumericProperty(UserProperty.Flags, i));

            slUsers.AddWithCheck(sEmail, new ProjectWiseUser(iID, sName,
                PWWrapper.aaApi_GetUserStringProperty(UserProperty.Desc, i),
                PWWrapper.aaApi_GetUserStringProperty(UserProperty.SecProvider, i),
                PWWrapper.aaApi_GetUserStringProperty(UserProperty.Email, i),
                PWWrapper.aaApi_GetUserStringProperty(UserProperty.Type, i),
                bDisabled));
        }

        return slUsers;
    }


    public static SortedList<string, PWWrapper.ProjectWiseUser> GetUsersByDomainAndUsername()
    {
        SortedList<string, ProjectWiseUser> slUsers = new SortedList<string, ProjectWiseUser>(StringComparer.InvariantCultureIgnoreCase);

        int iNumUsers = PWWrapper.aaApi_SelectAllUsers();

        for (int i = 0; i < iNumUsers; i++)
        {
            int iID = PWWrapper.aaApi_GetUserNumericProperty(UserProperty.ID, i);
            bool bDisabled = (1 == PWWrapper.aaApi_GetUserNumericProperty(UserProperty.Flags, i));
            string sDomain = PWWrapper.aaApi_GetUserStringProperty(UserProperty.SecProvider, i);
            string sName = PWWrapper.aaApi_GetUserStringProperty(UserProperty.Name, i);
            string sCombinedName = string.Empty;

            if (string.IsNullOrEmpty(sDomain))
                sCombinedName = sName;
            else
                sCombinedName = string.Format(@"{0}\{1}", sDomain, sName);

            if (!slUsers.ContainsKey(sCombinedName))
            {
                slUsers.Add(sCombinedName, new ProjectWiseUser(iID, sName,
                    PWWrapper.aaApi_GetUserStringProperty(UserProperty.Desc, i),
                    PWWrapper.aaApi_GetUserStringProperty(UserProperty.SecProvider, i),
                    PWWrapper.aaApi_GetUserStringProperty(UserProperty.Email, i),
                    PWWrapper.aaApi_GetUserStringProperty(UserProperty.Type, i),
                    bDisabled));
            }
        }

        return slUsers;
    }

    public static bool IsUserIdInUserList(string sUserList, int iUserId)
    {
        int iNumUserLists = PWWrapper.aaApi_SelectAllUserLists();

        for (int i = 0; i < iNumUserLists; i++)
        {
            if (sUserList.ToLower() == PWWrapper.aaApi_GetUserListStringProperty(UserListProperty.Name, i).ToLower())
            {
                SortedList<int, int> slUserListsVisited = new SortedList<int, int>();

                if (IsUserIdInUserListId(PWWrapper.aaApi_GetUserListNumericProperty(UserListProperty.ID, i), 
                        iUserId, ref slUserListsVisited))
                    return true;

                return false;
            }
        }

        return false;
    }

    private static bool IsUserIdInUserListId(int iUserListID, int iUserId, ref SortedList<int, int> slUserLists)
    {
        // if already checked
        if (!slUserLists.AddWithCheck(iUserListID, iUserListID))
            return false;

        IntPtr userBuf = aaApi_SelectUserListMemberDataBufferByProp(iUserListID, (int)ManagerTypeProperty.User, iUserId);

        if (userBuf != IntPtr.Zero)
        {
            if (1 == PWWrapper.aaApi_DmsDataBufferGetCount(userBuf))
            {
                PWWrapper.aaApi_DmsDataBufferFree(userBuf);
                return true;
            }

            PWWrapper.aaApi_DmsDataBufferFree(userBuf);
        }

        IntPtr groupBuf = PWWrapper.aaApi_SelectUserListMemberDataBufferByProp(iUserListID,
            (int)ManagerTypeProperty.Group, -1);

        if (groupBuf != IntPtr.Zero)
        {
            for (int j = 0; j < PWWrapper.aaApi_DmsDataBufferGetCount(groupBuf); j++)
            {
                if (PWWrapper.aaApi_SelectGroupMembers(
                    PWWrapper.aaApi_DmsDataBufferGetNumericProperty(groupBuf, (int)UserListMemberProperty.MemberID, j), iUserId) > 0)
                {
                    PWWrapper.aaApi_DmsDataBufferFree(groupBuf);
                    return true;
                }
            }

            PWWrapper.aaApi_DmsDataBufferFree(groupBuf);
        }

        IntPtr userListBuf = PWWrapper.aaApi_SelectUserListMemberDataBufferByProp(iUserListID, 
            (int)ManagerTypeProperty.UserList, -1);

        if (userListBuf != IntPtr.Zero)
        {
            for (int k = 0; k < PWWrapper.aaApi_DmsDataBufferGetCount(userListBuf); k++)
            {
                if (IsUserIdInUserListId(
                    PWWrapper.aaApi_DmsDataBufferGetNumericProperty(userListBuf, (int)UserListMemberProperty.MemberID, k), 
                        iUserId, ref slUserLists))
                {
                    PWWrapper.aaApi_DmsDataBufferFree(userListBuf);
                    return true;
                }
            }

            PWWrapper.aaApi_DmsDataBufferFree(userListBuf);
        }

        return false;
    }


    public static SortedList<string, PWWrapper.ProjectWiseUserList> GetUserListsByName()
    {
        SortedList<string, PWWrapper.ProjectWiseUserList> slUserLists =
            new SortedList<string, PWWrapper.ProjectWiseUserList>(StringComparer.InvariantCultureIgnoreCase);

        int iNumUsers = PWWrapper.aaApi_SelectAllUserLists();

        for (int i = 0; i < iNumUsers; i++)
        {
            string sName = PWWrapper.aaApi_GetUserListStringProperty(UserListProperty.Name, i);
            int iID = PWWrapper.aaApi_GetUserListNumericProperty(UserListProperty.ID, i);

            if (!slUserLists.ContainsKey(sName))
            {
                slUserLists.Add(sName, new ProjectWiseUserList(iID, sName,
                    PWWrapper.aaApi_GetUserListStringProperty(UserListProperty.Description, i),
                    PWWrapper.aaApi_GetUserListNumericProperty(UserListProperty.Owner, i),
                    PWWrapper.aaApi_GetUserListNumericProperty(UserListProperty.Type, i)));
            }
        }

        return slUserLists;
    }

    public static SortedList<int, PWWrapper.ProjectWiseUserList> GetUserListsById()
    {
        SortedList<int, PWWrapper.ProjectWiseUserList> slUserLists =
            new SortedList<int, PWWrapper.ProjectWiseUserList>();

        int iNumUsers = PWWrapper.aaApi_SelectAllUserLists();

        for (int i = 0; i < iNumUsers; i++)
        {
            string sName = PWWrapper.aaApi_GetUserListStringProperty(UserListProperty.Name, i);
            int iID = PWWrapper.aaApi_GetUserListNumericProperty(UserListProperty.ID, i);

            if (!slUserLists.ContainsKey(iID))
            {
                slUserLists.Add(iID, new ProjectWiseUserList(iID, sName,
                    PWWrapper.aaApi_GetUserListStringProperty(UserListProperty.Description, i),
                    PWWrapper.aaApi_GetUserListNumericProperty(UserListProperty.Owner, i),
                    PWWrapper.aaApi_GetUserListNumericProperty(UserListProperty.Type, i)));
            }
        }

        return slUserLists;
    }

    public static SortedList<string, PWWrapper.ProjectWiseGroup> GetGroupsByName()
    {
        SortedList<string, PWWrapper.ProjectWiseGroup> slGroups =
            new SortedList<string, PWWrapper.ProjectWiseGroup>(StringComparer.InvariantCultureIgnoreCase);

        int iNumUsers = PWWrapper.aaApi_SelectAllGroups();

        for (int i = 0; i < iNumUsers; i++)
        {
            string sName = PWWrapper.aaApi_GetGroupStringProperty(GroupProperty.Name, i);
            int iID = PWWrapper.aaApi_GetGroupNumericProperty(GroupProperty.ID, i);

            if (!slGroups.ContainsKey(sName))
            {
                slGroups.Add(sName, new ProjectWiseGroup(iID, sName,
                    PWWrapper.aaApi_GetGroupStringProperty(GroupProperty.Desc, i),
                    PWWrapper.aaApi_GetGroupStringProperty(GroupProperty.SecProvider, i),
                    PWWrapper.aaApi_GetGroupStringProperty(GroupProperty.Type, i)));
            }
        }

        return slGroups;
    }

    public static SortedList<int, PWWrapper.ProjectWiseGroup> GetGroupsById()
    {
        SortedList<int, PWWrapper.ProjectWiseGroup> slGroups =
            new SortedList<int, PWWrapper.ProjectWiseGroup>();

        int iNumUsers = PWWrapper.aaApi_SelectAllGroups();

        for (int i = 0; i < iNumUsers; i++)
        {
            string sName = PWWrapper.aaApi_GetGroupStringProperty(GroupProperty.Name, i);
            int iID = PWWrapper.aaApi_GetGroupNumericProperty(GroupProperty.ID, i);

            if (!slGroups.ContainsKey(iID))
            {
                slGroups.Add(iID, new ProjectWiseGroup(iID, sName,
                    PWWrapper.aaApi_GetGroupStringProperty(GroupProperty.Desc, i),
                    PWWrapper.aaApi_GetGroupStringProperty(GroupProperty.SecProvider, i),
                    PWWrapper.aaApi_GetGroupStringProperty(GroupProperty.Type, i)));
            }
        }

        return slGroups;
    }

    public static SortedList<int, PWWrapper.ProjectWiseUser> GetUsersById()
    {
        SortedList<int, ProjectWiseUser> slUsers = new SortedList<int, ProjectWiseUser>();

        int iNumUsers = PWWrapper.aaApi_SelectAllUsers();

        for (int i = 0; i < iNumUsers; i++)
        {
            string sName = PWWrapper.aaApi_GetUserStringProperty(UserProperty.Name, i);
            int iID = PWWrapper.aaApi_GetUserNumericProperty(UserProperty.ID, i);
            bool bDisabled = (1 == PWWrapper.aaApi_GetUserNumericProperty(UserProperty.Flags, i));

            if (!slUsers.ContainsKey(iID))
            {
                slUsers.Add(iID, new ProjectWiseUser(iID, sName,
                    PWWrapper.aaApi_GetUserStringProperty(UserProperty.Desc, i),
                    PWWrapper.aaApi_GetUserStringProperty(UserProperty.SecProvider, i),
                    PWWrapper.aaApi_GetUserStringProperty(UserProperty.Email, i),
                    PWWrapper.aaApi_GetUserStringProperty(UserProperty.Type, i),
                    bDisabled));
            }
        }

        return slUsers;
    }


    public static Hashtable GetWorkflows()
    {
        Hashtable htWorkflows = new Hashtable();

        for (int i = 0; i < PWWrapper.aaApi_SelectAllWorkflows(); i++)
        {
            string sWfName = PWWrapper.aaApi_GetWorkflowStringProperty(
                PWWrapper.WorkflowProperty.Name, i);

            int iWfId = PWWrapper.aaApi_GetWorkflowId(i);

            htWorkflows.Add(sWfName.ToLower(), iWfId);
        }

        return htWorkflows;
    }

    public static SortedList<string, int> GetStoragesByName()
    {
        SortedList<string, int> listStoragesByName = new SortedList<string, int>(StringComparer.CurrentCultureIgnoreCase);

        int iNumStorages = PWWrapper.aaApi_SelectAllStorages();

        for (int i = 0; i < iNumStorages; i++)
        {
            if (!listStoragesByName.ContainsKey(PWWrapper.aaApi_GetStorageStringProperty(PWWrapper.StorageProperty.Name, i).ToLower()))
                listStoragesByName.Add(PWWrapper.aaApi_GetStorageStringProperty(PWWrapper.StorageProperty.Name, i).ToLower(), PWWrapper.aaApi_GetStorageId(i));
        }

        return listStoragesByName;
    }

    public static SortedList<int, string> GetStoragesById()
    {
        SortedList<int, string> listStoragesById = new SortedList<int, string>();

        int iNumStorages = PWWrapper.aaApi_SelectAllStorages();

        for (int i = 0; i < iNumStorages; i++)
        {
            if (!listStoragesById.ContainsKey(PWWrapper.aaApi_GetStorageId(i)))
                listStoragesById.Add(PWWrapper.aaApi_GetStorageId(i), PWWrapper.aaApi_GetStorageStringProperty(PWWrapper.StorageProperty.Name, i));
        }

        return listStoragesById;
    }


    public static SortedList<int, string> GetWorkflowsById()
    {
        SortedList<int, string> listWorkflowsById = new SortedList<int, string>();

        int iNumWorkflows = PWWrapper.aaApi_SelectAllWorkflows();

        for (int i = 0; i < iNumWorkflows; i++)
        {
            if (!listWorkflowsById.ContainsKey(PWWrapper.aaApi_GetWorkflowId(i)))
                listWorkflowsById.Add(PWWrapper.aaApi_GetWorkflowId(i), PWWrapper.aaApi_GetWorkflowStringProperty(PWWrapper.WorkflowProperty.Name, i));
        }

        return listWorkflowsById;
    }

    public static SortedList<int, string> GetStatesById()
    {
        SortedList<int, string> listStatesById = new SortedList<int, string>();

        int iNumStates = PWWrapper.aaApi_SelectAllStates();

        for (int i = 0; i < iNumStates; i++)
        {
            if (!listStatesById.ContainsKey(PWWrapper.aaApi_GetStateId(i)))
                listStatesById.Add(PWWrapper.aaApi_GetStateId(i), PWWrapper.aaApi_GetStateStringProperty(StateProperty.Name, i));
        }

        return listStatesById;
    }

    public static SortedList<int, string> GetEnvironmentsById()
    {
        SortedList<int, string> listEnvironmentsById = new SortedList<int, string>();

        int iNumEnvs = PWWrapper.aaApi_SelectAllEnvs(true);

        for (int i = 0; i < iNumEnvs; i++)
        {
            if (!listEnvironmentsById.ContainsKey(PWWrapper.aaApi_GetEnvId(i)))
                listEnvironmentsById.Add(PWWrapper.aaApi_GetEnvId(i), PWWrapper.aaApi_GetEnvStringProperty(EnvironmentProperty.Name, i));
        }

        return listEnvironmentsById;
    }

    public static SortedList<string, int> GetWorkflowsByName()
    {
        SortedList<string, int> listWorkflowsByName = new SortedList<string, int>(StringComparer.CurrentCultureIgnoreCase);

        int iNumWorkflows = PWWrapper.aaApi_SelectAllWorkflows();

        for (int i = 0; i < iNumWorkflows; i++)
        {
            if (!listWorkflowsByName.ContainsKey(PWWrapper.aaApi_GetWorkflowStringProperty(PWWrapper.WorkflowProperty.Name, i)))
                listWorkflowsByName.Add(PWWrapper.aaApi_GetWorkflowStringProperty(PWWrapper.WorkflowProperty.Name, i), PWWrapper.aaApi_GetWorkflowId(i));
        }

        return listWorkflowsByName;
    }

    public static SortedList<string, int> GetStatesByName()
    {
        SortedList<string, int> listStatesByName = new SortedList<string, int>(StringComparer.CurrentCultureIgnoreCase);

        int iNumStates = PWWrapper.aaApi_SelectAllStates();

        for (int i = 0; i < iNumStates; i++)
        {
            if (!listStatesByName.ContainsKey(PWWrapper.aaApi_GetStateStringProperty(StateProperty.Name, i)))
                listStatesByName.Add(PWWrapper.aaApi_GetStateStringProperty(StateProperty.Name, i), PWWrapper.aaApi_GetStateId(i));
        }

        return listStatesByName;
    }

    public static SortedList<string, int> GetEnvironmentsByName()
    {
        SortedList<string, int> listEnvironmentsByName = new SortedList<string, int>(StringComparer.CurrentCultureIgnoreCase);

        int iNumEnvs = PWWrapper.aaApi_SelectAllEnvs(true);

        for (int i = 0; i < iNumEnvs; i++)
        {
            if (!listEnvironmentsByName.ContainsKey(PWWrapper.aaApi_GetEnvStringProperty(EnvironmentProperty.Name, i)))
                listEnvironmentsByName.Add(PWWrapper.aaApi_GetEnvStringProperty(EnvironmentProperty.Name, i), PWWrapper.aaApi_GetEnvId(i));
        }

        return listEnvironmentsByName;
    }


    public static string GetWorkflowName(int iWorkflowId)
    {
        if (iWorkflowId > 0)
            if (1 == PWWrapper.aaApi_SelectWorkflow(iWorkflowId))
                return PWWrapper.aaApi_GetWorkflowStringProperty(WorkflowProperty.Name, 0);
        return string.Empty;
    }

    public static string GetStateName(int iStateId)
    {
        if (iStateId > 0)
            if (1 == PWWrapper.aaApi_SelectState(iStateId))
                return PWWrapper.aaApi_GetStateStringProperty(StateProperty.Name, 0);

        return string.Empty;
    }

    public static int GetStateId(string sStateName)
    {
        if (string.IsNullOrEmpty(sStateName))
            return 0;

        for (int i = 0; i < PWWrapper.aaApi_SelectAllStates(); i++)
        {
            string sCurrStateName = PWWrapper.aaApi_GetStateStringProperty(StateProperty.Name, i);

            if (sCurrStateName.ToLower() == sStateName.ToLower())
                return PWWrapper.aaApi_GetStateId(i);
        }

        return 0;
    }

    public static Hashtable GetWorkflowStates(string sWorkflowName)
    {
        Hashtable htWfStates = new Hashtable();

        if (!string.IsNullOrEmpty(sWorkflowName))
        {
            int iWfId = GetWorkflowId(sWorkflowName);

            if (iWfId > 0)
            {
                int iNumStates = PWWrapper.aaApi_SelectStatesByWorkflow(iWfId);

                for (int i = 0; i < iNumStates; i++)
                {
                    string sStateName = PWWrapper.aaApi_GetStateStringProperty(StateProperty.Name, i);
                    int iStateId = PWWrapper.aaApi_GetStateId(i);

                    htWfStates.Add(sStateName.ToLower(), iStateId);
                }
            }
        }
        else
        {
            int iNumStates = PWWrapper.aaApi_SelectAllStates();

            for (int i = 0; i < iNumStates; i++)
            {
                string sStateName = PWWrapper.aaApi_GetStateStringProperty(StateProperty.Name, i);
                int iStateId = PWWrapper.aaApi_GetStateId(i);

                htWfStates.Add(sStateName.ToLower(), iStateId);
            }
        }

        return htWfStates;
    }

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_OdbcDateTimeStringToDbDateTimeString
        (
           String lpctstrTimeToFmt, /* i  time string in string format    */
           StringBuilder lptstrBuffer,     /* o  Buffer to receive the data type */
           int lSize             /* i  Size of lptstrBuffer (in chars) */
        );


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_EnableUserAccount(int iUserId, bool bEnable);

    public static int GetApplicationNumberForApplicationName(string sAppName)
    {
        int iApplications = PWWrapper.aaApi_SelectAllApplications();

        for (int i = 0; i < iApplications; i++)
        {
            if (sAppName.ToLower() ==
                (PWWrapper.aaApi_GetApplicationStringProperty(PWWrapper.ApplicationProperty.Name, i)).ToLower())
                return PWWrapper.aaApi_GetApplicationId(i);
        }

        return 0;
    }


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_DeleteProjectById(int iProjectId);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_DeleteProject
        (
        int lProjectId,     //LONG  lProjectId,  
        uint ulFlags,       //ULONG  ulFlags,  
        IntPtr fpCallBack,  //AAPROC_PROJECTDELETE  fpCallBack,  
        IntPtr aaUserParm,  //AAPARAM  aaUserParam,  
        ref int lplCount    //LPLONG  lplCount   
        );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_GetInternalDatasourceName
        (
        IntPtr hDataSource,           /* i  Datasource handle                     */
        StringBuilder lptstrDsName,          /* o  Internal datasource name              */
        int iBufferSize           /* i  lptstrDsName size in TCHARs           */
        );


    //[DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    //public static extern bool aaOApi_ExportOdsSchemaAsECXML
    //(
    //ref int classIds,                 // i - Array of ids for class to be exported
    //int nClassIds,                // i - Class id array size
    //ref int pTreeDefIdArray,          // i - Array of ids for tree definitions to be exported
    //int lTreeDefIdArrayLen,       // i - Tree definitions id array size
    //ref int pCustFolderIdArray,       // i - Array of ids for custom folders to be exported
    //int lCustFolderIdArrayLen,    // i - Custom folder id array size
    //bool bExportAllApplications,   // i - if TRUE, export all applications
    //bool bExportAllFunctions,      // i - if TRUE, export all functions
    //bool bExportAllRules,          // i - if TRUE, export all rules
    //bool bSuppressAccessRights,    // i - if TRUE, suppress exporting access rights
    //string ecprefix,                 // i - XML prefix to use in ECXML export file
    //string ecschemauri,              // i - URI to specify in ECXML export file
    //string ecSchemaXmlFile,          // i - name of ECXML export file
    //bool bNoValidate               // i - if TRUE, skip schema validation
    //);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaOApi_ImportOdsSchemaAsECXML
    (
    string ecSchemaXmlFile,    // i - ECSchema XML file to import
    bool bPreview,           // i - if TRUE, don't apply import - just rollback after test
    bool bNoValidate         // i - if TRUE, skip schema validation
    );


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_GUIDChangeDocumentFile
        (
        ref Guid pDocGuid,
        string newDocFileName
        );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_DeleteApplicationAction
        (
        int iApplId,
        int iUserId,
        ref Guid pActionTypeGuid,
        string sProgramName
        );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_GUIDModifyDocument
        (
        ref Guid pDocGuid,
        int fileType,
        int itemType,
        int applicationId,
        int departmentId,
        int workspaceProfileId,
        string docFileName,
        string docName,
        string docDesc
        );


    public static int GetApplicationId(string sApplicationName)
    {
        if (!string.IsNullOrEmpty(sApplicationName))
        {
            int iNumApplications = PWWrapper.aaApi_SelectAllApplications();

            for (int i = 0; i < iNumApplications; i++)
            {
                string sName =
                    PWWrapper.aaApi_GetApplicationStringProperty(ApplicationProperty.Name, i);

                if (sApplicationName.ToLower() == sName.ToLower())
                    return PWWrapper.aaApi_GetApplicationId(i);
            }
        }

        return 0;
    }

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_GUIDRefreshDocumentServerCopy(ref Guid guid);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GetProjectIdByNamePath
    (
       string lpctstrPath         /* i  Project name path  */
    );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_SetDocumentState
    (
       uint ulFlags,        /* i  Operation flags                    */
       int lProjectId,     /* i  Project number                     */
       int lDocumentId,    /* i  Document number                    */
       int lWorkflowId,    /* i  Workflow number                    */
       int lStateId,        /* i  State id (0 - next, -1 - prev)     */
       string lpctstrComment  /* i  Operation comments for audit trail */
    );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_ChangeDocumentToNextState(int lProjectNo,
        int lDocumentId,
        int lWorkflowId
    );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_ChangeDocumentToPrevState(int lProjectNo,
        int lDocumentId,
        int lWorkflowId
    );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_DeleteSet
    (
        int iSetProjectId,
        int iSetItemId,
        int iSetId
    );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_CreateSet
    (
       int lProjectId,         /* i  Project number for set         */
       ref int lpDocumentId,       /* o  Document number of set item    */
       ref int lpSetId,            /* o  Set number of created set      */
       int lSetType,           /* i  Type for set 2 = flat, 3 = hierachical */
       int lDepartmentId,      /* i  Department number for set      */
       string lpctstrName,        /* i  Set name                       */
       string lpctstrDesc,        /* i  Set description                */
       int lParentProjectId,   /* i  Parent project number of set   */
       int lParentDocumentId,  /* i  Parent document number of set  */
       int lChildProjectId,    /* i  Child project number of set    */
       int lChildDocumentId,   /* i  Child document number of set   */
       int lRelationType,      /* i  Set relation type 2=Group, 3=Redline, 4=Ref */
       string lpctstrTransfer,    /* i  Transfer type for set member "C" Copy out, "CO" checkout  */
       string sGuid,         /* i  Document GUID (for flat set)   */
       ref int lpMemberId          /* o  Member number of child in set  */
    );


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_AddSetMember
    (
       int lSetId,             /* i  Set number                    */
       int lSetType,           /* i  Type of set                   */
       int lParentProjectId,   /* i  Parent project number         */
       int lParentDocumentId,  /* i  Parent document number        */
       int lChildProjectId,    /* i  Child project number          */
       int lChildDocumentId,   /* i  Child document number         */
       int lRelationType,      /* i  Relation type for set member  */
       string lpctstrTransfer,    /* i  Transfer type for set member  */
       ref int lplMemberId         /* o  Member number of child in set */
    );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_DeleteSetMember
    (
       int lSetId,             /* i  Set number                    */
       int lMemberId,          /* i  ignored if 0                   */
       int lParentProjectId,   /* i  Parent project number         */
       int lParentDocumentId,  /* i  Parent document number        */
       int lChildProjectId,    /* i  Child project number          */
       int lChildDocumentId
    );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_SetProjectWorkflow
    (
       int lProjectId,   /* i  Project number                      */
       int lWorkflowId   /* i  Workflow number to set (0 to reset) */
    );


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GUIDSelectDocumentsByProjectId(ref Guid guid);


    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GetDocumentId(int lIndex);



    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaApi_GuidListCreate();

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_GuidListAddGuid(IntPtr guidListP, ref Guid guidP);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaApi_GUIDSelectNestedReferencesList(ref Guid masterGuidP,
        int iflags,    // ignored
        int imaxDepth   // -1 for all
        );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_GuidListDestroy(IntPtr guidListP);

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_GuidListGetSize(IntPtr guidListP);


    [DllImport("dmscli.dll", EntryPoint = "aaApi_GuidListGetAt", CharSet = CharSet.Unicode)]
    private static extern IntPtr __aaApi_GuidListGetAt(IntPtr guidListP, int iIndex);

    public static Guid aaApi_GuidListGetAt(IntPtr guidListP, int iIndex)
    {
        return (Guid)Marshal.PtrToStructure(__aaApi_GuidListGetAt(guidListP, iIndex), Type.GetType("System.Guid"));
    }

    [DllImport("dmscli.dll", EntryPoint = "aaApi_GuidListGetFirstGuid", CharSet = CharSet.Unicode)]
    private static extern IntPtr __aaApi_GuidListGetFirstGuid(IntPtr guidListP);

    public static Guid aaApi_GuidListGetFirstGuid(IntPtr guidListP)
    {
        return (Guid)Marshal.PtrToStructure(__aaApi_GuidListGetFirstGuid(guidListP), Type.GetType("System.Guid"));
    }

    [DllImport("dmscli.dll", EntryPoint = "aaApi_GuidListGetNextGuid", CharSet = CharSet.Unicode)]
    private static extern IntPtr __aaApi_GuidListGetNextGuid(IntPtr guidListP);

    public static Guid aaApi_GuidListGetNextGuid(IntPtr guidListP)
    {
        return (Guid)Marshal.PtrToStructure(__aaApi_GuidListGetNextGuid(guidListP), Type.GetType("System.Guid"));
    }

    [DllImport("dmscli.dll", CharSet = CharSet.Ansi)]
    public static extern int aaApi_GetGuidsFromFileName([In, Out] ref IntPtr docGuids, [In, Out]ref int iNumGuids,
        [In] string sFileName, [In] int iValidateWithChkl);

    #region DataTableFromSQLSelect
    // Oracle
    //Name                                      Null?    Type
    //----------------------------------------- -------- -------------Name       Native Type
    //A_BIGINT                                           NUMBER(20)   A_BIGINT	    2	24	8       A_BIGINT	23	24	8
    //A_CHAR                                             CHAR(25)     A_CHAR	    9	12	52      A_CHAR	    9	12	52
    //A_DOUBLE                                           FLOAT(126)   A_DOUBLE	    5	7	8       A_DOUBLE	5	7	8
    //A_GUID                                             CHAR(36)     A_GUID	    9	12	74      A_GUID	    21	21	16
    //AN_INTEGER                                         NUMBER(10)   AN_INTEGER	2	3	4       AN_INTEGER	3	3	4
    //A_TIMESTAMP                                        DATE         A_TIMESTAMP	17	12	50      A_TIMESTAMP	17	12	50
    //A_VARCHAR                                          VARCHAR2(25) A_VARCHAR	    10	12	52      A_VARCHAR	10	12	52
    //A_VARWCHAR                                         NVARCHAR2(25)A_VARWCHAR	13	12	52      A_VARWCHAR	13	12	52
    //A_WCHAR                                            NCHAR(25)    A_WCHAR	    12	12	52      A_WCHAR	    12	12	52
    //A_DATE                                             DATE         A_DATE	    17	12	50
    //A_ATTRNO                                           NUMBER(10)   A_ATTRNO	    2	3	4
    //A_VERSION                                          NUMBER(10)   A_VERSION	    2	3	4
    //O_PROJECTNO                                        NUMBER(10)   O_PROJECTNO	2	3	4
    //O_ITEMNO                                           NUMBER(10)   O_ITEMNO	    2	3	4
    //A_CREATORNO                                        NUMBER(10)   A_CREATORNO	2	3	4
    //A_CREATETIME                                       DATE         A_CREATETIME	17	12	50
    //A_UPDATORNO                                        NUMBER(10)   A_UPDATORNO	2	3	4
    //A_UPDATETIME                                       DATE         A_UPDATETIME	17	12	50

    //SQL Server
    //A_BIGINT	    23	24	8
    //A_CHAR	    9	12	52
    //A_DOUBLE	    5	7	8
    //A_GUID	    21	21	16
    //AN_INTEGER	3	3	4
    //A_TIMESTAMP	17	12	50
    //A_VARCHAR	    10	12	52
    //A_VARWCHAR	13	12	52
    //A_WCHAR	    12	12	52
    //A_BIT	        22	4	2
    //A_DATETIME	17	12	50
    //A_DECIMAL	    2	7	8
    //A_FLOAT	    5	7	8
    //A_REAL	    6	6	4
    //A_SMALLINT	4	4	2
    //A_TINYINT	    26	4	2
    //A_UNIQUEIDENTIFIER	21	21	16
    //a_attrno	    3	3	4
    //a_version	    3	3	4
    //o_projectno	3	3	4
    //o_itemno	    3	3	4
    //a_creatorno	3	3	4
    //a_createtime	17	12	50
    //a_updatorno	3	3	4
    //a_updatetime	17	12	50

    /// <summary>
    /// SQL Data Types - AASQL_*
    /// </summary>
    public enum SQLSelectPWTypes : int
    {
        Long = 24,
        String = 12,
        Double = 7,
        Guid = 21,
        Short = 4,
        Integer = 3,
        SQLReal = 6
    }

    /// <summary>
    /// See SQL Data Types - AASQL_*
    /// Updated by dww for import tools.
    /// </summary>
    public enum SQLSelectDBColumnTypes : int
    {
        SQLUnknown = 0,
        SQLNumeric = 1,
        OracleNumber = 2,   // not in SDK docs
        SQLDecimal = 2,
        SQLInteger = 3,
        SQLSmallInt = 4,
        SQLFloat = 5,
        SQLReal = 6,
        SQLDouble = 7,
        SQLDateTime = 8,
        Char = 9,           // not in SDK docs
        SQLChar = 9,
        VarChar = 10,       // not in SDK docs
        SQLVarChar = 10,
        SQLLongVarChar = 11,
        WChar = 12,         // not in SDK docs
        SQLWChar = 12,
        VarWChar = 13,      // not in SDK docs
        SQLVarWChar = 13,
        SQLLongVarWChar = 14,
        SQLDate = 15,
        SQLTime = 16,
        DateTime = 17,      // not in SDK docs
        SQLTimeStamp = 17,
        SQLBinary = 18,
        SQLVarBinary = 19,
        SQLLongVarBinary = 20,
        SQLGuid = 21,
        SQLBoolean = 22,    // not in SDK docs
        SQLBit = 22,
        SQLBigInt = 23,
        SQL_C_SBigInt = 24,
        SQL_C_UBigInt = 25,
        SQLTinyInt = 26,
        SQL_C_TimeSpan = 27
    }

    private static DataColumn GetDataColumn(string sName, SQLSelectDBColumnTypes iNativeType, SQLSelectPWTypes iPWType, int iLength)
    {
        if (iNativeType == SQLSelectDBColumnTypes.DateTime)
            return new DataColumn(sName, Type.GetType("System.DateTime"));
        if (iNativeType == SQLSelectDBColumnTypes.SQLBoolean)
            return new DataColumn(sName, Type.GetType("System.Boolean"));
        if (iNativeType == SQLSelectDBColumnTypes.SQLGuid)
            return new DataColumn(sName, Type.GetType("System.Guid"));


        if (iPWType == SQLSelectPWTypes.String)
            return new DataColumn(sName, Type.GetType("System.String"));
        if (iPWType == SQLSelectPWTypes.Integer)
            return new DataColumn(sName, Type.GetType("System.Int32"));
        if (iPWType == SQLSelectPWTypes.Double)
            return new DataColumn(sName, Type.GetType("System.Double"));
        if (iPWType == SQLSelectPWTypes.Long)
            return new DataColumn(sName, Type.GetType("System.Int64"));
        if (iPWType == SQLSelectPWTypes.Short)
            return new DataColumn(sName, Type.GetType("System.Int32"));
        if (iPWType == SQLSelectPWTypes.SQLReal)
            return new DataColumn(sName, Type.GetType("System.Double"));

        System.Diagnostics.Debug.WriteLine(string.Format("No mapping found for '{0}' type {1} nativetype {2}",
            sName, iNativeType, iPWType));

        return new DataColumn(sName, Type.GetType("System.String"));
    }

    public static DataTable CreateDataTableFromSQLSelect(string sSQL, string sTableName)
    {
        int iCols = 0;

        // System.Diagnostics.Debug.WriteLine(string.Format("Executing: '{0}'...", sSQL));

        int iNumRows = PWWrapper.aaApi_SqlSelect(sSQL, IntPtr.Zero, ref iCols);

        // System.Diagnostics.Debug.WriteLine(string.Format("'{2}' returned {0} rows and {1} columns", iNumRows, iCols, sSQL));

        DataTable dt = new DataTable();

        if (!string.IsNullOrEmpty(sTableName))
            dt.TableName = sTableName;

        if (iCols > 0)
        {
            // System.Diagnostics.Debug.WriteLine("Building data table...");

            for (int j = 0; j < iCols; j++)
            {
                //System.Diagnostics.Debug.WriteLine(string.Format("{0}\t{1}\t{2}\t{3}",
                //    PWWrapper.aaApi_SqlSelectGetStringProperty(PWWrapper.SqlSelectProperties.SQLSELECT_COLUMN_NAME, j),
                //    PWWrapper.aaApi_SqlSelectGetNumericProperty(PWWrapper.SqlSelectProperties.SQLSELECT_COLUMN_NATIVE_TYPE, j),
                //    PWWrapper.aaApi_SqlSelectGetNumericProperty(PWWrapper.SqlSelectProperties.SQLSELECT_COLUMN_TYPE, j),
                //    PWWrapper.aaApi_SqlSelectGetNumericProperty(PWWrapper.SqlSelectProperties.SQLSELECT_COLUMN_LENGTH, j)));

                dt.Columns.Add(GetDataColumn(PWWrapper.aaApi_SqlSelectGetStringProperty(PWWrapper.SqlSelectProperties.SQLSELECT_COLUMN_NAME, j),
                    (SQLSelectDBColumnTypes)PWWrapper.aaApi_SqlSelectGetNumericProperty(PWWrapper.SqlSelectProperties.SQLSELECT_COLUMN_NATIVE_TYPE, j),
                    (SQLSelectPWTypes)PWWrapper.aaApi_SqlSelectGetNumericProperty(PWWrapper.SqlSelectProperties.SQLSELECT_COLUMN_TYPE, j),
                    PWWrapper.aaApi_SqlSelectGetNumericProperty(PWWrapper.SqlSelectProperties.SQLSELECT_COLUMN_LENGTH, j)));
            }

            // if (iNumRows == 0)
            // System.Diagnostics.Debug.WriteLine("No rows returned");

            for (int i = 0; i < iNumRows; i++)
            {
                DataRow dr = dt.NewRow();

                for (int j = 0; j < iCols; j++)
                {
                    string sValue = PWWrapper.aaApi_SqlSelectGetData(i, j);

                    try
                    {
                        if (dt.Columns[j].DataType == Type.GetType("System.String"))
                            dr[j] = sValue;
                        else if (dt.Columns[j].DataType == Type.GetType("System.DateTime"))
                        {
                            DateTime date = DateTime.Now;

                            if (DateTime.TryParse(sValue, out date))
                            {
                                dr[j] = date;
                            }
                            else
                            {
                                System.Diagnostics.Debug.WriteLine(string.Format("Error parsing '{0}' into date for column {1}",
                                    sValue, dt.Columns[j].ColumnName));
                            }
                        }
                        else if (dt.Columns[j].DataType == Type.GetType("System.Boolean"))
                        {
                            bool bValue = false;
                            string sValue2 = sValue.ToLower();

                            if (bool.TryParse(sValue2, out bValue))
                                dr[j] = bValue;
                            else if (sValue2 == "no" || sValue2 == "false" || sValue2 == "0")
                                dr[j] = false;
                            else if (sValue2 == "yes" || sValue2 == "true" || sValue2 == "1")
                                dr[j] = true;
                            else
                            {
                                System.Diagnostics.Debug.WriteLine(string.Format("Error parsing '{0}' into boolean for column {1}",
                                    sValue, dt.Columns[j].ColumnName));
                            }
                        }
                        else if (dt.Columns[j].DataType == Type.GetType("System.Guid"))
                        {
                            try
                            {
                                Guid guid = new Guid(sValue);
                                dr[j] = guid;
                            }
                            catch
                            {
                                System.Diagnostics.Debug.WriteLine(string.Format("Error parsing '{0}' into guid for column {1}",
                                    sValue, dt.Columns[j].ColumnName));
                            }
                        }
                        else if (dt.Columns[j].DataType == Type.GetType("System.Int32"))
                        {
                            int iValue = 0;

                            if (int.TryParse(sValue, out iValue))
                                dr[j] = iValue;
                            else
                            {
                                System.Diagnostics.Debug.WriteLine(string.Format("Error parsing '{0}' into int for column {1}",
                                    sValue, dt.Columns[j].ColumnName));
                            }
                        }
                        else if (dt.Columns[j].DataType == Type.GetType("System.Double"))
                        {
                            double dValue = 0;

                            if (double.TryParse(sValue, out dValue))
                                dr[j] = dValue;
                            else
                            {
                                System.Diagnostics.Debug.WriteLine(string.Format("Error parsing '{0}' into double for column {1}",
                                    sValue, dt.Columns[j].ColumnName));
                            }
                        }
                        else if (dt.Columns[j].DataType == Type.GetType("System.Int64"))
                        {
                            Int64 iValue64 = 0;

                            if (Int64.TryParse(sValue, out iValue64))
                                dr[j] = iValue64;
                            else
                            {
                                System.Diagnostics.Debug.WriteLine(string.Format("Error parsing '{0}' into int64 value for column {1}",
                                    sValue, dt.Columns[j].ColumnName));
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine(string.Format("Error parsing data for row {0}, column '{1}', value '{2}'", i + 1, dt.Columns[j].ColumnName, sValue));
                        System.Diagnostics.Debug.WriteLine(ex.Message);
                        System.Diagnostics.Debug.WriteLine(ex.StackTrace);
                    }
                } // for each column

                try
                {
                    dt.Rows.Add(dr);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(string.Format("Error adding row {0}", i + 1));
                    System.Diagnostics.Debug.WriteLine(ex.Message);
                    System.Diagnostics.Debug.WriteLine(ex.StackTrace);
                }
            } // for each row

            // System.Diagnostics.Debug.WriteLine("Data table built.");
        }
        else
        {
            System.Diagnostics.Debug.WriteLine("No columns returned");
        }

        return dt;
    }
    #endregion

    private static void AppendToEnvironmentPath(string path)
    {
        // I based the maximum length of the path value on the documentation of 
        // SetEnvironmentVariable
        int maxLength = 32767;
#if true
        // this is my v8i support hack
        if (ConfigurationManager.AppSettings["Path"] != null)
        {
            Environment.SetEnvironmentVariable("Path", ConfigurationManager.AppSettings["Path"], EnvironmentVariableTarget.Process);
        }
        else
#endif
        {
            StringBuilder currentPathBuffer = new StringBuilder(maxLength);

            uint length = GetEnvironmentVariable("Path", currentPathBuffer, (uint)maxLength);
            string newPath;
            if (length > 0)
            {
                if (currentPathBuffer.ToString().IndexOf(path) != -1)
                    return;

                // newPath = currentPathBuffer.ToString() + ";" + path;
                newPath = path + ";" + currentPathBuffer.ToString();
                if (newPath.Length >= maxLength)
                    throw new ApplicationException("Can not add to 'Path' environment variable because the resulting value would be too long.");
            }
            else
                newPath = path;

            bool success = SetEnvironmentVariable("Path", newPath);
            if (!success)
                throw new ApplicationException("Could not write to 'Path' environment variable.");
        }
    }

    public static bool Is64Bit()
    {
        return (IntPtr.Size == 8);
    }

    public static string GetProjectWisePath()
    {
        string installDirectory = null;

        try
        {
            if (Is64Bit())
            {
                /********  DESIRED VERSION - only need to change this line to update the *********
                 ********  PW verison that will be required by all components of RaDS    
                 ********  this is a minimum version requirement *********/
                 // first version with 64-bit
                Version minVersion = new Version("08.11");
                string[] regKeys = new String[]{
                    "SOFTWARE\\Bentley\\ProjectWise Explorer",
                    // added for integration server only
                    "SOFTWARE\\Bentley\\ProjectWise"
                };

                foreach (string regKeyPath in regKeys)
                {
                    Microsoft.Win32.RegistryKey regLocalMachine = Microsoft.Win32.Registry.LocalMachine;
                    Microsoft.Win32.RegistryKey regKey = regLocalMachine.OpenSubKey(regKeyPath);

                    if (regKey != null)
                    {
                        string[] versions = regKey.GetSubKeyNames();
                        if (versions != null)
                        {
                            for (int i = 0; i < versions.Length; i++)
                            {
                                Microsoft.Win32.RegistryKey versionSubKey = regKey.OpenSubKey(versions[i]);

                                if (versionSubKey == null)
                                    continue;

                                //string sVersion = (string)versionSubKey.GetValue("Version");
                                //if (sVersion == null)
                                //    continue;
                                Version version = new Version(versions[i]);

                                if (version >= minVersion)
                                {
                                    installDirectory = (string)versionSubKey.GetValue("PathName");

                                    if (string.IsNullOrEmpty(installDirectory))
                                    {
                                        installDirectory = (string)versionSubKey.GetValue("Path");
                                    }

                                    minVersion = version;
                                }
                            }
                            if (!string.IsNullOrEmpty(installDirectory))
                                break;
                        }
                    }
                }
                if (string.IsNullOrEmpty(installDirectory))
                    throw new ApplicationException("Registry search could not find installation directory for a ProjectWise version matching minimum required version '" +
                        minVersion + "'.\nMake sure a ProjectWise version matching the above version is installed on this system.");
            }
            else
            {

                /********  DESIRED VERSION - only need to change this line to update the *********
                 ********  PW verison that will be required by all components of RaDS    
                 ********  this is a minimum version requirement *********/
                Version minVersion = new Version("08.01");
                string[] regKeys = new String[]{"SOFTWARE\\Wow6432Node\\Bentley\\ProjectWise Explorer",
                                                    "SOFTWARE\\Wow6432Node\\Bentley\\ProjectWise Administrator",
                                                    "SOFTWARE\\Bentley\\ProjectWise Explorer",
                                                   "SOFTWARE\\Bentley\\ProjectWise Administrator"};

                foreach (string regKeyPath in regKeys)
                {
                    Microsoft.Win32.RegistryKey regLocalMachine = Microsoft.Win32.Registry.LocalMachine;
                    Microsoft.Win32.RegistryKey regKey = regLocalMachine.OpenSubKey(regKeyPath);

                    if (regKey != null)
                    {
                        string[] versions = regKey.GetSubKeyNames();
                        if (versions != null)
                        {
                            for (int i = 0; i < versions.Length; i++)
                            {
                                Microsoft.Win32.RegistryKey versionSubKey = regKey.OpenSubKey(versions[i]);
                                if (versionSubKey == null)
                                    continue;

                                string sVersion = (string)versionSubKey.GetValue("Version");
                                if (sVersion == null)
                                    continue;
                                Version version = new Version(sVersion);

                                if (version >= minVersion)
                                {
                                    installDirectory = (string)versionSubKey.GetValue("PathName");
                                    minVersion = version;
                                }
                            }
                            if (installDirectory != null)
                                break;
                        }
                    }
                }
                if (installDirectory == null)
                    throw new ApplicationException("Registry search could not find installation directory for a ProjectWise version matching minimum required version '" +
                        minVersion + "'.\nMake sure a ProjectWise version matching the above version is installed on this system.");
            }
        }
        catch (Exception ex)
        {
            // EventLog log = new EventLog("Application", ".", "ProjectWise .NET API Wrapper");
            // log.WriteEntry(ex.Message, EventLogEntryType.Error);
            System.Diagnostics.Debug.WriteLine(ex.Message);
        }

        return installDirectory;
    }

    private static void AppendProjectWiseDllPathToEnvironmentPath()
    {
        try
        {
            string installDirectory = GetProjectWisePath();
            if (installDirectory != null)
            {
                installDirectory += "\\bin";
                AppendToEnvironmentPath(installDirectory);
            }
        }
        catch (Exception ex)
        {
            // EventLog log = new EventLog("Application", ".", "ProjectWise .NET API Wrapper");
            // log.WriteEntry(ex.Message, EventLogEntryType.Error);
            System.Diagnostics.Debug.WriteLine(ex.Message);
        }
    }

    [DllImport("KERNEL32.dll")]
    private static extern bool SetEnvironmentVariable(string name, string val);

    [DllImport("KERNEL32.dll")]
    private static extern uint GetEnvironmentVariable(string name, StringBuilder valueBuffer, uint bufferSize);


#region UtilityFunctions

    public static bool SetNonStandardUserStringSettingByUser(int iUserId, int iParamNo, string sUserSetting)
    {
        aaApi_ExecuteSqlStatement("create view v_dms_ucfg as select * from dms_ucfg");

        string sSql = string.Format("delete from v_dms_ucfg where o_userno = {0} and o_paramno = {1}", iUserId, iParamNo);

        aaApi_ExecuteSqlStatement(sSql);

        sSql =
            string.Format("insert into v_dms_ucfg (o_userno, o_paramno, o_intval, o_textval, o_compguid) values ({0},{1},0,'{2}','00000000-0000-0000-0000-000000000000')",
                iUserId, iParamNo, sUserSetting);

        return aaApi_ExecuteSqlStatement(sSql);
    }

    public static string GetNonStandardUserStringSettingByUser(int iUserId, int iParamNo)
    {
        StringBuilder StringBuffer = new StringBuilder(280);

        string sSql =
            string.Format("select o_textval from v_dms_ucfg where o_userno = {0} and o_paramno = {1}",
                iUserId, iParamNo);

        int lNumCols = 0;

        if (0 < aaApi_SqlSelect(sSql, System.IntPtr.Zero, ref lNumCols))
        {
            string sSettingVal = aaApi_SqlSelectGetData(0, 0);

            StringBuffer.Append(sSettingVal);

            return StringBuffer.ToString();
        }

        return string.Empty;
    }

    public static bool SetNonStandardUserNumericSettingByUser(int iUserId, int iParamNo, int iUserSetting)
    {
        aaApi_ExecuteSqlStatement("create view v_dms_ucfg as select * from dms_ucfg");

        string sSql = string.Format("delete from v_dms_ucfg where o_userno = {0} and o_paramno = {1}", iUserId, iParamNo);

        aaApi_ExecuteSqlStatement(sSql);

        sSql =
            string.Format("insert into v_dms_ucfg (o_userno, o_paramno, o_intval, o_textval, o_compguid) values ({0},{1},{2},'','00000000-0000-0000-0000-000000000000')",
                iUserId, iParamNo, iUserSetting);

        return aaApi_ExecuteSqlStatement(sSql);
    }

    public static int GetNonStandardUserNumericSettingByUser(int iUserId, int iParamNo)
    {
        string sSql =
            string.Format("select o_intval from v_dms_ucfg where o_userno = {0} and o_paramno = {1}",
                iUserId, iParamNo);

        int lNumCols = 0;

        if (0 < aaApi_SqlSelect(sSql, System.IntPtr.Zero, ref lNumCols))
        {
            string sSettingVal = aaApi_SqlSelectGetData(0, 0);

            return int.Parse(sSettingVal);
        }

        return -1;
    }

    public static ArrayList GetBranchProjectNos(int iProjectNo, bool bGetSubProjects)
    {
        ArrayList alProjects = new ArrayList();

        if (bGetSubProjects && iProjectNo > 0)
        {
            int iNumProjs = PWWrapper.aaApi_SelectProjectsFromBranch(iProjectNo, null, null, null, null);

            for (int i = 0; i < iNumProjs; i++)
            {
                alProjects.Add(PWWrapper.aaApi_GetProjectNumericProperty(PWWrapper.ProjectProperty.ID, i));
            }
        }
        else if (iProjectNo <= 0)
        {
            int iNumProjs = PWWrapper.aaApi_SelectAllProjects();

            for (int i = 0; i < iNumProjs; i++)
            {
                alProjects.Add(PWWrapper.aaApi_GetProjectNumericProperty(PWWrapper.ProjectProperty.ID, i));
            }
        }
        else
        {
            alProjects.Add(iProjectNo);
        }

        return alProjects;
    }

    public static string GetAttributeColumnValue(int iProjectNo, int iDocumentNo,
        string sColumnName)
    {
        int iEnvId = 0, iTableId = 0, iColumnId = 0;

        string sRetVal = string.Empty;

        if (!string.IsNullOrEmpty(sColumnName))
        {

            if (PWWrapper.aaApi_GetEnvTableInfoByProject(iProjectNo,
                ref iEnvId, ref iTableId, ref iColumnId))
            {
                int lNumberOfColumns = 0;

                int lNumLinks = PWWrapper.aaApi_SelectLinkDataByObject(iTableId,
                    PWWrapper.ObjectTypeForLinkData.Document,
                    iProjectNo,
                    iDocumentNo,
                    null,
                    ref lNumberOfColumns,
                    null,
                    0);

                for (int iRow = 0; iRow < lNumLinks; iRow++)
                {
                    for (int iCol = 0; iCol < lNumberOfColumns; iCol++)
                    {
                        string sCurrColumnName =
                            PWWrapper.aaApi_GetLinkDataColumnStringProperty(PWWrapper.LinkDataProperty.ColumnName, iCol);

                        if (!string.IsNullOrEmpty(sCurrColumnName))
                        {
                            if (sColumnName.ToLower() == sCurrColumnName.ToLower())
                            {
                                sRetVal = PWWrapper.aaApi_GetLinkDataColumnValue(iRow, iCol);
                                break;
                            }
                        }
                    } // for each column
                } // for each link
            } // if environment selected
        }

        return sRetVal;
    }

    public static void GetAttributeColumnValues
        (
        int iProjectNo, // in - project id
        int iDocumentNo, // in - document id
        ref Hashtable htAttrVals // in/out hashtable with desired column names (lower case) as keys
        )
    {
        int iEnvId = 0, iTableId = 0, iColumnId = 0;

        if (PWWrapper.aaApi_GetEnvTableInfoByProject(iProjectNo,
            ref iEnvId, ref iTableId, ref iColumnId))
        {
            int lNumberOfColumns = 0;

            int lNumLinks = PWWrapper.aaApi_SelectLinkDataByObject(iTableId,
                PWWrapper.ObjectTypeForLinkData.Document,
                iProjectNo,
                iDocumentNo,
                null,
                ref lNumberOfColumns,
                null,
                0);

            for (int iRow = 0; iRow < lNumLinks; iRow++)
            {
                for (int iCol = 0; iCol < lNumberOfColumns; iCol++)
                {
                    string sCurrColumnName =
                        PWWrapper.aaApi_GetLinkDataColumnStringProperty(PWWrapper.LinkDataProperty.ColumnName, iCol);

                    if (htAttrVals.ContainsKey(sCurrColumnName.ToLower()))
                    {
                        htAttrVals[sCurrColumnName.ToLower()] = PWWrapper.aaApi_GetLinkDataColumnValue(iRow, iCol);
                    } // for each column
                } // for each link
            } // if environment selected
        }
    }

    public static Hashtable GetAttributeColumnNamesFromEnvironment(int iEnvId)
    {
        Hashtable htAttrVals = new Hashtable();

        if (1 == PWWrapper.aaApi_SelectEnv(iEnvId))
        {
            int iTableId = PWWrapper.aaApi_GetEnvNumericProperty(EnvironmentProperty.TableID, 0);

            if (iTableId > 0)
            {
                int iNumCols = PWWrapper.aaApi_SelectColumnsByTable(iTableId);

                for (int i = 0; i < iNumCols; i++)
                {
                    if (!htAttrVals.ContainsKey(PWWrapper.aaApi_GetColumnStringProperty(ColumnProperty.Name, i).ToLower()))
                        htAttrVals.Add(PWWrapper.aaApi_GetColumnStringProperty(ColumnProperty.Name, i).ToLower(),
                            PWWrapper.aaApi_GetColumnNumericProperty(ColumnProperty.ColumnID, i));
                }
            }
        }

        return htAttrVals;
    }

    public static SortedList<string, int> GetEnvironmentColumnsKeyedByName(int iEnvId)
    {
        SortedList<string, int> slColumns = new SortedList<string, int>(StringComparer.CurrentCultureIgnoreCase);

        if (1 == PWWrapper.aaApi_SelectEnv(iEnvId))
        {
            int iTableId = PWWrapper.aaApi_GetEnvNumericProperty(EnvironmentProperty.TableID, 0);

            if (iTableId > 0)
            {
                int iNumCols = PWWrapper.aaApi_SelectColumnsByTable(iTableId);

                for (int i = 0; i < iNumCols; i++)
                {
                    slColumns.AddWithCheck(PWWrapper.aaApi_GetColumnStringProperty(ColumnProperty.Name, i),
                        PWWrapper.aaApi_GetColumnNumericProperty(ColumnProperty.ColumnID, i));
                }

            }
        }

        return slColumns;
    }


    public static Hashtable GetProjectProperties(int iProjectId)
    {
        Hashtable htProps = new Hashtable();

        IntPtr hProjBuf = PWWrapper.aaApi_SelectRichProjectOfFolder(iProjectId);

        if (hProjBuf != IntPtr.Zero)
        {
            if (1 == PWWrapper.aaApi_DmsDataBufferGetCount(hProjBuf))
            {
                int iClassId = PWWrapper.aaApi_DmsDataBufferGetNumericProperty(hProjBuf,
                    (int)PWWrapper.ProjectProperty.ComponentClassId, 0);
                int iInstanceId = PWWrapper.aaApi_DmsDataBufferGetNumericProperty(hProjBuf,
                    (int)PWWrapper.ProjectProperty.ComponentInstanceId, 0);

                IntPtr instP = PWWrapper.aaOApi_LoadInstanceByIds(iClassId, iInstanceId, 0);

                if (instP != IntPtr.Zero)
                {
                    int iNumAttrs = 0;

                    if (PWWrapper.aaOApi_GetInstanceNumAttrs(instP, 0, false, ref iNumAttrs))
                    {
                        for (int i = 0; i < iNumAttrs; i++)
                        {
                            int iAttrId = 0, iAttrType = 0, iAttrAddOn = 0, iAttrParent = 0, iAttrVisibllity = 0;

                            if (PWWrapper.aaOApi_GetInstanceAttrId(instP, i, 0, false, ref iAttrId, ref iAttrType,
                                ref iAttrAddOn, ref iAttrParent, ref iAttrVisibllity))
                            {
                                IntPtr attrP = PWWrapper.aaOApi_FindAttributePtr(iAttrId);

                                StringBuilder sbAttrName = new StringBuilder(512);

                                if (PWWrapper.aaOApi_GetAttributeStringProperty(attrP, PWWrapper.ODSAttributeProperty.Name,
                                    sbAttrName, sbAttrName.Capacity))
                                {
                                    StringBuilder sbAttrVal = new StringBuilder(512);

                                    if (PWWrapper.aaOApi_GetInstanceAttrStrValue(instP, iAttrId, 0, sbAttrVal,
                                        sbAttrVal.Capacity))
                                    {
                                        htProps.Add(sbAttrName.ToString(), sbAttrVal.ToString());
                                    }
                                }
                            }
                        }
                    }

                    PWWrapper.aaOApi_FreeInstance(instP);
                }
            }

            PWWrapper.aaApi_DmsDataBufferFree(hProjBuf);
        }
        else
        {
            // WriteLog("Rich Project not found");
        }

        return htProps;
    }

    public static int GetRichProjectId(int iProjectId)
    {
        int iRichProjectId = 0;

        IntPtr hProjBuf = PWWrapper.aaApi_SelectRichProjectOfFolder(iProjectId);

        if (hProjBuf != IntPtr.Zero)
        {
            if (1 == PWWrapper.aaApi_DmsDataBufferGetCount(hProjBuf))
            {
                iRichProjectId = PWWrapper.aaApi_DmsDataBufferGetNumericProperty(hProjBuf,
                    (int)PWWrapper.ProjectProperty.ID, 0);
            }

            PWWrapper.aaApi_DmsDataBufferFree(hProjBuf);
        }

        return iRichProjectId;
    }

    public static SortedList<string, string> GetProjectPropertyValuesInList(int iProjectId)
    {
        SortedList<string, string> slPropertyNamesProperyValues =
            new SortedList<string, string>(StringComparer.InvariantCultureIgnoreCase);

        IntPtr hProjBuf = PWWrapper.aaApi_SelectRichProjectOfFolder(iProjectId);

        if (hProjBuf != IntPtr.Zero)
        {
            if (1 == PWWrapper.aaApi_DmsDataBufferGetCount(hProjBuf))
            {
                int iClassId = PWWrapper.aaApi_DmsDataBufferGetNumericProperty(hProjBuf,
                    (int)PWWrapper.ProjectProperty.ComponentClassId, 0);
                int iInstanceId = PWWrapper.aaApi_DmsDataBufferGetNumericProperty(hProjBuf,
                    (int)PWWrapper.ProjectProperty.ComponentInstanceId, 0);

                slPropertyNamesProperyValues = GetInstancePropertyValuesInList(iClassId, iInstanceId);
            }

            PWWrapper.aaApi_DmsDataBufferFree(hProjBuf);
        }
        else
        {
            // WriteLog("Rich Project not found");
        }

        return slPropertyNamesProperyValues;
    }

#region MANAGED_WORKSPACE
    public enum WorkspaceAssocObjectType : int
    {
        WORKSPACEOBJECTTYPE_INVALID = 0,
        WORKSPACEOBJECTTYPE_DATASOURCE = 1,
        WORKSPACEOBJECTTYPE_USER = 2,
        WORKSPACEOBJECTTYPE_PROJECT = 3,
        WORKSPACEOBJECTTYPE_APPLICATION = 4,
        WORKSPACEOBJECTTYPE_DOCUMENT = 5,
        WORKSPACEOBJECTTYPE_GROUP = 6,
        WORKSPACEOBJECTTYPE_USERLIST = 7
    }

    public enum WorkspaceDmsBufferProperty : int
    {
        MWPCONFBLOCK_NUMERIC_PROP_ID = 1,
        MWPCONFBLOCK_NUMERIC_PROP_FLAGS = 2,
        MWPCONFBLOCK_NUMERIC_PROP_LEVEL = 3,
        MWPCONFBLOCK_STRING_PROP_NAME = 4,
        MWPCONFBLOCK_STRING_PROP_DESC = 5
    }

    public enum WorkspaceDmsBufferVariablesProperty : int
    {
        WPCBSETTINGPROP_ID = 0,
        WPCBSETTINGPROP_DESCRIPTION = 1,
        WPCBSETTINGPROP_NAME = 2,
        WPCBSETTINGPROP_FLAGS = 3,
        WPCBSETTINGPROP_VALUES = 4
    }

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaApi_MwpSelectWorkspaceDataBuffer(WorkspaceAssocObjectType workspaceAssocObjectType,
        ref Guid objectGuid, int iObjectId, int iUserId, uint uiFlags);

    [DllImport("PWManagedWorkspace.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr workspace_selectConfBlockVars(int iConfigBlockId);

    [DllImport("PWManagedWorkspace.dll", CharSet = CharSet.Unicode)]
    public static extern int workspace_getItemCountInDataBuffer(IntPtr hWorkspaceBuffer);

    [DllImport("PWManagedWorkspace.dll", EntryPoint = "workspace_getDataBufferNumericProperty", CharSet = CharSet.Unicode)]
    public static extern int workspace_getDataBufferNumericProperty(IntPtr hWorkspaceBuffer, WorkspaceDmsBufferVariablesProperty propertyId, int iIndex);


    [DllImport("PWManagedWorkspace.dll", EntryPoint = "workspace_getDataBufferStringProperty", CharSet = CharSet.Unicode)]
    private static extern IntPtr __workspace_getDataBufferStringProperty(IntPtr hWorkspaceBuffer, int iPropertyId, int iIndex);

    public static string workspace_getDataBufferStringProperty(IntPtr hWorkspaceBuffer, WorkspaceDmsBufferVariablesProperty propertyId, int iIndex)
    {
        return Marshal.PtrToStringUni(__workspace_getDataBufferStringProperty(hWorkspaceBuffer, (int)propertyId, iIndex));
    }

#endregion

    public static SortedList<string, string> GetInstancePropertyValuesInList(int iClassId, int iInstanceId)
    {
        SortedList<string, string> slPropertyNamesProperyValues =
            new SortedList<string, string>(StringComparer.InvariantCultureIgnoreCase);

        IntPtr instP = PWWrapper.aaOApi_LoadInstanceByIds(iClassId, iInstanceId, 0);

        if (instP != IntPtr.Zero)
        {
            int iNumAttrs = 0;

            if (PWWrapper.aaOApi_GetInstanceNumAttrs(instP, 0, false, ref iNumAttrs))
            {
                for (int i = 0; i < iNumAttrs; i++)
                {
                    int iAttrId = 0, iAttrType = 0, iAttrAddOn = 0, iAttrParent = 0, iAttrVisibllity = 0;

                    if (PWWrapper.aaOApi_GetInstanceAttrId(instP, i, 0, false, ref iAttrId, ref iAttrType,
                        ref iAttrAddOn, ref iAttrParent, ref iAttrVisibllity))
                    {
                        IntPtr attrP = PWWrapper.aaOApi_FindAttributePtr(iAttrId);

                        StringBuilder sbAttrName = new StringBuilder(512);

                        if (PWWrapper.aaOApi_GetAttributeStringProperty(attrP, PWWrapper.ODSAttributeProperty.Name,
                            sbAttrName, sbAttrName.Capacity))
                        {
                            int iDataLength =
                                PWWrapper.aaOApi_GetAttributeNumericProperty(attrP, ODSAttributeProperty.DataLength);

                            if (iDataLength > 0)
                            {
                                StringBuilder sbAttrVal = new StringBuilder(iDataLength + 2);

                                if (PWWrapper.aaOApi_GetInstanceAttrStrValue(instP, iAttrId, 0, sbAttrVal,
                                    sbAttrVal.Capacity))
                                {
                                    if (!slPropertyNamesProperyValues.ContainsKey(sbAttrName.ToString()))
                                        slPropertyNamesProperyValues.Add(sbAttrName.ToString(), sbAttrVal.ToString());
                                }
                            }
                        }
                    }
                }
            }

            PWWrapper.aaOApi_FreeInstance(instP);
        }
        else
        {
            // WriteLog("Rich Project not found");
        }

        return slPropertyNamesProperyValues;
    }

    public static SortedList<string, string> GetInstancePropertyValuesInList(int iClassId, int iInstanceId, SortedList<string, int> slProperties)
    {
        SortedList<string, string> slPropertyNamesProperyValues =
            new SortedList<string, string>(StringComparer.InvariantCultureIgnoreCase);

        IntPtr instP = PWWrapper.aaOApi_LoadInstanceByIds(iClassId, iInstanceId, 0);

        if (instP != IntPtr.Zero)
        {
            foreach (KeyValuePair<string, int> kvp in slProperties)
            {
                StringBuilder sbAttrVal = new StringBuilder(512);

                if (PWWrapper.aaOApi_GetInstanceAttrStrValue(instP, kvp.Value, 0, sbAttrVal,
                    sbAttrVal.Capacity))
                {
                    if (!slPropertyNamesProperyValues.ContainsKey(kvp.Key))
                        slPropertyNamesProperyValues.Add(kvp.Key, sbAttrVal.ToString());
                }
            }

            PWWrapper.aaOApi_FreeInstance(instP);
        }
        else
        {
            // WriteLog("Rich Project not found");
        }

        return slPropertyNamesProperyValues;
    }

    public static SortedList<string, int> GetProjectPropertyIdsInList(int iProjectId)
    {
        SortedList<string, int> slPropertyNamesProperyIds =
            new SortedList<string, int>(StringComparer.InvariantCultureIgnoreCase);

        IntPtr hProjBuf = PWWrapper.aaApi_SelectRichProjectOfFolder(iProjectId);

        if (hProjBuf != IntPtr.Zero)
        {
            if (1 == PWWrapper.aaApi_DmsDataBufferGetCount(hProjBuf))
            {
                int iClassId = PWWrapper.aaApi_DmsDataBufferGetNumericProperty(hProjBuf,
                    (int)PWWrapper.ProjectProperty.ComponentClassId, 0);
                int iInstanceId = PWWrapper.aaApi_DmsDataBufferGetNumericProperty(hProjBuf,
                    (int)PWWrapper.ProjectProperty.ComponentInstanceId, 0);

                IntPtr instP = PWWrapper.aaOApi_LoadInstanceByIds(iClassId, iInstanceId, 0);

                if (instP != IntPtr.Zero)
                {
                    int iNumAttrs = 0;

                    if (PWWrapper.aaOApi_GetInstanceNumAttrs(instP, 0, false, ref iNumAttrs))
                    {
                        for (int i = 0; i < iNumAttrs; i++)
                        {
                            int iAttrId = 0, iAttrType = 0, iAttrAddOn = 0, iAttrParent = 0, iAttrVisibllity = 0;

                            if (PWWrapper.aaOApi_GetInstanceAttrId(instP, i, 0, false, ref iAttrId, ref iAttrType,
                                ref iAttrAddOn, ref iAttrParent, ref iAttrVisibllity))
                            {
                                IntPtr attrP = PWWrapper.aaOApi_FindAttributePtr(iAttrId);

                                StringBuilder sbAttrName = new StringBuilder(512);

                                if (PWWrapper.aaOApi_GetAttributeStringProperty(attrP, PWWrapper.ODSAttributeProperty.Name,
                                    sbAttrName, sbAttrName.Capacity))
                                {
                                    if (!slPropertyNamesProperyIds.ContainsKey(sbAttrName.ToString()))
                                        slPropertyNamesProperyIds.Add(sbAttrName.ToString(), iAttrId);
                                }
                            }
                        }
                    }

                    PWWrapper.aaOApi_FreeInstance(instP);
                }
            }

            PWWrapper.aaApi_DmsDataBufferFree(hProjBuf);
        }
        else
        {
            // WriteLog("Rich Project not found");
        }

        return slPropertyNamesProperyIds;
    }

    public static int GetClassIdFromClassName(string sClassName)
    {
        IntPtr pClass = PWWrapper.aaOApi_FindClassPtrByName(sClassName);
        if (pClass != IntPtr.Zero)
            return PWWrapper.aaOApi_GetClassId(pClass);
        return 0;
    }

    public static string GetClassNameFromClassId(int iClassId)
    {
        IntPtr pClass = PWWrapper.aaOApi_FindClassPtr(iClassId);

        if (pClass != IntPtr.Zero)
        {
            StringBuilder sbClassName = new StringBuilder(256);
            if (PWWrapper.aaOApi_GetClassStringProperty(pClass, PWWrapper.ODSClassProperty.Name,
                sbClassName, sbClassName.Capacity))
            {
                return sbClassName.ToString();
            }
        }

        return string.Empty;
    }

    public static SortedList<string, int> GetClassPropertyIdsInList(string sClassName)
    {
        return GetClassPropertyIdsInList(PWWrapper.GetClassIdFromClassName(sClassName));
    }

    public static SortedList<string, int> GetClassPropertyIdsInList(int iClassId)
    {
        SortedList<string, int> slPropertyNamesProperyIds =
            new SortedList<string, int>(StringComparer.InvariantCultureIgnoreCase);

        IntPtr pClass = PWWrapper.aaOApi_FindClassPtr(iClassId);

        if (pClass != null && pClass != IntPtr.Zero)
        {
            int lNumAttrs = PWWrapper.aaOApi_GetClassAttrCount(pClass, true);

            // lBusinessKeyAttrId = PWWrapper.aaOApi_ClassGetBusinessKeyAttrId(pClass);

            for (int i = 0; i < lNumAttrs; ++i)
            {
                int lAttrId = 0, lAddonId = 0, lParentId = 0;

                if (PWWrapper.aaOApi_GetClassAttributes(pClass, 0, i, ref lAttrId,
                    ref lAddonId, ref lParentId))
                {
                    IntPtr pAttr = IntPtr.Zero;

                    pAttr = PWWrapper.aaOApi_FindAttributePtr(lAttrId);

                    // find the attribute using the id
                    if (pAttr != IntPtr.Zero)
                    {
                        StringBuilder wAttrName = new StringBuilder(256);
                        StringBuilder wAttrDesc = new StringBuilder(256);

                        int lVisibility = 0, lValueType = 0, iDataLen = 0;

                        if (PWWrapper.aaOApi_GetAttributeCmnProps(pAttr, ref lAttrId, wAttrName, wAttrDesc,
                            ref lVisibility, ref lValueType, ref iDataLen))
                        {
                            if (!slPropertyNamesProperyIds.ContainsKey(wAttrName.ToString()))
                                slPropertyNamesProperyIds.Add(wAttrName.ToString(), lAttrId);
                        }
                    }
                }
            }
        }
        else
        {
            // WriteLog("Rich Project not found");
        }

        return slPropertyNamesProperyIds;
    }

    /// <summary>
    /// Maps Column Id to Column Name for an environment.
    /// </summary>
    /// <param name="iEnvId"></param>
    /// <returns>Hashtable with key of Column Id and value of Column Name.</returns>
    public static Hashtable GetAttributeColumnIdsFromEnvironment(int iEnvId)
    {
        Hashtable htAttrVals = new Hashtable();

        if (1 == PWWrapper.aaApi_SelectEnv(iEnvId))
        {
            int iAttrDefCount = PWWrapper.aaApi_SelectEnvAttrDefs(iEnvId, -1, -1);

            for (int i = 0; i < iAttrDefCount; i++)
            {
                int iTableId = PWWrapper.aaApi_GetEnvAttrDefNumericProperty(AttributeDefinitionProperty.TableID, i);
                int iColId = PWWrapper.aaApi_GetEnvAttrDefNumericProperty(AttributeDefinitionProperty.ColumnID, i);

                if (1 == PWWrapper.aaApi_SelectColumn(iTableId, iColId))
                {
                    string sColumnName = PWWrapper.aaApi_GetColumnStringProperty(ColumnProperty.Name, 0);

                    if (!htAttrVals.ContainsKey(iColId))
                        htAttrVals.Add(iColId, sColumnName.ToLower());
                }
            }
        }

        return htAttrVals;
    }

    public static string GetColumnTypeName(SQLSelectDBColumnTypes iNativeType, SQLSelectPWTypes iPWType)
    {
        if (iNativeType == SQLSelectDBColumnTypes.DateTime)
            return "datetime";
        if (iNativeType == SQLSelectDBColumnTypes.SQLBoolean)
            return "boolean";
        if (iNativeType == SQLSelectDBColumnTypes.SQLGuid)
            return "guid";


        if (iPWType == SQLSelectPWTypes.String)
            return "string";
        if (iPWType == SQLSelectPWTypes.Integer)
            return "integer";
        if (iPWType == SQLSelectPWTypes.Double)
            return "double";
        if (iPWType == SQLSelectPWTypes.Long)
            return "integer64";
        if (iPWType == SQLSelectPWTypes.Short)
            return "integer";
        if (iPWType == SQLSelectPWTypes.SQLReal)
            return "double";

        System.Diagnostics.Debug.WriteLine(string.Format("No mapping found for '{0}' type {1} nativetype {2}",
            "", iNativeType, iPWType));

        return "string";
    }

    public class PWColumn
    {
        public string Name { get; set; }
        public int ColumnId { get; set; }
        public int TableId { get; set; }
        public string TypeName { get; set; }
        public int Length { get; set; }
    }

    public static List<PWColumn> GetListOfColumnsByTableId(int iTableId)
    {
        int iAttrDefCount = PWWrapper.aaApi_SelectColumnsByTable(iTableId);

        List<PWColumn> listOfColumns = new List<PWColumn>();

        for (int i = 0; i < iAttrDefCount; i++)
        {
            PWColumn pwCol = new PWColumn();

            pwCol.ColumnId = PWWrapper.aaApi_GetColumnNumericProperty(ColumnProperty.ColumnID, i);

            pwCol.Name = PWWrapper.aaApi_GetColumnStringProperty(ColumnProperty.Name, i);

            pwCol.Length = PWWrapper.aaApi_GetColumnNumericProperty(ColumnProperty.Length, i);

            pwCol.TypeName = GetColumnTypeName((SQLSelectDBColumnTypes)PWWrapper.aaApi_GetColumnNumericProperty(ColumnProperty.SQLType, i),
                (SQLSelectPWTypes)PWWrapper.aaApi_GetColumnNumericProperty(ColumnProperty.Type, i));

            pwCol.TableId = iTableId;

            listOfColumns.Add(pwCol);
        }

        return listOfColumns;
    }


    public static SortedList<int, PWColumn> GetListOfColumnsByTableIdKeyedById(int iTableId)
    {
        int iAttrDefCount = PWWrapper.aaApi_SelectColumnsByTable(iTableId);

        SortedList<int, PWColumn> listOfColumns = new SortedList<int, PWColumn>();

        for (int i = 0; i < iAttrDefCount; i++)
        {
            PWColumn pwCol = new PWColumn();

            pwCol.ColumnId = PWWrapper.aaApi_GetColumnNumericProperty(ColumnProperty.ColumnID, i);

            pwCol.Name = PWWrapper.aaApi_GetColumnStringProperty(ColumnProperty.Name, i);

            pwCol.Length = PWWrapper.aaApi_GetColumnNumericProperty(ColumnProperty.Length, i);

            pwCol.TypeName = GetColumnTypeName((SQLSelectDBColumnTypes)PWWrapper.aaApi_GetColumnNumericProperty(ColumnProperty.SQLType, i),
                (SQLSelectPWTypes)PWWrapper.aaApi_GetColumnNumericProperty(ColumnProperty.Type, i));

            pwCol.TableId = iTableId;

            listOfColumns.Add(pwCol.ColumnId, pwCol);
        }

        return listOfColumns;
    }

    public static SortedList<string, PWColumn> GetListOfColumnsByTableIdKeyedByName(int iTableId)
    {
        int iAttrDefCount = PWWrapper.aaApi_SelectColumnsByTable(iTableId);

        SortedList<string, PWColumn> listOfColumns = new SortedList<string, PWColumn>(StringComparer.InvariantCultureIgnoreCase);

        for (int i = 0; i < iAttrDefCount; i++)
        {
            PWColumn pwCol = new PWColumn();

            pwCol.ColumnId = PWWrapper.aaApi_GetColumnNumericProperty(ColumnProperty.ColumnID, i);

            pwCol.Name = PWWrapper.aaApi_GetColumnStringProperty(ColumnProperty.Name, i);

            pwCol.Length = PWWrapper.aaApi_GetColumnNumericProperty(ColumnProperty.Length, i);

            pwCol.TypeName = GetColumnTypeName((SQLSelectDBColumnTypes)PWWrapper.aaApi_GetColumnNumericProperty(ColumnProperty.SQLType, i),
                (SQLSelectPWTypes)PWWrapper.aaApi_GetColumnNumericProperty(ColumnProperty.Type, i));

            pwCol.TableId = iTableId;

            listOfColumns.Add(pwCol.Name, pwCol);
        }

        return listOfColumns;
    }


    public static Hashtable GetColumnIdsKeyedByNameFromTable(int iTableId)
    {
        Hashtable htAttrVals = new Hashtable();

        int iAttrDefCount = PWWrapper.aaApi_SelectColumnsByTable(iTableId);

        for (int i = 0; i < iAttrDefCount; i++)
        {
            int iColId = PWWrapper.aaApi_GetColumnNumericProperty(ColumnProperty.ColumnID, i);

            string sColumnName = PWWrapper.aaApi_GetColumnStringProperty(ColumnProperty.Name, i);

            if (!htAttrVals.ContainsKey(sColumnName.ToLower()))
                htAttrVals.Add(sColumnName.ToLower(), iColId);
        }

        return htAttrVals;
    }

    public static Hashtable GetColumnNamesKeyedByIdFromTable(int iTableId)
    {
        Hashtable htAttrVals = new Hashtable();

        int iAttrDefCount = PWWrapper.aaApi_SelectColumnsByTable(iTableId);

        for (int i = 0; i < iAttrDefCount; i++)
        {
            int iColId = PWWrapper.aaApi_GetColumnNumericProperty(ColumnProperty.ColumnID, i);

            string sColumnName = PWWrapper.aaApi_GetColumnStringProperty(ColumnProperty.Name, i);

            if (!htAttrVals.ContainsKey(iColId))
                htAttrVals.Add(iColId, sColumnName.ToLower());
        }

        return htAttrVals;
    }

    /// <summary>
    /// returns hashtable with column names (lower case) as keys and attribute values as values
    /// </summary>
    /// <param name="iProjectNo"></param>
    /// <param name="iDocumentNo"></param>
    /// <returns></returns>
    public static Hashtable GetAllAttributeColumnValues
        (
        int iProjectNo, // in - project id
        int iDocumentNo // in - document id
        )
    {
        int iEnvId = 0, iTableId = 0, iColumnId = 0;

        Hashtable htAttrVals = new Hashtable();

        if (PWWrapper.aaApi_GetEnvTableInfoByProject(iProjectNo,
            ref iEnvId, ref iTableId, ref iColumnId))
        {
            int lNumberOfColumns = 0;

            int lNumLinks = PWWrapper.aaApi_SelectLinkDataByObject(iTableId,
                PWWrapper.ObjectTypeForLinkData.Document,
                iProjectNo,
                iDocumentNo,
                null,
                ref lNumberOfColumns,
                null,
                0);

            for (int iRow = 0; iRow < lNumLinks; iRow++)
            {
                for (int iCol = 0; iCol < lNumberOfColumns; iCol++)
                {
                    string sCurrColumnName =
                        PWWrapper.aaApi_GetLinkDataColumnStringProperty(PWWrapper.LinkDataProperty.ColumnName, iCol);

                    try
                    {
                        htAttrVals.Add(sCurrColumnName.ToLower(), PWWrapper.aaApi_GetLinkDataColumnValue(iRow, iCol));
                    }
                    finally
                    {
                    }
                } // for each link
            } // if environment selected
        }

        return htAttrVals;
    }

    /// <summary>
    /// returns sorted list with column names (lower case) as keys and attribute values as values
    /// </summary>
    /// <param name="iProjectNo"></param>
    /// <param name="iDocumentNo"></param>
    /// <returns></returns>
    public static SortedList<string, string> GetAllAttributeColumnValuesInList
        (
        int iProjectNo, // in - project id
        int iDocumentNo // in - document id
        )
    {
        int iEnvId = 0, iTableId = 0, iColumnId = 0;

        SortedList<string, string> slAttrValues = new SortedList<string, string>(StringComparer.InvariantCultureIgnoreCase);

        // Hashtable slAttrValues = new Hashtable();

        if (PWWrapper.aaApi_GetEnvTableInfoByProject(iProjectNo,
            ref iEnvId, ref iTableId, ref iColumnId))
        {
            int lNumberOfColumns = 0;

            int lNumLinks = PWWrapper.aaApi_SelectLinkDataByObject(iTableId,
                PWWrapper.ObjectTypeForLinkData.Document,
                iProjectNo,
                iDocumentNo,
                null,
                ref lNumberOfColumns,
                null,
                0);

            if (lNumLinks > 0)
            {
                for (int iCol = 0; iCol < lNumberOfColumns; iCol++)
                {
                    string sCurrColumnName =
                        PWWrapper.aaApi_GetLinkDataColumnStringProperty(PWWrapper.LinkDataProperty.ColumnName, iCol);

                    if (!slAttrValues.ContainsKey(sCurrColumnName))
                        slAttrValues.Add(sCurrColumnName, PWWrapper.aaApi_GetLinkDataColumnValue(0, iCol));

                    //try
                    //{
                    //    slAttrValues.Add(sCurrColumnName.ToLower(), PWWrapper.aaApi_GetLinkDataColumnValue(iRow, iCol));
                    //}
                    //finally
                    //{
                    //}
                } // for each link
            } // if environment selected
        }

        return slAttrValues;
    }

    /// <summary>
    /// returns list of sorted lists with column names (lower case) as keys and attribute values as values
    /// </summary>
    /// <param name="iProjectNo"></param>
    /// <param name="iDocumentNo"></param>
    /// <returns></returns>
    public static List<SortedList<string, string>> GetAllAttributesAndAllColumnValuesInList
        (
        int iProjectNo, // in - project id
        int iDocumentNo // in - document id
        )
    {
        int iEnvId = 0, iTableId = 0, iColumnId = 0;

        List<SortedList<string, string>> listAttributes = new List<SortedList<string, string>>();

        // Hashtable slAttrValues = new Hashtable();

        if (PWWrapper.aaApi_GetEnvTableInfoByProject(iProjectNo,
            ref iEnvId, ref iTableId, ref iColumnId))
        {
            int lNumberOfColumns = 0;

            int lNumLinks = PWWrapper.aaApi_SelectLinkDataByObject(iTableId,
                PWWrapper.ObjectTypeForLinkData.Document,
                iProjectNo,
                iDocumentNo,
                null,
                ref lNumberOfColumns,
                null,
                0);

            if (lNumLinks > 0)
            {
                for (int iRow = 0; iRow < lNumLinks; iRow++)
                {
                    SortedList<string, string> slAttrValues = new SortedList<string, string>(StringComparer.InvariantCultureIgnoreCase);

                    for (int iCol = 0; iCol < lNumberOfColumns; iCol++)
                    {
                        string sCurrColumnName =
                            PWWrapper.aaApi_GetLinkDataColumnStringProperty(PWWrapper.LinkDataProperty.ColumnName, iCol);

                        if (!slAttrValues.ContainsKey(sCurrColumnName))
                            slAttrValues.Add(sCurrColumnName, PWWrapper.aaApi_GetLinkDataColumnValue(0, iCol));

                        //try
                        //{
                        //    slAttrValues.Add(sCurrColumnName.ToLower(), PWWrapper.aaApi_GetLinkDataColumnValue(iRow, iCol));
                        //}
                        //finally
                        //{
                        //}
                    } // for each attribute

                    listAttributes.Add(slAttrValues);
                } // for each link
            } // if environment selected
        }

        return listAttributes;
    }

    /// <summary>
    /// returns list of sorted lists with column names (lower case) as keys and attribute values as values
    /// </summary>
    /// <param name="docGuid"></param>
    /// <returns></returns>
    public static List<SortedList<string, string>> GetAllAttributesAndAllColumnValuesInListFromGuid
        (
        Guid docGuid
        )
    {
        int iProjectNo = 0, iDocumentNo = 0;

        PWWrapper.GetIdsFromGuidString(docGuid.ToString(), ref iProjectNo, ref iDocumentNo);

        return GetAllAttributesAndAllColumnValuesInList(iProjectNo, iDocumentNo);
    }

    /// <summary>
    /// returns list of sorted lists with column names (lower case) as keys and attribute values as values
    /// </summary>
    /// <param name="sDocGuid"></param>
    /// <returns></returns>
    public static List<SortedList<string, string>> GetAllAttributesAndAllColumnValuesInListFromGuidString
        (
        string sDocGuid
        )
    {
        int iProjectNo = 0, iDocumentNo = 0;

        PWWrapper.GetIdsFromGuidString(sDocGuid, ref iProjectNo, ref iDocumentNo);

        return GetAllAttributesAndAllColumnValuesInList(iProjectNo, iDocumentNo);
    }

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_DocumentGenNameWithPrefix
    (
    int lProjectId,
    string lpctstrPrefix,
    [MarshalAs(UnmanagedType.LPWStr)]
        StringBuilder lptstrDocName,
    int iBufferSize
    );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_DocumentGenFileNameWithPrefix
    (
    int lProjectId,
    string lpctstrPrefix,
    [MarshalAs(UnmanagedType.LPWStr)]
        StringBuilder lptstrDocName,
    int iBufferSize
    );

    public static string GetUniqueDocumentName(int iProjectId, string sDocumentName)
    {
        StringBuilder sbDocName = new StringBuilder(127);

        if (aaApi_DocumentGenNameWithPrefix(iProjectId, sDocumentName, sbDocName, sbDocName.Capacity))
        {
            return sbDocName.ToString();
        }

        return sDocumentName;
    }

    public static string GetUniqueFileName(int iProjectId, string sFileName)
    {
        StringBuilder sbDocName = new StringBuilder(127);

        if (aaApi_DocumentGenFileNameWithPrefix(iProjectId, sFileName, sbDocName, sbDocName.Capacity))
        {
            return sbDocName.ToString();
        }

        return sFileName;
    }

    public static string GetDocumentNamePath(int iProjectId, int iDocId)
    {
        StringBuilder sbPathBuffer = new StringBuilder(5096);

        if (aaApi_GetDocumentNamePath(iProjectId, iDocId, false, '\\',
            sbPathBuffer, sbPathBuffer.Capacity))
        {
            return sbPathBuffer.ToString();
        }

        return string.Empty;
    }

    // dww - this uses the depreciated function aaApi_GetProjectNamePath
    // consider using GetProjectNamePath2 instead
    public static string GetProjectNamePath(int iProjectId)
    {
        StringBuilder sbPathBuffer = new StringBuilder(5096);

        if (aaApi_GetProjectNamePath(iProjectId, false, '\\',
            sbPathBuffer, sbPathBuffer.Capacity))
        {
            return sbPathBuffer.ToString();
        }

        return string.Empty;
    }

    public static string GetProjectNamePath2(int iProjectId)
    {
        StringBuilder sbPathBuffer = new StringBuilder(5096);

        if (aaApi_GetProjectNamePath2(iProjectId, false, '\\',
            sbPathBuffer, sbPathBuffer.Capacity))
        {
            return sbPathBuffer.ToString();
        }

        return string.Empty;
    }

    public static string GetDocumentNamePath(Guid docGuid)
    {
        StringBuilder sbDocPath = new StringBuilder(5096);

        if (PWWrapper.aaApi_GUIDGetDocumentNamePath(ref docGuid, false, '\\', sbDocPath, sbDocPath.Capacity))
            return sbDocPath.ToString();

        return string.Empty;
    }

    public static string GetDocumentURL(int iProjectId, int iDocId, string sDSName)
    {
        return string.Format("pw://{0}/Documents/D{1}",
          sDSName, "{" + GetGuidStringFromIds(iProjectId, iDocId) + "}");
    }

    public static string GetDocumentURL(Guid docGuid, string sDSName)
    {
        return string.Format("pw://{0}/Documents/D{1}",
          sDSName, "{" + docGuid.ToString() + "}");
    }

    public static string GetDocumentURL(int iProjectId, int iDocId)
    {
        StringBuilder sbDSN = new StringBuilder(1024);

        if (PWWrapper.aaApi_GetActiveDatasourceName(sbDSN, 1024))
        {
            return string.Format("pw://{0}/Documents/D{1}",
              sbDSN.ToString(), "{" + GetGuidStringFromIds(iProjectId, iDocId) + "}");
        }

        return string.Empty;
    }

    public static string GetDocumentURL(Guid docGuid)
    {
        StringBuilder sbDSN = new StringBuilder(1024);

        if (PWWrapper.aaApi_GetActiveDatasourceName(sbDSN, 1024))
        {
            return string.Format("pw://{0}/Documents/D{1}",
              sbDSN.ToString(), "{" + docGuid.ToString() + "}");
        }

        return string.Empty;
    }

    public static string GetURLEncodedDocumentMoniker(int iProjectId, int iDocumentId)
    {
        return System.Web.HttpUtility.UrlEncode(GetMonikerStringFromDocumentIds(iProjectId, iDocumentId));
    }

    // http://localhost/default.aspx?location=BRUMDA2066REM.bentley.com%3APWv8i&link=pw%3A%2F%2FBRUMDA2066REM.bentley.com%3APWv8i%2FDocuments%2FMicroStation%26space%3BJ%2Fbikefrme.dgn&action=DM_Open_ReadOnly

    // http://rcorppwsvr01:8089/default.aspx?location=RCORPPWSVR01.bentley.com%3Arcorppworcl&link=pw%3A%2F%2FRCORPPWSVR01.bentley.com%3Arcorppworcl%2FDocuments%2FDave%26space%3BTest%2FTarget%26space%3BFolders%2FTarget%26space%3B1%2F

    public enum WebLinkActions : int
    {
        None,
        Open,
        OpenReadOnly,
        View,
        Markup
    }

    public static string GetURNWebLink(string sWebAddressIncludingASPXName, int iProjectId, WebLinkActions action)
    {
        string sMoniker = GetMonikerStringFromProjectId(iProjectId);

        return GetURNWebLink(sWebAddressIncludingASPXName, sMoniker, action);
    }

    public static string GetURNWebLink(string sWebAddressIncludingASPXName, Guid docGuid, WebLinkActions action)
    {
        string sMoniker = GetMonikerStringFromDocumentGuid(docGuid);

        return GetURNWebLink(sWebAddressIncludingASPXName, sMoniker, action);
    }

    public static string GetURNWebLink(string sWebAddressIncludingASPXName, int iProjectId, int iDocumentId, WebLinkActions action)
    {
        string sMoniker = GetMonikerStringFromDocumentIds(iProjectId, iDocumentId);

        return GetURNWebLink(sWebAddressIncludingASPXName, sMoniker, action);
    }

    public static string GetURNWebLink(string sWebAddressIncludingASPXName, string sMoniker, WebLinkActions action)
    {
        string sFormat = "{0}?location={1}&link={2}";

        // string sDatasource = GetDatasourceNameFromMonikerString(sMoniker);

        StringBuilder sbDSN = new StringBuilder(1024);

        if (!PWWrapper.aaApi_GetActiveDatasourceName(sbDSN, 1024))
        {
            return string.Empty;
        }

        string sDatasource = sbDSN.ToString();

        string sEncodedDSN = System.Web.HttpUtility.UrlEncode(sDatasource);

        string sEncodedPath = System.Web.HttpUtility.UrlEncode(sMoniker);

        string sURL = string.Format(sFormat, sWebAddressIncludingASPXName, sEncodedDSN, sEncodedPath);

        string sSuffix = string.Empty;

        switch (action)
        {
            case WebLinkActions.None:
                break;
            case WebLinkActions.Open:
                sSuffix = "&action=DM_Open";
                break;
            case WebLinkActions.OpenReadOnly:
                sSuffix = "&action=DM_Open_ReadOnly";
                break;
            case WebLinkActions.View:
                sSuffix = "&action=DM_View";
                break;
            case WebLinkActions.Markup:
                sSuffix = "&action=DM_Redline";
                break;
            default:
                break;
        }

        return sURL + sSuffix;
    }

    public static string GetFormattedWebLink(string sWebAddressIncludingASPXName, string sDatasource, string sDocumentNamePath, WebLinkActions action)
    {
        string sFormat = "{0}?location={1}&link={2}";

        string sEncodedDSN = System.Web.HttpUtility.UrlEncode(sDatasource);

        string sEncodedPath = System.Web.HttpUtility.UrlEncode(string.Format("pw://{0}/Documents/{1}", sDatasource, sDocumentNamePath));

        string sURL = string.Format(sFormat, sWebAddressIncludingASPXName, sEncodedDSN, sEncodedPath);

        string sSuffix = string.Empty;

        switch (action)
        {
            case WebLinkActions.None:
                break;
            case WebLinkActions.Open:
                sSuffix = "&action=DM_Open";
                break;
            case WebLinkActions.OpenReadOnly:
                sSuffix = "&action=DM_Open_ReadOnly";
                break;
            case WebLinkActions.View:
                sSuffix = "&action=DM_View";
                break;
            case WebLinkActions.Markup:
                sSuffix = "&action=DM_Redline";
                break;
            default:
                break;
        }

        return sURL + sSuffix;
    }

    public static string GetProjectWebLink(string sWebAddressIncludingASPXName, int iProjectId)
    {
        StringBuilder sbDSN = new StringBuilder(1024);

        if (PWWrapper.aaApi_GetActiveDatasourceName(sbDSN, 1024))
        {
            return GetFormattedWebLink(sWebAddressIncludingASPXName, sbDSN.ToString(),
                GetProjectNamePath(iProjectId) + "//", WebLinkActions.None);
        }

        return string.Empty;
    }


    public static string GetDocumentWebLink(string sWebAddressIncludingASPXName, Guid docGuid, WebLinkActions action)
    {
        StringBuilder sbDSN = new StringBuilder(1024);

        if (PWWrapper.aaApi_GetActiveDatasourceName(sbDSN, 1024))
        {
            return GetFormattedWebLink(sWebAddressIncludingASPXName, sbDSN.ToString(), GetDocumentNamePath(docGuid), action);
        }

        return string.Empty;
    }

    public static string GetDocumentWebLink(string sWebAddressIncludingASPXName, int iProjectId, int iDocumentId, WebLinkActions action)
    {
        StringBuilder sbDSN = new StringBuilder(1024);

        if (PWWrapper.aaApi_GetActiveDatasourceName(sbDSN, 1024))
        {
            return GetFormattedWebLink(sWebAddressIncludingASPXName, sbDSN.ToString(), GetDocumentNamePath(iProjectId, iDocumentId), action);
        }

        return string.Empty;
    }

    public static string GetDocumentWebLink(string sWebAddressIncludingASPXName, string sDatasource, Guid docGuid, WebLinkActions action)
    {
        return GetFormattedWebLink(sWebAddressIncludingASPXName, sDatasource, GetDocumentNamePath(docGuid), action);
    }

    public static string GetDocumentWebLink(string sWebAddressIncludingASPXName, string sDatasource, int iProjectId, int iDocumentId, WebLinkActions action)
    {
        return GetFormattedWebLink(sWebAddressIncludingASPXName, sDatasource, GetDocumentNamePath(iProjectId, iDocumentId), action);
    }

    public static string GetProjectURL(int iProjectId)
    {
        Guid[] guids = new Guid[1];

        if (PWWrapper.aaApi_GetProjectGUIDsByIds(1, ref iProjectId, guids))
        {
            StringBuilder sbDSN = new StringBuilder(1024);

            if (PWWrapper.aaApi_GetActiveDatasourceName(sbDSN, 1024))
            {
                return string.Format("pw://{0}/Documents/P{1}/",
                    sbDSN.ToString(), "{" + guids[0].ToString() + "}");
            }
        }

        return string.Empty;
    }

    public static string GetGuidStringFromIds(int iProjectId, int iDocId)
    {
        PWWrapper.AaDocItem[] documents = new PWWrapper.AaDocItem[1];
        Guid[] guids = new Guid[1];

        documents[0].lProjectId = iProjectId;
        documents[0].lDocumentId = iDocId;

        if (PWWrapper.aaApi_GetDocumentGUIDsByIds(1, documents, guids))
            return guids[0].ToString();

        return guids[0].ToString();
    }

    public static string GetProjectGuidStringFromId(int iProjectId)
    {
        int[] projIds = new int[1];
        Guid[] guids = new Guid[1];

        if (aaApi_GetProjectGUIDsByIds(1, ref iProjectId, guids))
        {
            return guids[0].ToString();
        }

        return string.Empty;
    }

    public static bool GetIdsFromGuidString(string sDocGuid, ref int iProjId, ref int iDocId)
    {
        PWWrapper.AaDocItem docItem = new PWWrapper.AaDocItem();
        Guid[] guids = new Guid[1];

        try
        {
            guids[0] = new Guid(sDocGuid);

            if (PWWrapper.aaApi_GetDocumentIdsByGUIDs(1, guids, ref docItem))
            {
                iProjId = docItem.lProjectId;
                iDocId = docItem.lDocumentId;
                return true;
            }
        }
        finally
        {
        }

        return false;
    }

    public static bool SetAttributesValues
    (
        int iProjectNo, // in - project id
        int iDocumentNo, // in - document id
        Dictionary<string, string> dictPossiblyCaseSensitive, // in diction with column names as keys and new values as values
        bool bAddSheet
    )
    {
        int iEnvId = 0, iTableId = 0, iColumnId = 0;

        SortedList<string, string> slCaseInsensitive = new SortedList<string, string>(StringComparer.CurrentCultureIgnoreCase);

        foreach (KeyValuePair<string, string> kvp in dictPossiblyCaseSensitive)
        {
            if (!slCaseInsensitive.ContainsKey(kvp.Key))
                slCaseInsensitive.Add(kvp.Key, kvp.Value);
        }

        if (PWWrapper.aaApi_GetEnvTableInfoByProject(iProjectNo,
            ref iEnvId, ref iTableId, ref iColumnId))
        {
            int iNumLinks = PWWrapper.aaApi_SelectLinks(iProjectNo, iDocumentNo);

            if (iNumLinks == 0 || bAddSheet)
            {
                PWWrapper.aaApi_FreeLinkDataInsertDesc();

                bool bUpdatedValue = false;

                int iNumAttrs = PWWrapper.aaApi_SelectColumnsByTable(iTableId);

                SortedList<string, int> slColumnNamesToIds = new SortedList<string, int>(StringComparer.CurrentCultureIgnoreCase);

                for (int iCol = 0; iCol < iNumAttrs; iCol++)
                {
                    string sColumnName =
                        PWWrapper.aaApi_GetColumnStringProperty(PWWrapper.ColumnProperty.Name, iCol);

                    int iCurrColId =
                        PWWrapper.aaApi_GetColumnNumericProperty(PWWrapper.ColumnProperty.ColumnID, iCol);

                    if (iCurrColId != iColumnId)
                    {
                        slColumnNamesToIds.Add(sColumnName, iCurrColId);
                    }
                }

                foreach (KeyValuePair<string, string> kvp in slCaseInsensitive)
                {
                    if (!string.IsNullOrEmpty(kvp.Value.ToString()))
                    {
                        if (slColumnNamesToIds.ContainsKey(kvp.Key.ToString()))
                        {
                            int iCurrColId = (int)slColumnNamesToIds[kvp.Key.ToString()];

                            string sAttrValue = kvp.Value.ToString();

                            if (iCurrColId != iColumnId)
                            {
                                if (PWWrapper.aaApi_SetLinkDataColumnValue(iTableId,
                                    iCurrColId,
                                    sAttrValue))
                                {
                                    bUpdatedValue = true;
                                }
                            }
                        }
                    }
                } // for each passed in attribute value

                if (bUpdatedValue)
                {
                    int iLinkColId = 0;
                    StringBuilder sbVal = new StringBuilder(30);

                    if (PWWrapper.aaApi_CreateLinkDataAndLink(iTableId, 1,
                        iProjectNo, iDocumentNo, ref iLinkColId, sbVal, sbVal.Capacity))
                    {
                        return true;
                    }
                }
            }
            else // numlinks > 0 && bAddSheet = false
            {
                for (int iRow = 0; iRow < iNumLinks; iRow++)
                {
                    string sUniqueVal =
                        PWWrapper.aaApi_GetLinkStringProperty(PWWrapper.LinkProperty.ColumnValue, iRow);

                    int lNumberOfColumns = 0;

                    PWWrapper.aaApi_SelectLinkDataByObject(iTableId,
                        PWWrapper.ObjectTypeForLinkData.Document,
                        iProjectNo,
                        iDocumentNo,
                        null,
                        ref lNumberOfColumns,
                        null,
                        0);

                    PWWrapper.aaApi_FreeLinkDataUpdateDesc();

                    bool bUpdatedValue = false;

                    for (int iCol = 0; iCol < lNumberOfColumns; iCol++)
                    {
                        string sCurrColumnName =
                            PWWrapper.aaApi_GetLinkDataColumnStringProperty(PWWrapper.LinkDataProperty.ColumnName, iCol);

                        if (slCaseInsensitive.ContainsKey(sCurrColumnName.ToLower()))
                        {
                            try
                            {
                                string sValue = slCaseInsensitive[sCurrColumnName.ToLower()].ToString();

                                int iCurrColId =
                                    PWWrapper.aaApi_GetLinkDataColumnNumericProperty(PWWrapper.LinkDataProperty.ColumnID, iCol);

                                if (iCurrColId != iColumnId)
                                {
                                    if (PWWrapper.aaApi_UpdateLinkDataColumnValue(iTableId,
                                        iCurrColId,
                                        sValue))
                                    {
                                        bUpdatedValue = true;
                                    }
                                }
                            }
                            catch
                            {
                            }
                        }
                    }

                    if (bUpdatedValue)
                    {
                        if (PWWrapper.aaApi_UpdateEnvAttr(iTableId, Convert.ToInt32(sUniqueVal)))
                        // if (PWWrapper.aaApi_UpdateLinkData(iTableId, iColumnId, sUniqueVal))
                        {
                            return true;
                        }
                    }
                } // for each attribute sheet
            }
        } // get table and column ids

        return false;
    }

    public static bool SetAttributesValues
    (
        int iProjectNo, // in - project id
        int iDocumentNo, // in - document id
        SortedList<string, string> slPossiblyCaseSensitive, // in sortedlist with column names as keys and new values as values
        bool bAddSheet
    )
    {
        int iEnvId = 0, iTableId = 0, iColumnId = 0;

        // Need to transform input SortedList to be case insensitive in the keys

        SortedList<string, string> slCaseInsensitive = new SortedList<string, string>(StringComparer.CurrentCultureIgnoreCase);

        foreach (KeyValuePair<string, string> kvp in slPossiblyCaseSensitive)
        {
            if (!slCaseInsensitive.ContainsKey(kvp.Key))
                slCaseInsensitive.Add(kvp.Key, kvp.Value);
        }

        if (PWWrapper.aaApi_GetEnvTableInfoByProject(iProjectNo,
            ref iEnvId, ref iTableId, ref iColumnId))
        {
            int iNumLinks = PWWrapper.aaApi_SelectLinks(iProjectNo, iDocumentNo);

            if (iNumLinks == 0 || bAddSheet)
            {
                PWWrapper.aaApi_FreeLinkDataInsertDesc();

                bool bUpdatedValue = false;

                int iNumAttrs = PWWrapper.aaApi_SelectColumnsByTable(iTableId);

                SortedList<string, int> slColumnNamesToIds = new SortedList<string, int>(StringComparer.CurrentCultureIgnoreCase);

                for (int iCol = 0; iCol < iNumAttrs; iCol++)
                {
                    string sColumnName =
                        PWWrapper.aaApi_GetColumnStringProperty(PWWrapper.ColumnProperty.Name, iCol);

                    int iCurrColId =
                        PWWrapper.aaApi_GetColumnNumericProperty(PWWrapper.ColumnProperty.ColumnID, iCol);

                    if (iCurrColId != iColumnId)
                    {
                        slColumnNamesToIds.Add(sColumnName, iCurrColId);
                    }
                }

                foreach (KeyValuePair<string, string> kvp in slCaseInsensitive)
                {
                    if (!string.IsNullOrEmpty(kvp.Value.ToString()))
                    {
                        if (slColumnNamesToIds.ContainsKey(kvp.Key.ToString()))
                        {
                            int iCurrColId = (int)slColumnNamesToIds[kvp.Key.ToString()];

                            string sAttrValue = kvp.Value.ToString();

                            if (iCurrColId != iColumnId)
                            {
                                if (PWWrapper.aaApi_SetLinkDataColumnValue(iTableId,
                                    iCurrColId,
                                    sAttrValue))
                                {
                                    bUpdatedValue = true;
                                }
                            }
                        }
                    }
                } // for each passed in attribute value

                if (bUpdatedValue)
                {
                    int iLinkColId = 0;
                    StringBuilder sbVal = new StringBuilder(30);

                    if (PWWrapper.aaApi_CreateLinkDataAndLink(iTableId, 1,
                        iProjectNo, iDocumentNo, ref iLinkColId, sbVal, sbVal.Capacity))
                    {
                        return true;
                    }
                }
            }
            else // numlinks > 0 && bAddSheet = false
            {
                for (int iRow = 0; iRow < iNumLinks; iRow++)
                {
                    string sUniqueVal =
                        PWWrapper.aaApi_GetLinkStringProperty(PWWrapper.LinkProperty.ColumnValue, iRow);

                    int lNumberOfColumns = 0;

                    PWWrapper.aaApi_SelectLinkDataByObject(iTableId,
                        PWWrapper.ObjectTypeForLinkData.Document,
                        iProjectNo,
                        iDocumentNo,
                        null,
                        ref lNumberOfColumns,
                        null,
                        0);

                    PWWrapper.aaApi_FreeLinkDataUpdateDesc();

                    bool bUpdatedValue = false;

                    for (int iCol = 0; iCol < lNumberOfColumns; iCol++)
                    {
                        string sCurrColumnName =
                            PWWrapper.aaApi_GetLinkDataColumnStringProperty(PWWrapper.LinkDataProperty.ColumnName, iCol);

                        // so, this should not care about the case now
                        if (slCaseInsensitive.ContainsKey(sCurrColumnName.ToLower()))
                        {
                            try
                            {
                                string sValue = slCaseInsensitive[sCurrColumnName.ToLower()].ToString();

                                int iCurrColId =
                                    PWWrapper.aaApi_GetLinkDataColumnNumericProperty(PWWrapper.LinkDataProperty.ColumnID, iCol);

                                if (iCurrColId != iColumnId)
                                {
                                    if (PWWrapper.aaApi_UpdateLinkDataColumnValue(iTableId,
                                        iCurrColId,
                                        sValue))
                                    {
                                        bUpdatedValue = true;
                                    }
                                }
                            }
                            catch
                            {
                            }
                        }
                    }

                    if (bUpdatedValue)
                    {
                        if (PWWrapper.aaApi_UpdateEnvAttr(iTableId, Convert.ToInt32(sUniqueVal)))
                        // if (PWWrapper.aaApi_UpdateLinkData(iTableId, iColumnId, sUniqueVal))
                        {
                            return true;
                        }
                    }
                } // for each attribute sheet
            }
        } // get table and column ids

        return false;
    }

    public static bool SetAttributesValues
    (
        int iProjectNo, // in - project id
        int iDocumentNo, // in - document id
        Hashtable htAttrVals // in hashtable with column names (lower case) as keys and new values as values
    )
    {
        int iEnvId = 0, iTableId = 0, iColumnId = 0;

        if (PWWrapper.aaApi_GetEnvTableInfoByProject(iProjectNo,
            ref iEnvId, ref iTableId, ref iColumnId))
        {
            int iNumLinks = PWWrapper.aaApi_SelectLinks(iProjectNo, iDocumentNo);

            if (iNumLinks == 0)
            {
                PWWrapper.aaApi_FreeLinkDataInsertDesc();

                bool bUpdatedValue = false;

                int iNumAttrs = PWWrapper.aaApi_SelectColumnsByTable(iTableId);

                Hashtable htColumnNamesToIds = new Hashtable();

                for (int iCol = 0; iCol < iNumAttrs; iCol++)
                {
                    string sColumnName =
                        PWWrapper.aaApi_GetColumnStringProperty(PWWrapper.ColumnProperty.Name, iCol);

                    int iCurrColId =
                        PWWrapper.aaApi_GetColumnNumericProperty(PWWrapper.ColumnProperty.ColumnID, iCol);

                    if (iCurrColId != iColumnId)
                    {
                        htColumnNamesToIds.Add(sColumnName.ToLower(), iCurrColId);
                    }
                }

                foreach (DictionaryEntry de in htAttrVals)
                {
                    if (!string.IsNullOrEmpty(de.Value.ToString()))
                    {
                        if (htColumnNamesToIds.ContainsKey(de.Key.ToString()))
                        {
                            int iCurrColId = (int)htColumnNamesToIds[de.Key.ToString()];

                            string sAttrValue = de.Value.ToString();

                            if (iCurrColId != iColumnId)
                            {
                                if (PWWrapper.aaApi_SetLinkDataColumnValue(iTableId,
                                    iCurrColId,
                                    sAttrValue))
                                {
                                    bUpdatedValue = true;
                                }
                            }
                        }
                    }
                } // for each passed in attribute value

                if (bUpdatedValue)
                {
                    int iLinkColId = 0;
                    StringBuilder sbVal = new StringBuilder(30);

                    if (PWWrapper.aaApi_CreateLinkDataAndLink(iTableId, 1,
                        iProjectNo, iDocumentNo, ref iLinkColId, sbVal, sbVal.Capacity))
                    {
                        return true;
                    }
                }
            }
            else
            {
                for (int iRow = 0; iRow < iNumLinks; iRow++)
                {
                    string sUniqueVal =
                        PWWrapper.aaApi_GetLinkStringProperty(PWWrapper.LinkProperty.ColumnValue, iRow);

                    int lNumberOfColumns = 0;

                    PWWrapper.aaApi_SelectLinkDataByObject(iTableId,
                        PWWrapper.ObjectTypeForLinkData.Document,
                        iProjectNo,
                        iDocumentNo,
                        null,
                        ref lNumberOfColumns,
                        null,
                        0);

                    PWWrapper.aaApi_FreeLinkDataUpdateDesc();

                    bool bUpdatedValue = false;

                    for (int iCol = 0; iCol < lNumberOfColumns; iCol++)
                    {
                        string sCurrColumnName =
                            PWWrapper.aaApi_GetLinkDataColumnStringProperty(PWWrapper.LinkDataProperty.ColumnName, iCol);

                        if (htAttrVals.ContainsKey(sCurrColumnName.ToLower()))
                        {
                            try
                            {

                                string sValue = htAttrVals[sCurrColumnName.ToLower()].ToString();

                                int iCurrColId =
                                    PWWrapper.aaApi_GetLinkDataColumnNumericProperty(PWWrapper.LinkDataProperty.ColumnID, iCol);

                                if (iCurrColId != iColumnId)
                                {
                                    if (PWWrapper.aaApi_UpdateLinkDataColumnValue(iTableId,
                                        iCurrColId,
                                        sValue))
                                    {
                                        bUpdatedValue = true;
                                    }
                                }
                            }
                            catch
                            {
                            }
                        }
                    }

                    if (bUpdatedValue)
                    {
                        if (PWWrapper.aaApi_UpdateEnvAttr(iTableId, Convert.ToInt32(sUniqueVal)))
                        // if (PWWrapper.aaApi_UpdateLinkData(iTableId, iColumnId, sUniqueVal))
                        {
                            return true;
                        }
                    }
                } // for each attribute sheet
            }
        } // get table and column ids

        return false;
    }

    public static bool SetAttributesValuesFromColumnIds
    (
        int iProjectNo, // in - project id
        int iDocumentNo, // in - document id
        Hashtable htAttrVals // in hashtable with column ids as keys and new values as values
    )
    {
        int iEnvId = 0, iTableId = 0, iColumnId = 0;

        if (PWWrapper.aaApi_GetEnvTableInfoByProject(iProjectNo,
            ref iEnvId, ref iTableId, ref iColumnId))
        {
            int iNumLinks = PWWrapper.aaApi_SelectLinks(iProjectNo, iDocumentNo);

            if (iNumLinks == 0)
            {
                PWWrapper.aaApi_FreeLinkDataInsertDesc();

                bool bUpdatedValue = false;

                foreach (DictionaryEntry de in htAttrVals)
                {
                    if (!string.IsNullOrEmpty(de.Value.ToString()))
                    {
                        int iCurrColId = (int)de.Key;

                        string sAttrValue = de.Value.ToString();

                        if (iCurrColId != iColumnId)
                        {
                            if (PWWrapper.aaApi_SetLinkDataColumnValue(iTableId,
                                iCurrColId,
                                sAttrValue))
                            {
                                bUpdatedValue = true;
                            }
                        }
                    }
                } // for each passed in attribute value

                if (bUpdatedValue)
                {
                    int iLinkColId = 0;
                    StringBuilder sbVal = new StringBuilder(30);

                    if (PWWrapper.aaApi_CreateLinkDataAndLink(iTableId, 1,
                        iProjectNo, iDocumentNo, ref iLinkColId, sbVal, sbVal.Capacity))
                    {
                        return true;
                    }
                }
            }
            else
            {
                for (int iRow = 0; iRow < iNumLinks; iRow++)
                {
                    string sUniqueVal =
                        PWWrapper.aaApi_GetLinkStringProperty(PWWrapper.LinkProperty.ColumnValue, iRow);

                    PWWrapper.aaApi_FreeLinkDataUpdateDesc();

                    bool bUpdatedValue = false;

                    foreach (DictionaryEntry de in htAttrVals)
                    {
                        if (!string.IsNullOrEmpty(de.Value.ToString()))
                        {
                            int iCurrColId = (int)de.Key;

                            string sAttrValue = de.Value.ToString();

                            if (iCurrColId != iColumnId)
                            {
                                if (PWWrapper.aaApi_UpdateLinkDataColumnValue(iTableId,
                                    iCurrColId,
                                    sAttrValue))
                                {
                                    bUpdatedValue = true;
                                }
                            }
                        }
                    } // for each passed in attribute value

                    if (bUpdatedValue)
                    {
                        if (PWWrapper.aaApi_UpdateEnvAttr(iTableId, Convert.ToInt32(sUniqueVal)))
                        // if (PWWrapper.aaApi_UpdateLinkData(iTableId, iColumnId, sUniqueVal))
                        {
                            return true;
                        }
                    }
                } // for each attribute sheet
            }
        } // get table and column ids

        return false;
    }

    public static bool SetAttributesValuesFromColumnIds
    (
        int iProjectNo, // in - project id
        int iDocumentNo, // in - document id
        SortedList<int, string> slAttrVals // in hashtable with column ids as keys and new values as values
    )
    {
        int iEnvId = 0, iTableId = 0, iColumnId = 0;

        if (PWWrapper.aaApi_GetEnvTableInfoByProject(iProjectNo,
            ref iEnvId, ref iTableId, ref iColumnId))
        {
            int iNumLinks = PWWrapper.aaApi_SelectLinks(iProjectNo, iDocumentNo);

            if (iNumLinks == 0)
            {
                PWWrapper.aaApi_FreeLinkDataInsertDesc();

                bool bUpdatedValue = false;

                foreach (KeyValuePair<int, string> kvp in slAttrVals)
                {
                    if (kvp.Key != iColumnId)
                    {
                        if (PWWrapper.aaApi_SetLinkDataColumnValue(iTableId,
                            kvp.Key,
                            kvp.Value))
                        {
                            bUpdatedValue = true;
                        }
                    }
                } // for each passed in attribute value

                if (bUpdatedValue)
                {
                    int iLinkColId = 0;
                    StringBuilder sbVal = new StringBuilder(30);

                    if (PWWrapper.aaApi_CreateLinkDataAndLink(iTableId, 1,
                        iProjectNo, iDocumentNo, ref iLinkColId, sbVal, sbVal.Capacity))
                    {
                        return true;
                    }
                }
            }
            else
            {
                for (int iRow = 0; iRow < iNumLinks; iRow++)
                {
                    string sUniqueVal =
                        PWWrapper.aaApi_GetLinkStringProperty(PWWrapper.LinkProperty.ColumnValue, iRow);

                    PWWrapper.aaApi_FreeLinkDataUpdateDesc();

                    bool bUpdatedValue = false;

                    foreach (KeyValuePair<int, string> kvp in slAttrVals)
                    {
                        if (kvp.Key != iColumnId)
                        {
                            if (PWWrapper.aaApi_UpdateLinkDataColumnValue(iTableId,
                                kvp.Key,
                                kvp.Value))
                            {
                                bUpdatedValue = true;
                            }
                        }
                    } // for each passed in attribute value

                    if (bUpdatedValue)
                    {
                        // if (PWWrapper.aaApi_UpdateLinkData(iTableId, iColumnId, sUniqueVal))
                        if (PWWrapper.aaApi_UpdateEnvAttr(iTableId, Convert.ToInt32(sUniqueVal)))
                        {
                            return true;
                        }
                    }
                } // for each attribute sheet
            }
        } // get table and column ids

        return false;
    }

    public static bool SetAttributeValue(int iProjectNo, int iDocumentNo, string sColumnName, string sValue)
    {
        int iEnvId = 0, iTableId = 0, iColumnId = 0;

        if (PWWrapper.aaApi_GetEnvTableInfoByProject(iProjectNo,
            ref iEnvId, ref iTableId, ref iColumnId))
        {
            int iNumLinks = PWWrapper.aaApi_SelectLinks(iProjectNo, iDocumentNo);

            if (iNumLinks == 0)
            {
                PWWrapper.aaApi_FreeLinkDataInsertDesc();

                bool bUpdatedValue = false;

                int iNumAttrs = PWWrapper.aaApi_SelectColumnsByTable(iTableId);

                int iTargetColumnId = 0;

                for (int iCol = 0; iCol < iNumAttrs; iCol++)
                {
                    string sCurrColumnName =
                        PWWrapper.aaApi_GetColumnStringProperty(PWWrapper.ColumnProperty.Name, iCol);

                    if (sCurrColumnName.ToLower() == sColumnName.ToLower())
                    {
                        iTargetColumnId =
                            PWWrapper.aaApi_GetColumnNumericProperty(PWWrapper.ColumnProperty.ColumnID, iCol);
                        break;
                    }
                }

                if (iTargetColumnId > 0)
                {
                    if (iTargetColumnId != iColumnId)
                    {
                        if (PWWrapper.aaApi_SetLinkDataColumnValue(iTableId,
                            iTargetColumnId,
                            sValue))
                        {
                            bUpdatedValue = true;
                        }
                    }
                } // for each passed in attribute value

                if (bUpdatedValue)
                {
                    int iLinkColId = 0;
                    StringBuilder sbVal = new StringBuilder(30);

                    if (PWWrapper.aaApi_CreateLinkDataAndLink(iTableId, 1,
                        iProjectNo, iDocumentNo, ref iLinkColId, sbVal, sbVal.Capacity))
                    {
                        return true;
                    }
                }
            }
            else
            {
                for (int iRow = 0; iRow < iNumLinks; iRow++)
                {
                    string sUniqueVal =
                        PWWrapper.aaApi_GetLinkStringProperty(PWWrapper.LinkProperty.ColumnValue, iRow);

                    int lNumberOfColumns = 0;

                    PWWrapper.aaApi_SelectLinkDataByObject(iTableId,
                        PWWrapper.ObjectTypeForLinkData.Document,
                        iProjectNo,
                        iDocumentNo,
                        null,
                        ref lNumberOfColumns,
                        null,
                        0);

                    PWWrapper.aaApi_FreeLinkDataUpdateDesc();

                    bool bUpdatedValue = false;

                    for (int iCol = 0; iCol < lNumberOfColumns; iCol++)
                    {
                        string sCurrColumnName =
                            PWWrapper.aaApi_GetLinkDataColumnStringProperty(PWWrapper.LinkDataProperty.ColumnName, iCol);

                        if (sColumnName.ToLower() == sCurrColumnName.ToLower())
                        {
                            int iCurrColId =
                                PWWrapper.aaApi_GetLinkDataColumnNumericProperty(PWWrapper.LinkDataProperty.ColumnID, iCol);

                            int iColLen =
                                PWWrapper.aaApi_GetLinkDataColumnNumericProperty(PWWrapper.LinkDataProperty.ColumnLength, iCol);

                            if (iColLen > 2)
                            {
                                if (sValue.Length > iColLen)
                                {
                                    sValue = sValue.Substring(0, iColLen - 1);
                                }
                            }

                            if (iCurrColId != iColumnId)
                            {
                                if (PWWrapper.aaApi_UpdateLinkDataColumnValue(iTableId,
                                    iCurrColId,
                                    sValue))
                                {
                                    bUpdatedValue = true;
                                    break;
                                }
                            }
                        }
                    }

                    if (bUpdatedValue)
                    {
                        if (PWWrapper.aaApi_UpdateEnvAttr(iTableId, Convert.ToInt32(sUniqueVal)))
                            // if (PWWrapper.aaApi_UpdateLinkData(iTableId, iColumnId, sUniqueVal))
                            return true;
                    }
                } // for each attribute sheet
            }
        } // get table and column ids

        return false;
    }


#endregion

#region Views

    public static string aaApi_ViewGetName(IntPtr hView)
    {
        return Marshal.PtrToStringUni(_aaApi_ViewGetName(hView));
    }

    [DllImport("dmawin.dll", CharSet = CharSet.Unicode, EntryPoint = "aaApi_ViewGetName")]
    private static extern IntPtr _aaApi_ViewGetName(IntPtr hView);

    [DllImport("dmawin.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaApi_ViewGetFirst();

    [DllImport("dmawin.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaApi_ViewGetNext(IntPtr hView);

    [DllImport("dmawin.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaApi_ViewGetFirstForProject(int iProjectId);

    [DllImport("dmawin.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaApi_ViewGetHandle(string sViewName);

    [DllImport("dmawin.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaApi_ViewGetHandleById(int iViewId);

    [DllImport("dmawin.dll", CharSet = CharSet.Unicode)]
    public static extern uint aaApi_ViewColumnGetFirst(IntPtr pViewHandle);

    [DllImport("dmawin.dll", CharSet = CharSet.Unicode)]
    public static extern uint aaApi_ViewColumnGetNext(uint uiViewColumn);

    [DllImport("dmawin.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_ViewColumnGetTable(uint uiViewColumn);

    [DllImport("dmawin.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_ViewColumnGetField(uint uiViewColumn);

    [DllImport("dmawin.dll", EntryPoint = "aaApi_ViewColumnGetFieldName", CharSet = CharSet.Unicode)]
    private static extern IntPtr _aaApi_ViewColumnGetFieldName(uint uiViewColumn);

    public static string aaApi_ViewColumnGetFieldName(uint uiViewColumn)
    {
        return Marshal.PtrToStringUni(_aaApi_ViewColumnGetFieldName(uiViewColumn));
    }

    [DllImport("dmawin.dll", EntryPoint = "aaApi_ViewColumnGetName", CharSet = CharSet.Unicode)]
    private static extern IntPtr _aaApi_ViewColumnGetName(uint uiViewColumn);

    public static string aaApi_ViewColumnGetName(uint uiViewColumn)
    {
        return Marshal.PtrToStringUni(_aaApi_ViewColumnGetName(uiViewColumn));
    }

    [DllImport("dmawin.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaApi_GetActiveDocumentList(); // can't believe this wasn't in here...

    [DllImport("dmawin.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaApi_GetDocumentListCtrl(IntPtr hDocumentList);

    [DllImport("dmactrl.dll", CharSet = CharSet.Unicode)]
    public static extern int aaApi_CopyListControlContent(IntPtr hWndList,
      uint ulFlags,
      ref int lplSubItems,
      int lSubItemCount,
      string lpctstrColSeparator,
      string lpctstrRowSeparator
     );

    [DllImport("dmawin.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_DocListGetData(IntPtr hWndDocList,
        int iItem,
        ref int lplProjectId,
        ref int lplDocumentId,
        ref int lplSetId
    );


    [DllImport("dmawin.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_DocListSetView(IntPtr hWndDocList,
        IntPtr hView);

    public static bool SetDocumentListView(string sViewName)
    {
        IntPtr hView = aaApi_ViewGetHandle(sViewName);

        if (hView != IntPtr.Zero)
        {
            return aaApi_DocListSetView(aaApi_GetActiveDocumentList(), hView);
        }

        return false;
    }

    public static bool SetDocumentListViewByProject(int iProjectId, string sViewName)
    {
        IntPtr hView = aaApi_ViewGetFirstForProject(iProjectId);

        while (hView != IntPtr.Zero)
        {
            string sCurrentViewName = aaApi_ViewGetName(hView);

            if (sCurrentViewName.ToLower() == sViewName.ToLower())
            {
                return aaApi_DocListSetView(aaApi_GetActiveDocumentList(), hView);
            }

            hView = aaApi_ViewGetNext(hView);
        }

        return false;
    }

#endregion

}

public class PWSearch
{
    // added per Dan 2020-01-17
    [DllImport("PWSearchWrapperX64.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall, EntryPoint = "CreateSearch")]
    public static extern int CreateSearchX64
          (
              int iProjectId,
              bool bIncludeSubVaults,
              string wcFullTextString,
              bool bWholePhrase,
              bool bAnyWord, // otherwise all words
              bool bSearchAttributes, // otherwise full text at this point
              string szDocumentNameP,
              string szFileNameP,
              string szDocumentDescP,
              bool bOriginalsOnly,
              int iEnvId,
              [In][MarshalAsAttribute(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] arAttributeNames,
              [In][MarshalAsAttribute(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] arAttributeValues,
              int iSize,
              int iWorkflowId,
              string szStatesP, // comma delimited states
              string szUpdatedAfterP, // early date 2009-10-22 01:00:00
              string szUpdatedBeforeP, // late date 2010-10-22 01:00:00
             string szSavedSearchNameP,
              int lParentProjectId,
              int lParentQueryId,
              string sViewName
          );

    [DllImport("PWSearchWrapper.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall)]
    public static extern int CreateSearch
    (
        int iProjectId,
        bool bIncludeSubVaults,
        string wcFullTextString,
        bool bWholePhrase,
        bool bAnyWord, // otherwise all words
        bool bSearchAttributes, // otherwise full text at this point
        string szDocumentNameP,
        string szFileNameP,
        string szDocumentDescP,
        bool bOriginalsOnly,
        int iEnvId,
        [In][MarshalAsAttribute(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] arAttributeNames,
        [In][MarshalAsAttribute(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] arAttributeValues,
        int iSize,
        int iWorkflowId,
        string szStatesP, // comma delimited states
        string szUpdatedAfterP, // early date 2009-10-22 01:00:00
        string szUpdatedBeforeP, // late date 2010-10-22 01:00:00
        string szSavedSearchNameP,
        int lParentProjectId,
        int lParentQueryId,
        string sViewName
    );

    [DllImport("PWSearchWrapper.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall)]
    public static extern int CreateSearchFolder
    (
        string szSavedSearchNameP,
        int lParentProjectId,
        int lParentQueryId
    );

    [DllImport("PWSearchWrapperX64.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall, EntryPoint = "CreateSearchFolder")]
    public static extern int CreateSearchFolderX64
    (
        string szSavedSearchNameP,
        int lParentProjectId,
        int lParentQueryId
    );

    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern bool aaApi_SQueryDelete(int iQueryId);

    // above added per Dan 2020-01-17

    [DllImport("PWSearchWrapper.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall)]
    private extern static int SearchForProjectsByProperties(
            int iParentProjectId,
            string sProjectType,
            [In][MarshalAsAttribute(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] propertyNames,
            [In][MarshalAsAttribute(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] propertyValues,
            int size,
            [Out] out IntPtr ppProjects,
            [Out] out int iCountP
        );

    [DllImport("PWSearchWrapperX64.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall, EntryPoint = "SearchForProjectsByProperties")]
    private extern static int SearchForProjectsByPropertiesX64(
            int iParentProjectId,
            string sProjectType,
            [In][MarshalAsAttribute(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] propertyNames,
            [In][MarshalAsAttribute(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] propertyValues,
            int size,
            [Out] out IntPtr ppProjects,
            [Out] out int iCountP
        );

    [DllImport("PWSearchWrapper.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall)]
    private extern static int SearchForProjectsByName(
            string sProjectName,
            [Out] out IntPtr ppProjects,
            [Out] out int iCountP
        );

    [DllImport("PWSearchWrapperX64.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall, EntryPoint = "SearchForProjectsByName")]
    private extern static int SearchForProjectsByNameX64(
            string sProjectName,
            [Out] out IntPtr ppProjects,
            [Out] out int iCountP
        );


    [DllImport("PWSearchWrapper.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall)]
    public extern static int GetSavedSearchId(
            string sSearchName
        );

    [DllImport("PWSearchWrapperX64.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall, EntryPoint = "GetSavedSearchId")]
    public extern static int GetSavedSearchIdX64(
            string sSearchName
        );

    [DllImport("PWSearchWrapper.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall)]
    private extern static int SearchForProjectsByTree(
            int iParentProjectId,
            [Out] out IntPtr ppProjects,
             [Out] out IntPtr ppComponentClassIds,
             [Out] out IntPtr ppComponentInstanceIds,
             [Out] out IntPtr ppEnvironmentIds,
             [Out] out IntPtr ppWorkflowIds,
             [Out] out IntPtr ppParentIds,
             [Out] out IntPtr ppStorageIds,
             [Out] out IntPtr ppIsParent,
             [Out] out IntPtr ppProjectTypes,
                [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
                    out string[] arProjectNames,
                [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
                    out string[] arProjectDescriptions,
                [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
                    out string[] arProjectCodes,
                [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
                    out string[] arProjectGuids,
             [Out] out int iCountP
        );

    [DllImport("PWSearchWrapperX64.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall, EntryPoint = "SearchForProjectsByTree")]
    private extern static int SearchForProjectsByTreeX64(
            int iParentProjectId,
            [Out] out IntPtr ppProjects,
             [Out] out IntPtr ppComponentClassIds,
             [Out] out IntPtr ppComponentInstanceIds,
             [Out] out IntPtr ppEnvironmentIds,
             [Out] out IntPtr ppWorkflowIds,
             [Out] out IntPtr ppParentIds,
             [Out] out IntPtr ppStorageIds,
             [Out] out IntPtr ppIsParent,
             [Out] out IntPtr ppProjectTypes,
                [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
                    out string[] arProjectNames,
                [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
                    out string[] arProjectDescriptions,
                [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
                    out string[] arProjectCodes,
                [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
                    out string[] arProjectGuids,
             [Out] out int iCountP
        );

    [DllImport("PWSearchWrapper.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall)]
    private extern static int SearchForProjectsByClassId(
            int iParentProjectId,
            int iClassId,
            [Out] out IntPtr ppProjects,
             [Out] out IntPtr ppComponentClassIds,
             [Out] out IntPtr ppComponentInstanceIds,
             [Out] out IntPtr ppEnvironmentIds,
             [Out] out IntPtr ppWorkflowIds,
             [Out] out IntPtr ppParentIds,
             [Out] out IntPtr ppStorageIds,
             [Out] out IntPtr ppIsParent,
             [Out] out IntPtr ppProjectTypes,
                [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
                    out string[] arProjectNames,
                [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
                    out string[] arProjectDescriptions,
                [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
                    out string[] arProjectCodes,
             [Out] out int iCountP
        );

    [DllImport("PWSearchWrapperX64.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall, EntryPoint = "SearchForProjectsByClassId")]
    private extern static int SearchForProjectsByClassIdX64(
            int iParentProjectId,
            int iClassId,
            [Out] out IntPtr ppProjects,
             [Out] out IntPtr ppComponentClassIds,
             [Out] out IntPtr ppComponentInstanceIds,
             [Out] out IntPtr ppEnvironmentIds,
             [Out] out IntPtr ppWorkflowIds,
             [Out] out IntPtr ppParentIds,
             [Out] out IntPtr ppStorageIds,
             [Out] out IntPtr ppIsParent,
             [Out] out IntPtr ppProjectTypes,
                [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
                    out string[] arProjectNames,
                [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
                    out string[] arProjectDescriptions,
                [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
                    out string[] arProjectCodes,
             [Out] out int iCountP
        );


    [DllImport("PWSearchWrapper.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall)]
    private extern static int SearchForProjectsByTreeForEmptyFolders(
            int iParentProjectId,
            [Out] out IntPtr ppProjects,
             [Out] out IntPtr ppComponentClassIds,
             [Out] out IntPtr ppComponentInstanceIds,
             [Out] out IntPtr ppEnvironmentIds,
             [Out] out IntPtr ppWorkflowIds,
             [Out] out IntPtr ppParentIds,
             [Out] out IntPtr ppStorageIds,
             [Out] out IntPtr ppIsParent,
             [Out] out IntPtr ppProjectTypes,
                [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
                    out string[] arProjectNames,
                [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
                    out string[] arProjectDescriptions,
                [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
                    out string[] arProjectCodes,
             [Out] out int iCountP
        );

    [DllImport("PWSearchWrapperX64.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall, EntryPoint = "SearchForProjectsByTreeForEmptyFolders")]
    private extern static int SearchForProjectsByTreeForEmptyFoldersX64(
            int iParentProjectId,
            [Out] out IntPtr ppProjects,
             [Out] out IntPtr ppComponentClassIds,
             [Out] out IntPtr ppComponentInstanceIds,
             [Out] out IntPtr ppEnvironmentIds,
             [Out] out IntPtr ppWorkflowIds,
             [Out] out IntPtr ppParentIds,
             [Out] out IntPtr ppStorageIds,
             [Out] out IntPtr ppIsParent,
             [Out] out IntPtr ppProjectTypes,
                [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
                    out string[] arProjectNames,
                [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
                    out string[] arProjectDescriptions,
                [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
                    out string[] arProjectCodes,
             [Out] out int iCountP
        );

    [DllImport("PWSearchWrapper.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall)]
    private extern static int SearchForProjectsByTreeAndEnvironment(
            int iParentProjectId,
            int iEnvironmentId,
            [Out] out IntPtr ppProjects,
             [Out] out IntPtr ppComponentClassIds,
             [Out] out IntPtr ppComponentInstanceIds,
             [Out] out IntPtr ppEnvironmentIds,
             [Out] out IntPtr ppWorkflowIds,
             [Out] out IntPtr ppParentIds,
             [Out] out IntPtr ppStorageIds,
             [Out] out IntPtr ppIsParent,
             [Out] out IntPtr ppProjectTypes,
                [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
                    out string[] arProjectNames,
                [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
                    out string[] arProjectDescriptions,
                [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
                    out string[] arProjectCodes,
             [Out] out int iCountP
        );

    [DllImport("PWSearchWrapperX64.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall, EntryPoint = "SearchForProjectsByTreeAndEnvironment")]
    private extern static int SearchForProjectsByTreeAndEnvironmentX64(
            int iParentProjectId,
            int iEnvironmentId,
            [Out] out IntPtr ppProjects,
             [Out] out IntPtr ppComponentClassIds,
             [Out] out IntPtr ppComponentInstanceIds,
             [Out] out IntPtr ppEnvironmentIds,
             [Out] out IntPtr ppWorkflowIds,
             [Out] out IntPtr ppParentIds,
             [Out] out IntPtr ppStorageIds,
             [Out] out IntPtr ppIsParent,
             [Out] out IntPtr ppProjectTypes,
                [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
                    out string[] arProjectNames,
                [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
                    out string[] arProjectDescriptions,
                [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
                    out string[] arProjectCodes,
             [Out] out int iCountP
        );


    [DllImport("PWSearchWrapper.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall)]
    private extern static int SearchForDocuments(
        int iProjectId,
        bool bIncludeSubFolders,
        string sFullTextSearchString,
        bool bWholePhrase,
        bool bAnyWord,
        bool bSearchAttributes, // otherwise full text
        string sDocumentName,
        string sFileName,
        string sDocumentDescription,
        bool bOriginalsOnly,
        int iEnvironmentId,
        [In][MarshalAsAttribute(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] arAttributeNames,
        [In][MarshalAsAttribute(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] arAttributeValues,
        int size,
        [Out] out IntPtr ppProjects,
        [Out] out IntPtr ppDocumentIds,
        [Out] out IntPtr ppVersionSeqNumbers,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentGuidStrings,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentNames,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentFileNames,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentDescriptions,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentUpdateDates,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arVersions
        );

    [DllImport("PWSearchWrapperX64.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall, EntryPoint = "SearchForDocuments")]
    private extern static int SearchForDocumentsX64(
        int iProjectId,
        bool bIncludeSubFolders,
        string sFullTextSearchString,
        bool bWholePhrase,
        bool bAnyWord,
        bool bSearchAttributes, // otherwise full text
        string sDocumentName,
        string sFileName,
        string sDocumentDescription,
        bool bOriginalsOnly,
        int iEnvironmentId,
        [In][MarshalAsAttribute(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] arAttributeNames,
        [In][MarshalAsAttribute(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] arAttributeValues,
        int size,
        [Out] out IntPtr ppProjects,
        [Out] out IntPtr ppDocumentIds,
        [Out] out IntPtr ppVersionSeqNumbers,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentGuidStrings,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentNames,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentFileNames,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentDescriptions,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentUpdateDates,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arVersions
        );

    [DllImport("PWSearchWrapper.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall)]
    private extern static int SearchForDocumentsByItemType
    (
        int iProjectId,
        bool bIncludeSubVaults,
        int iItemType,
        string szDocumentNameP,
        string szFileNameP,
        string szDocumentDescP,
        bool bOriginalsOnly,
        int iEnvId,
        [In][MarshalAsAttribute(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] arAttributeNames,
        [In][MarshalAsAttribute(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] arAttributeValues,
        int size,
        [Out] out IntPtr ppProjects,
        [Out] out IntPtr ppDocumentIds,
        [Out] out IntPtr ppVersionSeqNumbers,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentGuidStrings,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentNames,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentFileNames,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentDescriptions,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentUpdateDates,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arVersions
    );

    [DllImport("PWSearchWrapperX64.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall, EntryPoint = "SearchForDocumentsByItemType")]
    private extern static int SearchForDocumentsByItemTypeX64
    (
        int iProjectId,
        bool bIncludeSubVaults,
        int iItemType,
        string szDocumentNameP,
        string szFileNameP,
        string szDocumentDescP,
        bool bOriginalsOnly,
        int iEnvId,
        [In][MarshalAsAttribute(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] arAttributeNames,
        [In][MarshalAsAttribute(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] arAttributeValues,
        int size,
        [Out] out IntPtr ppProjects,
        [Out] out IntPtr ppDocumentIds,
        [Out] out IntPtr ppVersionSeqNumbers,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentGuidStrings,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentNames,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentFileNames,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentDescriptions,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentUpdateDates,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arVersions
    );

    [DllImport("PWSearchWrapper.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall, EntryPoint = "SearchForDocsReturnWFStateSizeStorageMimeType")]
    private extern static int SearchForDocumentsReturningWorkflowStateSizesAsStrings(
        int iProjectId,
        bool bIncludeSubFolders,
        string sFullTextSearchString,
        bool bWholePhrase,
        bool bAnyWord,
        bool bSearchAttributes, // otherwise full text
        string sDocumentName,
        string sFileName,
        string sDocumentDescription,
        bool bOriginalsOnly,
        int iEnvironmentId,
        [In][MarshalAsAttribute(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] arAttributeNames,
        [In][MarshalAsAttribute(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] arAttributeValues,
        int size,
        [Out] out IntPtr ppProjects,
        [Out] out IntPtr ppDocumentIds,
        [Out] out IntPtr ppVersionSeqNumbers,
        [Out] out IntPtr ppWorkflowIds,
        [Out] out IntPtr ppStateIds,
        [Out] out IntPtr ppStorageIds,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentGuidStrings,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentNames,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentFileNames,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentDescriptions,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentUpdateDates,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arVersions,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arFileSizes,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arMimeTypes
        );

    [DllImport("PWSearchWrapperX64.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall, EntryPoint = "SearchForDocsReturnWFStateSizeStorageMimeType")]
    private extern static int SearchForDocumentsReturningWorkflowStateSizesAsStringsX64(
        int iProjectId,
        bool bIncludeSubFolders,
        string sFullTextSearchString,
        bool bWholePhrase,
        bool bAnyWord,
        bool bSearchAttributes, // otherwise full text
        string sDocumentName,
        string sFileName,
        string sDocumentDescription,
        bool bOriginalsOnly,
        int iEnvironmentId,
        [In][MarshalAsAttribute(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] arAttributeNames,
        [In][MarshalAsAttribute(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] arAttributeValues,
        int size,
        [Out] out IntPtr ppProjects,
        [Out] out IntPtr ppDocumentIds,
        [Out] out IntPtr ppVersionSeqNumbers,
        [Out] out IntPtr ppWorkflowIds,
        [Out] out IntPtr ppStateIds,
        [Out] out IntPtr ppStorageIds,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentGuidStrings,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentNames,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentFileNames,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentDescriptions,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentUpdateDates,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arVersions,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arFileSizes,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arMimeTypes
        );

    [DllImport("PWSearchWrapper.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall, EntryPoint = "SearchForDocsMultipleValuesReturnWFStateSizeStorageMimeType")]
    private extern static int SearchForDocumentsMultiValuesReturningWorkflowStateSizesAsStrings(
        int iProjectId,
        bool bIncludeSubFolders,
        string sDocumentName,
        string sFileName,
        string sDocumentDescription,
        bool bOriginalsOnly,
        int iEnvironmentId,
        string sAttributeName,
        [In][MarshalAsAttribute(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] arAttributeValues,
        int size,
        [Out] out IntPtr ppProjects,
        [Out] out IntPtr ppDocumentIds,
        [Out] out IntPtr ppVersionSeqNumbers,
        [Out] out IntPtr ppWorkflowIds,
        [Out] out IntPtr ppStateIds,
        [Out] out IntPtr ppStorageIds,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentGuidStrings,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentNames,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentFileNames,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentDescriptions,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentUpdateDates,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arVersions,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arFileSizes,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arMimeTypes
        );

    [DllImport("PWSearchWrapperX64.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall, EntryPoint = "SearchForDocsMultipleValuesReturnWFStateSizeStorageMimeType")]
    private extern static int SearchForDocumentsMultiValuesReturningWorkflowStateSizesAsStringsX64(
        int iProjectId,
        bool bIncludeSubFolders,
        string sDocumentName,
        string sFileName,
        string sDocumentDescription,
        bool bOriginalsOnly,
        int iEnvironmentId,
        string sAttributeName,
        [In][MarshalAsAttribute(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] arAttributeValues,
        int size,
        [Out] out IntPtr ppProjects,
        [Out] out IntPtr ppDocumentIds,
        [Out] out IntPtr ppVersionSeqNumbers,
        [Out] out IntPtr ppWorkflowIds,
        [Out] out IntPtr ppStateIds,
        [Out] out IntPtr ppStorageIds,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentGuidStrings,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentNames,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentFileNames,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentDescriptions,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentUpdateDates,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arVersions,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arFileSizes,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arMimeTypes
        );

    [DllImport("PWSearchWrapper.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall, EntryPoint = "SearchForDocsByQueryIdReturnWFStateSizesStorageMimeType")]
    private extern static int SearchForDocumentsByQueryIdReturningWorkflowStateSizesAsStrings(
        int iQueryId,
        [Out] out IntPtr ppProjects,
        [Out] out IntPtr ppDocumentIds,
        [Out] out IntPtr ppVersionSeqNumbers,
        [Out] out IntPtr ppWorkflowIds,
        [Out] out IntPtr ppStateIds,
        [Out] out IntPtr ppStorageIds,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentGuidStrings,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentNames,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentFileNames,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentDescriptions,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentUpdateDates,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arVersions,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arFileSizes,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arMimeTypes
        );

    [DllImport("PWSearchWrapperX64.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall, EntryPoint = "SearchForDocumentsByQueryIdReturningWorkflowStateSizesAsStrings")]
    private extern static int SearchForDocumentsByQueryIdReturningWorkflowStateSizesAsStringsX64_Old(
        int iQueryId,
        [Out] out IntPtr ppProjects,
        [Out] out IntPtr ppDocumentIds,
        [Out] out IntPtr ppVersionSeqNumbers,
        [Out] out IntPtr ppWorkflowIds,
        [Out] out IntPtr ppStateIds,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentGuidStrings,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentNames,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentFileNames,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentDescriptions,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentUpdateDates,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arVersions,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arFileSizes
        );

    [DllImport("PWSearchWrapperX64.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall, EntryPoint = "SearchForDocsByQueryIdReturnWFStateSizesStorageMimeType")]
    private extern static int SearchForDocumentsByQueryIdReturningWorkflowStateSizesAsStringsX64(
        int iQueryId,
        [Out] out IntPtr ppProjects,
        [Out] out IntPtr ppDocumentIds,
        [Out] out IntPtr ppVersionSeqNumbers,
        [Out] out IntPtr ppWorkflowIds,
        [Out] out IntPtr ppStateIds,
        [Out] out IntPtr ppStorageIds,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentGuidStrings,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentNames,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentFileNames,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentDescriptions,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentUpdateDates,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arVersions,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arFileSizes,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arMimeTypes
        );
    
    [DllImport("PWSearchWrapper.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall)]
    private extern static int SearchForDocumentsWithStatesAndDates(
        int iProjectId,
        bool bIncludeSubFolders,
        string sFullTextSearchString,
        bool bWholePhrase,
        bool bAnyWord,
        bool bSearchAttributes, // otherwise full text
        string sDocumentName,
        string sFileName,
        string sDocumentDescription,
        bool bOriginalsOnly,
        int iEnvironmentId,
        [In][MarshalAsAttribute(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] arAttributeNames,
        [In][MarshalAsAttribute(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] arAttributeValues,
        int size,
        string sStates, // comma delimited state numbers
        string sUpdatedAfter, // early date 2009-10-22 01:00:00
        string sUpdatedBefore, // late date 2010-10-22 01:00:00
        [Out] out IntPtr ppProjects,
        [Out] out IntPtr ppDocumentIds,
        [Out] out IntPtr ppVersionSequenceNums,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentGuidStrings,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentNames,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentFileNames,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentDescriptions,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentUpdateDates,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentVersions
        );

    [DllImport("PWSearchWrapperX64.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall, EntryPoint = "SearchForDocumentsWithStatesAndDates")]
    private extern static int SearchForDocumentsWithStatesAndDatesX64(
        int iProjectId,
        bool bIncludeSubFolders,
        string sFullTextSearchString,
        bool bWholePhrase,
        bool bAnyWord,
        bool bSearchAttributes, // otherwise full text
        string sDocumentName,
        string sFileName,
        string sDocumentDescription,
        bool bOriginalsOnly,
        int iEnvironmentId,
        [In][MarshalAsAttribute(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] arAttributeNames,
        [In][MarshalAsAttribute(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] arAttributeValues,
        int size,
        string sStates, // comma delimited state numbers
        string sUpdatedAfter, // early date 2009-10-22 01:00:00
        string sUpdatedBefore, // late date 2010-10-22 01:00:00
        [Out] out IntPtr ppProjects,
        [Out] out IntPtr ppDocumentIds,
        [Out] out IntPtr ppVersionSequenceNums,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
                    out string[] arDocumentGuidStrings,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
                    out string[] arDocumentNames,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
                    out string[] arDocumentFileNames,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
                    out string[] arDocumentDescriptions,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
                    out string[] arDocumentUpdateDates,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
                    out string[] arDocumentVersions
    );

    [DllImport("PWSearchWrapper.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall)]
    private extern static int SearchForDocumentsWithStatesAndDatesAndReturnColumns(
        int iProjectId,
        bool bIncludeSubFolders,
        string sFullTextSearchString,
        bool bWholePhrase,
        bool bAnyWord,
        bool bSearchAttributes, // otherwise full text
        string sDocumentName,
        string sFileName,
        string sDocumentDescription,
        bool bOriginalsOnly,
        int iEnvironmentId,
        [In][MarshalAsAttribute(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] arAttributeNames,
        [In][MarshalAsAttribute(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] arAttributeValues,
        int size,
        string sStates, // comma delimited state numbers
        string sFileUpdatedAfter, // early date 2009-10-22 01:00:00
        string sFileUpdatedBefore, // late date 2010-10-22 01:00:00
        string sDocUpdatedAfter, // early date 2009-10-22 01:00:00 // added 2020-04-02
        string sDocUpdatedBefore, // late date 2010-10-22 01:00:00 // added 2020-04-02
        [In][MarshalAsAttribute(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] arReturnColumns,
        int iReturnColumnsSize,
        string sDelimiter,
        [Out] out IntPtr ppProjects,
        [Out] out IntPtr ppDocumentIds,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentGuidStrings,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentNames,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentFileNames,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentDescriptions,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentUpdateDates,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentVersions,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arAttributeColumnValues
        );

    [DllImport("PWSearchWrapperX64.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall, EntryPoint = "SearchForDocumentsUltimate")]
    private extern static int SearchForDocumentsUltimateX64
    (
        int iProjectId,
        bool bIncludeSubVaults,
        string wcFullTextString,
        bool bWholePhrase,
        bool bAnyWord, // otherwise all words
        bool bSearchAttributes, // otherwise full text at this point
        string szDocumentNameP,
        string szFileNameP,
        string szDocumentDescP,
        bool bOriginalsOnly,
        [In][MarshalAsAttribute(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] arAttributeNames,
        [In][MarshalAsAttribute(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] arAttributeValues,
        int iAttributesLength,
        string szEnvironmentsP, // comma delimited environment ids
        string szStatesP, // comma delimited state ids
        string szWorkflowsP, // comma delimited workflow ids
        string szFileUpdatedAfterP, // early date 2009-10-22 01:00:00
        string szFileUpdatedBeforeP, // late date 2010-10-22 01:00:00
        string szDocUpdatedAfterP, // early date 2009-10-22 01:00:00 // added 2020-04-02
        string szDocUpdatedBeforeP, // late date 2010-10-22 01:00:00 // added 2020-04-02
        string szDocCreatedAfterP, // early date 2009-10-22 01:00:00
        string szDocCreatedBeforeP, // late date 2010-10-22 01:00:00
        string szDocCheckedOutAfterP, // early date 2009-10-22 01:00:00
        string szDocCheckedOutBeforeP, // late date 2010-10-22 01:00:00
        double dLatitudeMin,
        double dintitudeMin,
        double dLatitudeMax,
        double dintitudeMax,
        bool bSpatialSecondPass,
        string szStoragesP, // comma delimited storage ids
        string szStatusesP, // comma delimited statuses ('CO','CI','I') etc.
        string szItemTypesP, // comma delimited types (12 flat set, 15 abstract, 0 and 10 are normal, 13 redline).
        int iFinalStatus, // -1 ignore, 0 not final status, 1 final status
        string szCreatorIdsP, // comma delimited creator ids
        string szLastUserIdsP, // comma delimited last user ids
        string szCheckedOutUserIdsP, // comma delimited checked out user ids
        string szApplicationIdsP, // comma delimited application ids
        [In][MarshalAsAttribute(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] arReturnColumns,
        int iReturnColumnsSize,
        string szDelimiterP, // delimiter for additional column values
        string sQueryName,
        int iQueryParentProjectId,
        int iQueryParentId,
        [Out] out IntPtr  ppProjects,
        [Out] out IntPtr  ppDocumentIds,
        [Out] out IntPtr  ppVersionSeqNumbers,
        [Out] out IntPtr  ppOriginalNumbers,
        [Out] out IntPtr  ppWorkflowIds,
        [Out] out IntPtr  ppStateIds,
        [Out] out IntPtr  ppStorageIds,
        [Out] out IntPtr  ppCreatorIds,
        [Out] out IntPtr  ppUpdaterIds,
        [Out] out IntPtr  ppCheckedOutUserIds,
        [Out] out IntPtr  ppItemTypes, // normal, set or abstract
        [Out] out IntPtr  ppApplicationIds,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentGuidStrings,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentNames,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentFileNames,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentDescriptions,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentUpdateDates,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arFileUpdateDates,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentCreateDates,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentCheckedOutDates,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentVersions,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentFileSizes,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentStatuses,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentMimeTypes,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arAttribeColumnValues
    );

    [DllImport("PWSearchWrapper.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall, EntryPoint = "SearchForDocumentsUltimate")]
    private extern static int SearchForDocumentsUltimate
    (
        int iProjectId,
        bool bIncludeSubVaults,
        string wcFullTextString,
        bool bWholePhrase,
        bool bAnyWord, // otherwise all words
        bool bSearchAttributes, // otherwise full text at this point
        string szDocumentNameP,
        string szFileNameP,
        string szDocumentDescP,
        bool bOriginalsOnly,
        [In][MarshalAsAttribute(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] arAttributeNames,
        [In][MarshalAsAttribute(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] arAttributeValues,
        int iAttributesLength,
        string szEnvironmentsP, // comma delimited environment ids
        string szStatesP, // comma delimited state ids
        string szWorkflowsP, // comma delimited workflow ids
        string szFileUpdatedAfterP, // early date 2009-10-22 01:00:00
        string szFileUpdatedBeforeP, // late date 2010-10-22 01:00:00
        string szDocUpdatedAfterP, // early date 2009-10-22 01:00:00 // added 2020-04-02
        string szDocUpdatedBeforeP, // late date 2010-10-22 01:00:00 // added 2020-04-02
        string szDocCreatedAfterP, // early date 2009-10-22 01:00:00
        string szDocCreatedBeforeP, // late date 2010-10-22 01:00:00
        string szDocCheckedOutAfterP, // early date 2009-10-22 01:00:00
        string szDocCheckedOutBeforeP, // late date 2010-10-22 01:00:00
        double dLatitudeMin,
        double dintitudeMin,
        double dLatitudeMax,
        double dintitudeMax,
        bool bSpatialSecondPass,
        string szStoragesP, // comma delimited storage ids
        string szStatusesP, // comma delimited statuses ('CO','CI','I') etc.
        string szItemTypesP, // comma delimited types (12 flat set, 15 abstract, 0 and 10 are normal, 13 redline).
        int iFinalStatus, // -1 ignore, 0 not final status, 1 final status
        string szCreatorIdsP, // comma delimited creator ids
        string szLastUserIdsP, // comma delimited last user ids
        string szCheckedOutUserIdsP, // comma delimited checked out user ids
        string szApplicationIdsP, // comma delimited application ids
        [In][MarshalAsAttribute(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] arReturnColumns,
        int iReturnColumnsSize,
        string szDelimiterP, // delimiter for additional column values
        string sQueryName,
        int iQueryParentProjectId,
        int iQueryParentId,
        [Out] out IntPtr ppProjects,
        [Out] out IntPtr ppDocumentIds,
        [Out] out IntPtr ppVersionSeqNumbers,
        [Out] out IntPtr ppOriginalNumbers,
        [Out] out IntPtr ppWorkflowIds,
        [Out] out IntPtr ppStateIds,
        [Out] out IntPtr ppStorageIds,
        [Out] out IntPtr ppCreatorIds,
        [Out] out IntPtr ppUpdaterIds,
        [Out] out IntPtr ppCheckedOutUserIds,
        [Out] out IntPtr ppItemTypes, // normal, set or abstract
        [Out] out IntPtr ppApplicationIds,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentGuidStrings,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentNames,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentFileNames,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentDescriptions,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentUpdateDates,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arFileUpdateDates,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentCreateDates,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentCheckedOutDates,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentVersions,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentFileSizes,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentStatuses,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentMimeTypes,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arAttribeColumnValues
    );

    [DllImport("PWSearchWrapperX64.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall, EntryPoint = "SearchForDocumentsWithStatesAndDatesAndReturnColumns")]
    private extern static int SearchForDocumentsWithStatesAndDatesAndReturnColumnsX64(
        int iProjectId,
        bool bIncludeSubFolders,
        string sFullTextSearchString,
        bool bWholePhrase,
        bool bAnyWord,
        bool bSearchAttributes, // otherwise full text
        string sDocumentName,
        string sFileName,
        string sDocumentDescription,
        bool bOriginalsOnly,
        int iEnvironmentId,
        [In][MarshalAsAttribute(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] arAttributeNames,
        [In][MarshalAsAttribute(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] arAttributeValues,
        int size,
        string sStates, // comma delimited state numbers
        string sFileUpdatedAfter, // early date 2009-10-22 01:00:00
        string sFileUpdatedBefore, // late date 2010-10-22 01:00:00
        string sDocUpdatedAfter, // early date 2009-10-22 01:00:00 // added 2020-04-02
        string sDocUpdatedBefore, // late date 2010-10-22 01:00:00 // added 2020-04-02
        [In][MarshalAsAttribute(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] arReturnColumns,
        int iReturnColumnsSize,
        string sDelimiter,
        [Out] out IntPtr ppProjects,
        [Out] out IntPtr ppDocumentIds,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentGuidStrings,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentNames,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentFileNames,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentDescriptions,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentUpdateDates,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentVersions,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arAttributeColumnValues
        );

    [DllImport("PWSearchWrapper.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall)]
    private extern static int SearchForDocumentsWithStatesAndUpdateDatesAndReturnColumns(
        int iProjectId,
        bool bIncludeSubFolders,
        string sFullTextSearchString,
        bool bWholePhrase,
        bool bAnyWord,
        bool bSearchAttributes, // otherwise full text
        string sDocumentName,
        string sFileName,
        string sDocumentDescription,
        bool bOriginalsOnly,
        int iEnvironmentId,
        [In][MarshalAsAttribute(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] arAttributeNames,
        [In][MarshalAsAttribute(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] arAttributeValues,
        int size,
        string sStates, // comma delimited state numbers
        string sDocUpdatedAfter, // early date 2009-10-22 01:00:00
        string sDocUpdatedBefore, // late date 2010-10-22 01:00:00
        [In][MarshalAsAttribute(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] arReturnColumns,
        int iReturnColumnsSize,
        string sDelimiter,
        [Out] out IntPtr ppProjects,
        [Out] out IntPtr ppDocumentIds,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentGuidStrings,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentNames,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentFileNames,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentDescriptions,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentUpdateDates,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentVersions,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arAttributeColumnValues
        );

    [DllImport("PWSearchWrapperX64.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall, EntryPoint = "SearchForDocumentsWithStatesAndUpdateDatesAndReturnColumns")]
    private extern static int SearchForDocumentsWithStatesAndUpdateDatesAndReturnColumnsX64(
        int iProjectId,
        bool bIncludeSubFolders,
        string sFullTextSearchString,
        bool bWholePhrase,
        bool bAnyWord,
        bool bSearchAttributes, // otherwise full text
        string sDocumentName,
        string sFileName,
        string sDocumentDescription,
        bool bOriginalsOnly,
        int iEnvironmentId,
        [In][MarshalAsAttribute(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] arAttributeNames,
        [In][MarshalAsAttribute(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] arAttributeValues,
        int size,
        string sStates, // comma delimited state numbers
        string sDocUpdatedAfter, // early date 2009-10-22 01:00:00
        string sDocUpdatedBefore, // late date 2010-10-22 01:00:00
        [In][MarshalAsAttribute(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] arReturnColumns,
        int iReturnColumnsSize,
        string sDelimiter,
        [Out] out IntPtr ppProjects,
        [Out] out IntPtr ppDocumentIds,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentGuidStrings,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentNames,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentFileNames,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentDescriptions,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentUpdateDates,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentVersions,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arAttributeColumnValues
        );

    [DllImport("PWSearchWrapper.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall)]
    private extern static int SearchForDocumentsByQueryId(
        int iQueryId,
        [Out] out IntPtr ppProjects,
        [Out] out IntPtr ppDocumentIds,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentGuidStrings,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentNames,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentFileNames,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentDescriptions,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentUpdateDates,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arVersions
        );

    [DllImport("PWSearchWrapperX64.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall, EntryPoint = "SearchForDocumentsByQueryId")]
    private extern static int SearchForDocumentsByQueryIdX64(
        int iQueryId,
        [Out] out IntPtr ppProjects,
        [Out] out IntPtr ppDocumentIds,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentGuidStrings,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentNames,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentFileNames,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentDescriptions,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentUpdateDates,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arVersions
        );

    [DllImport("PWSearchWrapper.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall)]
    private extern static int SearchForDocumentsMinMemory(
        int iProjectId,
        bool bIncludeSubFolders,
        string sFullTextSearchString,
        bool bWholePhrase,
        bool bAnyWord,
        bool bSearchAttributes, // otherwise full text
        string sDocumentName,
        string sFileName,
        string sDocumentDescription,
        bool bOriginalsOnly,
        int iEnvironmentId,
        [In][MarshalAsAttribute(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] arAttributeNames,
        [In][MarshalAsAttribute(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] arAttributeValues,
        int size,
        [Out] out IntPtr ppProjects,
        [Out] out IntPtr ppDocumentIds
        );

    [DllImport("PWSearchWrapperX64.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall, EntryPoint = "SearchForDocumentsMinMemory")]
    private extern static int SearchForDocumentsMinMemoryX64(
        int iProjectId,
        bool bIncludeSubFolders,
        string sFullTextSearchString,
        bool bWholePhrase,
        bool bAnyWord,
        bool bSearchAttributes, // otherwise full text
        string sDocumentName,
        string sFileName,
        string sDocumentDescription,
        bool bOriginalsOnly,
        int iEnvironmentId,
        [In][MarshalAsAttribute(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] arAttributeNames,
        [In][MarshalAsAttribute(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] arAttributeValues,
        int size,
        [Out] out IntPtr ppProjects,
        [Out] out IntPtr ppDocumentIds
        );

    [DllImport("PWSearchWrapper.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall)]
    private extern static int SearchForDocumentsByQueryIdMinMemory(
        int iQueryId,
        [Out] out IntPtr ppProjects,
        [Out] out IntPtr ppDocumentIds);

    [DllImport("PWSearchWrapperX64.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall, EntryPoint = "SearchForDocumentsByQueryIdMinMemory")]
    private extern static int SearchForDocumentsByQueryIdMinMemoryX64(
        int iQueryId,
        [Out] out IntPtr ppProjects,
        [Out] out IntPtr ppDocumentIds);

    public static bool Is64Bit()
    {
        return (IntPtr.Size == 8);
    }

    private static void MarshalUnmananagedIntArrayToManagedIntArray
    (
      IntPtr pUnmanagedIntArray,
      int iCount,
      out int[] ManagedIntArray
    )
    {
        ManagedIntArray = new int[iCount];

        if (pUnmanagedIntArray != IntPtr.Zero)
        {
            Marshal.Copy(pUnmanagedIntArray, ManagedIntArray, 0, iCount);

            Marshal.FreeCoTaskMem(pUnmanagedIntArray);
        }
    }

    public static SortedList<int, string> GetListOfRichProjectsFromProperties(int iParentProjectId, string sClassType, SortedList<string, string> slProps, bool bGetPath)
    {
        SortedList<int, string> slProjects = new SortedList<int, string>();

        string[] saPropNames = new string[slProps.Count];
        string[] saPropValues = new string[slProps.Count];

        slProps.Keys.CopyTo(saPropNames, 0);

        slProps.Values.CopyTo(saPropValues, 0);

        IntPtr intPtrProjects = IntPtr.Zero;

        int iCount = 0;

        try
        {
            if (Is64Bit())
            {
                SearchForProjectsByPropertiesX64(iParentProjectId, sClassType, saPropNames, saPropValues, slProps.Count, out intPtrProjects, out iCount);
            }
            else
            {
                SearchForProjectsByProperties(iParentProjectId, sClassType, saPropNames, saPropValues, slProps.Count, out intPtrProjects, out iCount);
            }

            if (iCount > 0 && intPtrProjects != IntPtr.Zero)
            {
                int[] arProjects = new int[iCount];

                Marshal.Copy(intPtrProjects, arProjects, 0, iCount);

                foreach (int iProject in arProjects)
                {
                    if (!slProjects.ContainsKey(iProject))
                    {
                        if (bGetPath)
                        {
                            slProjects.Add(iProject, PWWrapper.GetProjectNamePath(iProject));
                        }
                        else
                        {
                            slProjects.Add(iProject, string.Empty);
                        }
                    }
                }

                Marshal.FreeCoTaskMem(intPtrProjects);
            }
        }
        catch (Exception ex)
        {
            BPSUtilities.WriteLog("Error: {0}\n{1}", ex.Message, ex.StackTrace);
        }

        return slProjects;
    }

    public static SortedList<int, string> GetListOfProjectsFromName(string sNamePattern, bool bGetPath)
    {
        SortedList<int, string> slProjects = new SortedList<int, string>();

        IntPtr intPtrProjects = IntPtr.Zero;

        int iCount = 0;

        if (Is64Bit())
            SearchForProjectsByNameX64(sNamePattern, out intPtrProjects, out iCount);
        else
            SearchForProjectsByName(sNamePattern, out intPtrProjects, out iCount);

        if (iCount > 0 && intPtrProjects != IntPtr.Zero)
        {
            int[] arProjects = new int[iCount];

            Marshal.Copy(intPtrProjects, arProjects, 0, iCount);

            // BPSUtilities.WriteLog("Found {0} projects", iCount);

            foreach (int iProject in arProjects)
            {
                if (!slProjects.ContainsKey(iProject))
                {
                    if (bGetPath)
                    {
                        slProjects.Add(iProject, PWWrapper.GetProjectNamePath(iProject));
                    }
                    else
                    {
                        slProjects.Add(iProject, string.Empty);
                    }
                }
            }

            Marshal.FreeCoTaskMem(intPtrProjects);
        }

        return slProjects;
    }

    public static DataTable GetAllProjectsByClass(int iParentProjectId, string sClassName, bool bGetPath)
    {
        IntPtr ppProjects = IntPtr.Zero;
        IntPtr ppComponentClassIds = IntPtr.Zero;
        IntPtr ppComponentInstanceIds = IntPtr.Zero;
        IntPtr ppEnvironmentIds = IntPtr.Zero;
        IntPtr ppWorkflowIds = IntPtr.Zero;
        IntPtr ppParentIds = IntPtr.Zero;
        IntPtr ppStorageIds = IntPtr.Zero;
        IntPtr ppIsParent = IntPtr.Zero;
        IntPtr ppProjectTypes = IntPtr.Zero;
        IntPtr pppProjectNames = IntPtr.Zero;
        IntPtr pppProjectDescriptions = IntPtr.Zero;
        IntPtr pppProjectCodes = IntPtr.Zero;

        int iCount = 0;

        string[] arProjectNames = null;
        string[] arProjectDescriptions = null;
        string[] arProjectCodes = null;

        int iClassId = 0;

        if (!string.IsNullOrEmpty(sClassName))
            iClassId = PWWrapper.GetClassIdFromClassName(sClassName);

        if (Is64Bit())
            SearchForProjectsByClassIdX64(iParentProjectId, iClassId, out ppProjects, out ppComponentClassIds,
                out ppComponentInstanceIds, out ppEnvironmentIds, out ppWorkflowIds,
                out ppParentIds, out ppStorageIds,
                out ppIsParent, out ppProjectTypes, out arProjectNames, out arProjectDescriptions, out arProjectCodes,
                out iCount);
        else
            SearchForProjectsByClassId(iParentProjectId, iClassId, out ppProjects, out ppComponentClassIds,
                out ppComponentInstanceIds, out ppEnvironmentIds, out ppWorkflowIds,
                out ppParentIds, out ppStorageIds,
                out ppIsParent, out ppProjectTypes, out arProjectNames, out arProjectDescriptions, out arProjectCodes,
                out iCount);


        DataTable dt = new DataTable("Projects");

        dt.Columns.Add("ProjectName", Type.GetType("System.String"));
        dt.Columns.Add("ProjectDescription", Type.GetType("System.String"));
        dt.Columns.Add("ProjectCode", Type.GetType("System.String"));
        dt.Columns.Add("ProjectID", Type.GetType("System.Int32"));
        dt.Columns.Add("ClassID", Type.GetType("System.Int32"));
        dt.Columns.Add("InstanceID", Type.GetType("System.Int32"));
        dt.Columns.Add("EnvironmentID", Type.GetType("System.Int32"));
        dt.Columns.Add("WorkflowID", Type.GetType("System.Int32"));
        dt.Columns.Add("ParentID", Type.GetType("System.Int32"));
        dt.Columns.Add("StorageID", Type.GetType("System.Int32"));
        dt.Columns.Add("IsParent", Type.GetType("System.Int32"));
        dt.Columns.Add("PWPath", Type.GetType("System.String"));
        dt.Columns.Add("ProjectType", Type.GetType("System.Int32"));

        DataColumn[] pk = new DataColumn[1];
        pk[0] = dt.Columns["ProjectId"];
        dt.PrimaryKey = pk;

        if (iCount > 0)
        {
            System.Diagnostics.Debug.WriteLine(string.Format("Found {0} projects", iCount));
            Console.WriteLine(string.Format("Found {0} projects", iCount));

            int[] arProjects = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppProjects, iCount, out arProjects);
            int[] arClassIds = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppComponentClassIds, iCount, out arClassIds);
            int[] arInstanceIds = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppComponentInstanceIds, iCount, out arInstanceIds);
            int[] arEnvironmentIds = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppEnvironmentIds, iCount, out arEnvironmentIds);
            int[] arWorkflowIds = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppWorkflowIds, iCount, out arWorkflowIds);
            int[] arParentIds = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppParentIds, iCount, out arParentIds);
            int[] arStorageIds = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppStorageIds, iCount, out arStorageIds);
            int[] arIsParent = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppIsParent, iCount, out arIsParent);
            int[] arProjectTypes = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppProjectTypes, iCount, out arProjectTypes);

            for (int i = 0; i < iCount; i++)
            {
                DataRow dr = dt.NewRow();

                dr["ProjectID"] = arProjects[i];
                dr["ClassID"] = arClassIds[i];
                dr["InstanceID"] = arInstanceIds[i];
                dr["EnvironmentID"] = arEnvironmentIds[i];
                dr["WorkflowID"] = arWorkflowIds[i];
                dr["ParentID"] = arParentIds[i];
                dr["StorageID"] = arStorageIds[i];
                dr["IsParent"] = arIsParent[i];
                dr["ProjectName"] = arProjectNames[i];
                dr["ProjectDescription"] = arProjectDescriptions[i];
                dr["ProjectCode"] = arProjectCodes[i];
                dr["ProjectType"] = arProjectTypes[i];
                // if (bGetPath)
                // dr["PWPath"] = PWWrapper.GetProjectNamePath(arProjects[i]);

                try
                {
                    dt.Rows.Add(dr);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(string.Format("Error: {0}\n{1}", ex.Message, ex.StackTrace));
                }
            }

            if (bGetPath)
            {
                System.Diagnostics.Debug.WriteLine("Filling paths...");
                Console.WriteLine("Filling paths...");

                foreach (DataRow dr in dt.Rows)
                {
                    StringBuilder sb = new StringBuilder();

                    sb.Append(dr["ProjectName"].ToString());

                    if ((int)dr["ParentID"] > 0)
                    {
                        DataRow drParent = dt.Rows.Find(dr["ParentID"]);

                        if (drParent != null)
                        {
                            sb.Insert(0, drParent["ProjectName"].ToString() + "\\");

                            while ((int)drParent["ParentID"] > 0)
                            {
                                drParent = dt.Rows.Find(drParent["ParentID"]);

                                if (drParent == null)
                                    break;

                                sb.Insert(0, drParent["ProjectName"].ToString() + "\\");
                            }
                        }
                    }

                    dr["PWPath"] = sb.ToString();
                } // foreach row in the table returned from the search
            } // if (bGetPath)
        }

        return dt;
    }

    public static DataTable GetAllProjectsInBranch(int iParentProjectId, bool bGetPath)
    {
        IntPtr ppProjects = IntPtr.Zero;
        IntPtr ppComponentClassIds = IntPtr.Zero;
        IntPtr ppComponentInstanceIds = IntPtr.Zero;
        IntPtr ppEnvironmentIds = IntPtr.Zero;
        IntPtr ppWorkflowIds = IntPtr.Zero;
        IntPtr ppParentIds = IntPtr.Zero;
        IntPtr ppStorageIds = IntPtr.Zero;
        IntPtr ppIsParent = IntPtr.Zero;
        IntPtr pppProjectNames = IntPtr.Zero;
        IntPtr pppProjectDescriptions = IntPtr.Zero;
        IntPtr pppProjectCodes = IntPtr.Zero;
        IntPtr ppProjectTypes = IntPtr.Zero;

        int iCount = 0;

        string[] arProjectNames = null;
        string[] arProjectDescriptions = null;
        string[] arProjectCodes = null;
        string[] arProjectGuids = null;

        if (Is64Bit())
            SearchForProjectsByTreeX64(iParentProjectId, out ppProjects, out ppComponentClassIds,
                out ppComponentInstanceIds, out ppEnvironmentIds, out ppWorkflowIds,
                out ppParentIds, out ppStorageIds,
                out ppIsParent, out ppProjectTypes, out arProjectNames, out arProjectDescriptions, out arProjectCodes,
                out arProjectGuids,
                out iCount);
        else
            SearchForProjectsByTree(iParentProjectId, out ppProjects, out ppComponentClassIds,
                out ppComponentInstanceIds, out ppEnvironmentIds, out ppWorkflowIds,
                out ppParentIds, out ppStorageIds,
                out ppIsParent, out ppProjectTypes, out arProjectNames, out arProjectDescriptions, out arProjectCodes,
                out arProjectGuids,
                out iCount);

        DataTable dt = new DataTable("Projects");

        dt.Columns.Add("ProjectName", Type.GetType("System.String"));
        dt.Columns.Add("ProjectDescription", Type.GetType("System.String"));
        dt.Columns.Add("ProjectCode", Type.GetType("System.String"));
        dt.Columns.Add("ProjectID", Type.GetType("System.Int32"));
        dt.Columns.Add("ProjectGUID", Type.GetType("System.String"));
        dt.Columns.Add("ClassID", Type.GetType("System.Int32"));
        dt.Columns.Add("InstanceID", Type.GetType("System.Int32"));
        dt.Columns.Add("EnvironmentID", Type.GetType("System.Int32"));
        dt.Columns.Add("WorkflowID", Type.GetType("System.Int32"));
        dt.Columns.Add("ParentID", Type.GetType("System.Int32"));
        dt.Columns.Add("StorageID", Type.GetType("System.Int32"));
        dt.Columns.Add("IsParent", Type.GetType("System.Int32"));
        dt.Columns.Add("PWPath", Type.GetType("System.String"));
        dt.Columns.Add("ProjectType", Type.GetType("System.Int32"));

        DataColumn[] pk = new DataColumn[1];
        pk[0] = dt.Columns["ProjectId"];
        dt.PrimaryKey = pk;

        if (iCount > 0)
        {
            // System.Diagnostics.Debug.WriteLine(string.Format("Found {0} projects", iCount));
            // Console.WriteLine(string.Format("Found {0} projects", iCount));

            int[] arProjects = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppProjects, iCount, out arProjects);
            int[] arClassIds = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppComponentClassIds, iCount, out arClassIds);
            int[] arInstanceIds = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppComponentInstanceIds, iCount, out arInstanceIds);
            int[] arEnvironmentIds = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppEnvironmentIds, iCount, out arEnvironmentIds);
            int[] arWorkflowIds = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppWorkflowIds, iCount, out arWorkflowIds);
            int[] arParentIds = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppParentIds, iCount, out arParentIds);
            int[] arStorageIds = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppStorageIds, iCount, out arStorageIds);
            int[] arIsParent = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppIsParent, iCount, out arIsParent);
            int[] arProjectTypes = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppProjectTypes, iCount, out arProjectTypes);

            for (int i = 0; i < iCount; i++)
            {
                DataRow dr = dt.NewRow();

                dr["ProjectID"] = arProjects[i];
                dr["ClassID"] = arClassIds[i];
                dr["InstanceID"] = arInstanceIds[i];
                dr["EnvironmentID"] = arEnvironmentIds[i];
                dr["WorkflowID"] = arWorkflowIds[i];
                dr["ParentID"] = arParentIds[i];
                dr["StorageID"] = arStorageIds[i];
                dr["IsParent"] = arIsParent[i];
                dr["ProjectName"] = arProjectNames[i];
                dr["ProjectDescription"] = arProjectDescriptions[i];
                dr["ProjectCode"] = arProjectCodes[i];
                dr["ProjectType"] = arProjectTypes[i];
                dr["ProjectGUID"] = arProjectGuids[i];
                // if (bGetPath)
                // dr["PWPath"] = PWWrapper.GetProjectNamePath(arProjects[i]);

                try
                {
                    dt.Rows.Add(dr);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(string.Format("Error: {0}\n{1}", ex.Message, ex.StackTrace));
                }
            }

            if (bGetPath)
            {
                // System.Diagnostics.Debug.WriteLine("Filling paths...");
                // Console.WriteLine("Filling paths...");

                foreach (DataRow dr in dt.Rows)
                {
                    StringBuilder sb = new StringBuilder();

                    sb.Append(dr["ProjectName"].ToString());

                    if ((int)dr["ParentID"] > 0)
                    {
                        DataRow drParent = dt.Rows.Find(dr["ParentID"]);

                        if (drParent != null)
                        {
                            sb.Insert(0, drParent["ProjectName"].ToString() + "\\");

                            while ((int)drParent["ParentID"] > 0)
                            {
                                drParent = dt.Rows.Find(drParent["ParentID"]);

                                if (drParent == null)
                                    break;

                                sb.Insert(0, drParent["ProjectName"].ToString() + "\\");
                            }
                        }
                    }

                    dr["PWPath"] = sb.ToString();
                } // foreach row in the table returned from the search
            } // if (bGetPath)
        }

        return dt;
    }

    public static DataTable GetAllProjectsInBranchWithEnvironment(int iParentProjectId, int iEnvironmentId, bool bGetPath)
    {
        IntPtr ppProjects = IntPtr.Zero;
        IntPtr ppComponentClassIds = IntPtr.Zero;
        IntPtr ppComponentInstanceIds = IntPtr.Zero;
        IntPtr ppEnvironmentIds = IntPtr.Zero;
        IntPtr ppWorkflowIds = IntPtr.Zero;
        IntPtr ppParentIds = IntPtr.Zero;
        IntPtr ppStorageIds = IntPtr.Zero;
        IntPtr ppIsParent = IntPtr.Zero;
        IntPtr pppProjectNames = IntPtr.Zero;
        IntPtr pppProjectDescriptions = IntPtr.Zero;
        IntPtr pppProjectCodes = IntPtr.Zero;
        IntPtr ppProjectTypes = IntPtr.Zero;

        int iCount = 0;

        string[] arProjectNames = null;
        string[] arProjectDescriptions = null;
        string[] arProjectCodes = null;

        if (Is64Bit())
            SearchForProjectsByTreeAndEnvironmentX64(iParentProjectId, iEnvironmentId, out ppProjects, out ppComponentClassIds,
                out ppComponentInstanceIds, out ppEnvironmentIds, out ppWorkflowIds,
                out ppParentIds, out ppStorageIds,
                out ppIsParent, out ppProjectTypes, out arProjectNames, out arProjectDescriptions, out arProjectCodes,
                out iCount);
        else
            SearchForProjectsByTreeAndEnvironment(iParentProjectId, iEnvironmentId, out ppProjects, out ppComponentClassIds,
                out ppComponentInstanceIds, out ppEnvironmentIds, out ppWorkflowIds,
                out ppParentIds, out ppStorageIds,
                out ppIsParent, out ppProjectTypes, out arProjectNames, out arProjectDescriptions, out arProjectCodes,
                out iCount);

        DataTable dt = new DataTable("Projects");

        dt.Columns.Add("ProjectName", Type.GetType("System.String"));
        dt.Columns.Add("ProjectDescription", Type.GetType("System.String"));
        dt.Columns.Add("ProjectCode", Type.GetType("System.String"));
        dt.Columns.Add("ProjectID", Type.GetType("System.Int32"));
        dt.Columns.Add("ClassID", Type.GetType("System.Int32"));
        dt.Columns.Add("InstanceID", Type.GetType("System.Int32"));
        dt.Columns.Add("EnvironmentID", Type.GetType("System.Int32"));
        dt.Columns.Add("WorkflowID", Type.GetType("System.Int32"));
        dt.Columns.Add("ParentID", Type.GetType("System.Int32"));
        dt.Columns.Add("StorageID", Type.GetType("System.Int32"));
        dt.Columns.Add("IsParent", Type.GetType("System.Int32"));
        dt.Columns.Add("PWPath", Type.GetType("System.String"));
        dt.Columns.Add("ProjectType", Type.GetType("System.Int32"));

        DataColumn[] pk = new DataColumn[1];
        pk[0] = dt.Columns["ProjectId"];
        dt.PrimaryKey = pk;

        if (iCount > 0)
        {
            System.Diagnostics.Debug.WriteLine(string.Format("Found {0} projects", iCount));
            Console.WriteLine(string.Format("Found {0} projects", iCount));

            int[] arProjects = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppProjects, iCount, out arProjects);
            int[] arClassIds = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppComponentClassIds, iCount, out arClassIds);
            int[] arInstanceIds = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppComponentInstanceIds, iCount, out arInstanceIds);
            int[] arEnvironmentIds = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppEnvironmentIds, iCount, out arEnvironmentIds);
            int[] arWorkflowIds = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppWorkflowIds, iCount, out arWorkflowIds);
            int[] arParentIds = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppParentIds, iCount, out arParentIds);
            int[] arStorageIds = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppStorageIds, iCount, out arStorageIds);
            int[] arIsParent = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppIsParent, iCount, out arIsParent);
            int[] arProjectTypes = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppProjectTypes, iCount, out arProjectTypes);

            for (int i = 0; i < iCount; i++)
            {
                DataRow dr = dt.NewRow();

                dr["ProjectID"] = arProjects[i];
                dr["ClassID"] = arClassIds[i];
                dr["InstanceID"] = arInstanceIds[i];
                dr["EnvironmentID"] = arEnvironmentIds[i];
                dr["WorkflowID"] = arWorkflowIds[i];
                dr["ParentID"] = arParentIds[i];
                dr["StorageID"] = arStorageIds[i];
                dr["IsParent"] = arIsParent[i];
                dr["ProjectName"] = arProjectNames[i];
                dr["ProjectDescription"] = arProjectDescriptions[i];
                dr["ProjectCode"] = arProjectCodes[i];
                dr["ProjectType"] = arProjectTypes[i];
                // if (bGetPath)
                // dr["PWPath"] = PWWrapper.GetProjectNamePath(arProjects[i]);

                try
                {
                    dt.Rows.Add(dr);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(string.Format("Error: {0}\n{1}", ex.Message, ex.StackTrace));
                }
            }

            if (bGetPath)
            {
                System.Diagnostics.Debug.WriteLine("Filling paths...");
                Console.WriteLine("Filling paths...");

                foreach (DataRow dr in dt.Rows)
                {
                    StringBuilder sb = new StringBuilder();

                    sb.Append(dr["ProjectName"].ToString());

                    if ((int)dr["ParentID"] > 0)
                    {
                        DataRow drParent = dt.Rows.Find(dr["ParentID"]);

                        if (drParent != null)
                        {
                            sb.Insert(0, drParent["ProjectName"].ToString() + "\\");

                            while ((int)drParent["ParentID"] > 0)
                            {
                                drParent = dt.Rows.Find(drParent["ParentID"]);

                                if (drParent == null)
                                    break;

                                sb.Insert(0, drParent["ProjectName"].ToString() + "\\");
                            }
                        }
                    }

                    dr["PWPath"] = sb.ToString();
                } // foreach row in the table returned from the search
            } // if (bGetPath)
        }

        return dt;
    }

    public static DataTable GetAllEmptyProjectsInBranch(int iParentProjectId, bool bGetPath)
    {
        IntPtr ppProjects = IntPtr.Zero;
        IntPtr ppComponentClassIds = IntPtr.Zero;
        IntPtr ppComponentInstanceIds = IntPtr.Zero;
        IntPtr ppEnvironmentIds = IntPtr.Zero;
        IntPtr ppWorkflowIds = IntPtr.Zero;
        IntPtr ppParentIds = IntPtr.Zero;
        IntPtr ppStorageIds = IntPtr.Zero;
        IntPtr ppIsParent = IntPtr.Zero;
        IntPtr pppProjectNames = IntPtr.Zero;
        IntPtr pppProjectDescriptions = IntPtr.Zero;
        IntPtr pppProjectCodes = IntPtr.Zero;
        IntPtr ppProjectTypes = IntPtr.Zero;

        int iCount = 0;

        string[] arProjectNames = null;
        string[] arProjectDescriptions = null;
        string[] arProjectCodes = null;

        if (Is64Bit())
            SearchForProjectsByTreeForEmptyFoldersX64(iParentProjectId, out ppProjects, out ppComponentClassIds,
                out ppComponentInstanceIds, out ppEnvironmentIds, out ppWorkflowIds,
                out ppParentIds, out ppStorageIds,
                out ppIsParent, out ppProjectTypes, out arProjectNames, out arProjectDescriptions, out arProjectCodes,
                out iCount);
        else
            SearchForProjectsByTreeForEmptyFolders(iParentProjectId, out ppProjects, out ppComponentClassIds,
                out ppComponentInstanceIds, out ppEnvironmentIds, out ppWorkflowIds,
                out ppParentIds, out ppStorageIds,
                out ppIsParent, out ppProjectTypes, out arProjectNames, out arProjectDescriptions, out arProjectCodes,
                out iCount);

        DataTable dt = new DataTable("Projects");

        dt.Columns.Add("ProjectName", Type.GetType("System.String"));
        dt.Columns.Add("ProjectDescription", Type.GetType("System.String"));
        dt.Columns.Add("ProjectCode", Type.GetType("System.String"));
        dt.Columns.Add("ProjectID", Type.GetType("System.Int32"));
        dt.Columns.Add("ClassID", Type.GetType("System.Int32"));
        dt.Columns.Add("InstanceID", Type.GetType("System.Int32"));
        dt.Columns.Add("EnvironmentID", Type.GetType("System.Int32"));
        dt.Columns.Add("WorkflowID", Type.GetType("System.Int32"));
        dt.Columns.Add("ParentID", Type.GetType("System.Int32"));
        dt.Columns.Add("StorageID", Type.GetType("System.Int32"));
        dt.Columns.Add("IsParent", Type.GetType("System.Int32"));
        dt.Columns.Add("PWPath", Type.GetType("System.String"));
        dt.Columns.Add("ProjectType", Type.GetType("System.Int32"));

        DataColumn[] pk = new DataColumn[1];
        pk[0] = dt.Columns["ProjectId"];
        dt.PrimaryKey = pk;

        if (iCount > 0)
        {
            System.Diagnostics.Debug.WriteLine(string.Format("Found {0} projects", iCount));
            Console.WriteLine(string.Format("Found {0} projects", iCount));

            int[] arProjects = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppProjects, iCount, out arProjects);
            int[] arClassIds = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppComponentClassIds, iCount, out arClassIds);
            int[] arInstanceIds = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppComponentInstanceIds, iCount, out arInstanceIds);
            int[] arEnvironmentIds = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppEnvironmentIds, iCount, out arEnvironmentIds);
            int[] arWorkflowIds = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppWorkflowIds, iCount, out arWorkflowIds);
            int[] arParentIds = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppParentIds, iCount, out arParentIds);
            int[] arStorageIds = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppStorageIds, iCount, out arStorageIds);
            int[] arIsParent = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppIsParent, iCount, out arIsParent);
            int[] arProjectTypes = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppProjectTypes, iCount, out arProjectTypes);

            for (int i = 0; i < iCount; i++)
            {
                DataRow dr = dt.NewRow();

                dr["ProjectID"] = arProjects[i];
                dr["ClassID"] = arClassIds[i];
                dr["InstanceID"] = arInstanceIds[i];
                dr["EnvironmentID"] = arEnvironmentIds[i];
                dr["WorkflowID"] = arWorkflowIds[i];
                dr["ParentID"] = arParentIds[i];
                dr["StorageID"] = arStorageIds[i];
                dr["IsParent"] = arIsParent[i];
                dr["ProjectName"] = arProjectNames[i];
                dr["ProjectDescription"] = arProjectDescriptions[i];
                dr["ProjectCode"] = arProjectCodes[i];
                dr["ProjectType"] = arProjectTypes[i];

                // if (bGetPath)
                // dr["PWPath"] = PWWrapper.GetProjectNamePath(arProjects[i]);

                try
                {
                    dt.Rows.Add(dr);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(string.Format("Error: {0}\n{1}", ex.Message, ex.StackTrace));
                }
            }

            if (bGetPath)
            {
                System.Diagnostics.Debug.WriteLine("Filling paths...");
                Console.WriteLine("Filling paths...");

                foreach (DataRow dr in dt.Rows)
                {
                    StringBuilder sb = new StringBuilder();

                    sb.Append(dr["ProjectName"].ToString());

                    if ((int)dr["ParentID"] > 0)
                    {
                        DataRow drParent = dt.Rows.Find(dr["ParentID"]);

                        if (drParent != null)
                        {
                            sb.Insert(0, drParent["ProjectName"].ToString() + "\\");

                            while ((int)drParent["ParentID"] > 0)
                            {
                                drParent = dt.Rows.Find(drParent["ParentID"]);

                                if (drParent == null)
                                    break;

                                sb.Insert(0, drParent["ProjectName"].ToString() + "\\");
                            }
                        }
                    }

                    dr["PWPath"] = sb.ToString();
                } // foreach row in the table returned from the search
            } // if (bGetPath)
        }

        return dt;
    }

    public static DataTable SearchForDocumentsWithStatesAndDates(int iProjectId, bool bSearchSubFolders,
        string sFullText, bool bWholePhrase, bool bAnyWords, bool bSearchAttributes,
        string sDocumentName, string sFileName, string sDocumentDescription, bool bOriginalsOnly,
        int iEnvironmentId,
        SortedList<string, string> slAttributes,
        List<int> listStates,
        string sUpdatedAfter, // early date 2009-10-22 01:00:00
        string sUpdatedBefore, // late date 2010-10-22 01:00:00
        bool bGetPath)
    {
        IntPtr ppProjects = IntPtr.Zero;
        IntPtr ppDocumentIds = IntPtr.Zero;
        IntPtr ppVersionSequenceNums = IntPtr.Zero;

        int iCount = 0;

        string[] arDocumentGuidStrings = null;
        string[] arDocumentNames = null;
        string[] arDocumentFileNames = null;
        string[] arDocumentDescriptions = null;
        string[] arDocumentUpdateDates = null;
        string[] arDocumentVersions = null;

        string[] arAttributeNames = new string[Math.Max(slAttributes.Count, 1)];
        string[] arAttributeValues = new string[Math.Max(slAttributes.Count, 1)];

        int iIndex = 0;

        foreach (KeyValuePair<string, string> kvp in slAttributes)
        {
            arAttributeNames[iIndex] = kvp.Key;
            arAttributeValues[iIndex] = kvp.Value;
            iIndex++;
        }

        System.Diagnostics.Debug.WriteLine("Starting query...");

        string sStates = string.Empty;

        StringBuilder sbStates = new StringBuilder();

        foreach (int iState in listStates)
        {
            if (sbStates.Length > 0)
                sbStates.Append(",");
            sbStates.Append(iState);
        }

        sStates = sbStates.ToString();

        try
        {
            if (Is64Bit())
                iCount = SearchForDocumentsWithStatesAndDatesX64(iProjectId, bSearchSubFolders, sFullText, bWholePhrase,
                    bAnyWords, bSearchAttributes, sDocumentName, sFileName, sDocumentDescription, bOriginalsOnly,
                    iEnvironmentId,
                    arAttributeNames, arAttributeValues, slAttributes.Count,
                    sStates,
                    sUpdatedAfter,
                    sUpdatedBefore,
                    out ppProjects, out ppDocumentIds, out ppVersionSequenceNums, out arDocumentGuidStrings, out arDocumentNames, out arDocumentFileNames,
                    out arDocumentDescriptions, out arDocumentUpdateDates, out arDocumentVersions);
            else
                iCount = SearchForDocumentsWithStatesAndDates(iProjectId, bSearchSubFolders, sFullText, bWholePhrase,
                    bAnyWords, bSearchAttributes, sDocumentName, sFileName, sDocumentDescription, bOriginalsOnly,
                    iEnvironmentId,
                    arAttributeNames, arAttributeValues, slAttributes.Count,
                    sStates,
                    sUpdatedAfter,
                    sUpdatedBefore,
                    out ppProjects, out ppDocumentIds, out ppVersionSequenceNums, out arDocumentGuidStrings, out arDocumentNames, out arDocumentFileNames,
                    out arDocumentDescriptions, out arDocumentUpdateDates, out arDocumentVersions);
        }
        catch (Exception ex)
        {
            BPSUtilities.WriteLog("Error: {0}\n{1}", ex.Message, ex.StackTrace);
        }

        System.Diagnostics.Debug.WriteLine(string.Format("Returned {0} matches", iCount));

        DataTable dt = new DataTable("Documents");

        dt.Columns.Add("DocumentGUID", Type.GetType("System.String"));
        dt.Columns.Add("DocumentName", Type.GetType("System.String"));
        dt.Columns.Add("DocumentFileName", Type.GetType("System.String"));
        dt.Columns.Add("DocumentDescription", Type.GetType("System.String"));
        dt.Columns.Add("DocumentUpdateDate", Type.GetType("System.String"));
        dt.Columns.Add("DocumentVersion", Type.GetType("System.String"));
        dt.Columns.Add("DocumentVersionSequence", Type.GetType("System.Int32"));
        dt.Columns.Add("ProjectId", Type.GetType("System.Int32"));
        dt.Columns.Add("DocumentId", Type.GetType("System.Int32"));
        dt.Columns.Add("ProjectPath", Type.GetType("System.String"));

        DataColumn[] pk = new DataColumn[1];
        pk[0] = dt.Columns["DocumentGUID"];
        dt.PrimaryKey = pk;

        SortedList<int, string> slProjectPaths = new SortedList<int, string>();

        if (iCount > 0)
        {
            int[] arProjects = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppProjects, iCount, out arProjects);
            int[] arDocumentIds = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppDocumentIds, iCount, out arDocumentIds);
            int[] arDocumentVersionSequenceNums = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppVersionSequenceNums, iCount, out arDocumentVersionSequenceNums);

            for (int i = 0; i < iCount; i++)
            {
                DataRow dr = dt.NewRow();

                dr["DocumentId"] = arDocumentIds[i];
                dr["ProjectID"] = arProjects[i];
                dr["DocumentGUID"] = arDocumentGuidStrings[i];
                dr["DocumentName"] = arDocumentNames[i];
                dr["DocumentFileName"] = arDocumentFileNames[i];
                dr["DocumentDescription"] = arDocumentDescriptions[i];
                dr["DocumentUpdateDate"] = arDocumentUpdateDates[i];
                dr["DocumentVersion"] = arDocumentVersions[i];
                dr["DocumentVersionSequence"] = arDocumentVersionSequenceNums[i];

                if (bGetPath)
                {
                    string sProjectPath = string.Empty;
                    if (!slProjectPaths.TryGetValue(arProjects[i], out sProjectPath))
                    {
                        sProjectPath = PWWrapper.GetProjectNamePath(arProjects[i]);
                    }

                    dr["ProjectPath"] = sProjectPath;
                }

                try
                {
                    dt.Rows.Add(dr);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(string.Format("Error: {0}\n{1}", ex.Message, ex.StackTrace));
                }
            }

            // because I always forget the column names
            StringBuilder sbColumns = new StringBuilder();

            foreach (DataColumn dc in dt.Columns)
            {
                sbColumns.Append(dc.ColumnName + ";");
            }

            System.Diagnostics.Debug.WriteLine(string.Format("Columns: {0}", sbColumns.ToString()));
        }

        GC.Collect();

        return dt;
    }

    public static DataTable SearchForDocumentsWithStatesAndDatesAndReturnColumns(int iProjectId, bool bSearchSubFolders,
        string sFullText, bool bWholePhrase, bool bAnyWords, bool bSearchAttributes,
        string sDocumentName, string sFileName, string sDocumentDescription, bool bOriginalsOnly,
        int iEnvironmentId,
        SortedList<string, string> slAttributes,
        List<int> listStates,
        string sFileUpdatedAfter, // early date 2009-10-22 01:00:00
        string sFileUpdatedBefore, // late date 2010-10-22 01:00:00
        string sDocUpdatedAfter, // early date 2009-10-22 01:00:00 // added 2020-04-02
        string sDocUpdatedBefore, // late date 2010-10-22 01:00:00 // added 2020-04-02
        bool bGetPath,
        List<string> listColumns)
    {
        IntPtr ppProjects = IntPtr.Zero;
        IntPtr ppDocumentIds = IntPtr.Zero;

        int iCount = 0;

        string[] arDocumentGuidStrings = null;
        string[] arDocumentNames = null;
        string[] arDocumentFileNames = null;
        string[] arDocumentDescriptions = null;
        string[] arDocumentUpdateDates = null;
        string[] arDocumentVersions = null;
        string[] arAdditionalAttributeValues = null;

        string[] arAttributeNames = new string[Math.Max(slAttributes.Count, 1)];
        string[] arAttributeValues = new string[Math.Max(slAttributes.Count, 1)];

        string[] arColumns = listColumns.ToArray();

        int iIndex = 0;

        foreach (KeyValuePair<string, string> kvp in slAttributes)
        {
            arAttributeNames[iIndex] = kvp.Key;
            arAttributeValues[iIndex] = kvp.Value;
            iIndex++;
        }

        System.Diagnostics.Debug.WriteLine("Starting query...");

        string sStates = string.Empty;

        StringBuilder sbStates = new StringBuilder();

        foreach (int iState in listStates)
        {
            if (sbStates.Length > 0)
                sbStates.Append(",");
            sbStates.Append(iState);
        }

        sStates = sbStates.ToString();

        string sDelimiter = "^";

        try
        {
            if (Is64Bit())
                iCount = SearchForDocumentsWithStatesAndDatesAndReturnColumnsX64(iProjectId, bSearchSubFolders, sFullText, bWholePhrase,
                    bAnyWords, bSearchAttributes, sDocumentName, sFileName, sDocumentDescription, bOriginalsOnly,
                    iEnvironmentId,
                    arAttributeNames, arAttributeValues, slAttributes.Count,
                    sStates,
                    sFileUpdatedAfter,
                    sFileUpdatedBefore,
                    sDocUpdatedAfter,
                    sDocUpdatedBefore,
                    arColumns, arColumns.Length,
                    sDelimiter,
                    out ppProjects, out ppDocumentIds, out arDocumentGuidStrings, out arDocumentNames, out arDocumentFileNames,
                    out arDocumentDescriptions, out arDocumentUpdateDates, out arDocumentVersions, out arAdditionalAttributeValues);
            else
                iCount = SearchForDocumentsWithStatesAndDatesAndReturnColumns(iProjectId, bSearchSubFolders, sFullText, bWholePhrase,
                    bAnyWords, bSearchAttributes, sDocumentName, sFileName, sDocumentDescription, bOriginalsOnly,
                    iEnvironmentId,
                    arAttributeNames, arAttributeValues, slAttributes.Count,
                    sStates,
                    sFileUpdatedAfter,
                    sFileUpdatedBefore,
                    sDocUpdatedAfter,
                    sDocUpdatedBefore,
                    arColumns, arColumns.Length,
                    sDelimiter,
                    out ppProjects, out ppDocumentIds, out arDocumentGuidStrings, out arDocumentNames, out arDocumentFileNames,
                    out arDocumentDescriptions, out arDocumentUpdateDates, out arDocumentVersions, out arAdditionalAttributeValues);
        }
        catch (Exception ex)
        {
            BPSUtilities.WriteLog("Error: {0}\n{1}", ex.Message, ex.StackTrace);
        }

        System.Diagnostics.Debug.WriteLine(string.Format("Returned {0} matches", iCount));

        DataTable dt = new DataTable("Documents");

        dt.Columns.Add("DocumentGUID", Type.GetType("System.String"));
        dt.Columns.Add("DocumentName", Type.GetType("System.String"));
        dt.Columns.Add("DocumentFileName", Type.GetType("System.String"));
        dt.Columns.Add("DocumentDescription", Type.GetType("System.String"));
        dt.Columns.Add("DocumentUpdateDate", Type.GetType("System.String"));
        dt.Columns.Add("DocumentVersion", Type.GetType("System.String"));
        dt.Columns.Add("ProjectId", Type.GetType("System.Int32"));
        dt.Columns.Add("DocumentId", Type.GetType("System.Int32"));
        dt.Columns.Add("ProjectPath", Type.GetType("System.String"));

        foreach (string sColumnName in listColumns)
        {
            if (!dt.Columns.Contains(sColumnName))
                dt.Columns.Add(sColumnName, Type.GetType("System.String"));
        }

        DataColumn[] pk = new DataColumn[1];
        pk[0] = dt.Columns["DocumentGUID"];
        dt.PrimaryKey = pk;

        SortedList<int, string> slProjectPaths = new SortedList<int, string>();

        if (iCount > 0)
        {
            int[] arProjects = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppProjects, iCount, out arProjects);
            int[] arDocumentIds = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppDocumentIds, iCount, out arDocumentIds);

            for (int i = 0; i < iCount; i++)
            {
                DataRow dr = dt.NewRow();

                dr["DocumentId"] = arDocumentIds[i];

                dr["ProjectID"] = arProjects[i];
                dr["DocumentGUID"] = arDocumentGuidStrings[i];
                dr["DocumentName"] = arDocumentNames[i];
                dr["DocumentFileName"] = arDocumentFileNames[i];
                dr["DocumentDescription"] = arDocumentDescriptions[i];
                dr["DocumentUpdateDate"] = arDocumentUpdateDates[i];
                dr["DocumentVersion"] = arDocumentVersions[i];

                if (bGetPath)
                {
                    string sProjectPath = string.Empty;
                    if (!slProjectPaths.TryGetValue(arProjects[i], out sProjectPath))
                    {
                        sProjectPath = PWWrapper.GetProjectNamePath(arProjects[i]);
                    }

                    dr["ProjectPath"] = sProjectPath;
                }

                if (arAdditionalAttributeValues.Length >= iCount)
                {
                    if (!string.IsNullOrEmpty(arAdditionalAttributeValues[i]))
                    {
                        string[] sAdditionalValues = arAdditionalAttributeValues[i].Split(sDelimiter.ToCharArray());

                        if (sAdditionalValues.Length == arColumns.Length)
                        {
                            for (int k = 0; k < arColumns.Length; k++)
                            {
                                if (dt.Columns.Contains(arColumns[k]))
                                {
                                    dr[arColumns[k]] = sAdditionalValues[k].Trim();
                                }
                            }
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine(string.Format("Additional values length was {0} while list of columns contained {1} values.",
                                sAdditionalValues.Length, arColumns.Length));
                        }
                    }
                }

                try
                {
                    dt.Rows.Add(dr);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(string.Format("Error: {0}\n{1}", ex.Message, ex.StackTrace));
                }
            }

            // because I always forget the column names
            StringBuilder sbColumns = new StringBuilder();

            foreach (DataColumn dc in dt.Columns)
            {
                sbColumns.Append(dc.ColumnName + ";");
            }

            System.Diagnostics.Debug.WriteLine(string.Format("Columns: {0}", sbColumns.ToString()));
        }

        GC.Collect();

        return dt;
    }

    public static DataTable SearchForDocumentsWithStatesAndUpdateDatesAndReturnColumns(int iProjectId, bool bSearchSubFolders,
        string sFullText, bool bWholePhrase, bool bAnyWords, bool bSearchAttributes,
        string sDocumentName, string sFileName, string sDocumentDescription, bool bOriginalsOnly,
        int iEnvironmentId,
        SortedList<string, string> slAttributes,
        List<int> listStates,
        string sDocUpdatedAfter, // early date 2009-10-22 01:00:00
        string sDocUpdatedBefore, // late date 2010-10-22 01:00:00
        bool bGetPath,
        List<string> listColumns)
    {
        IntPtr ppProjects = IntPtr.Zero;
        IntPtr ppDocumentIds = IntPtr.Zero;

        int iCount = 0;

        string[] arDocumentGuidStrings = null;
        string[] arDocumentNames = null;
        string[] arDocumentFileNames = null;
        string[] arDocumentDescriptions = null;
        string[] arDocumentUpdateDates = null;
        string[] arDocumentVersions = null;
        string[] arAdditionalAttributeValues = null;

        string[] arAttributeNames = new string[Math.Max(slAttributes.Count, 1)];
        string[] arAttributeValues = new string[Math.Max(slAttributes.Count, 1)];

        string[] arColumns = listColumns.ToArray();

        int iIndex = 0;

        foreach (KeyValuePair<string, string> kvp in slAttributes)
        {
            arAttributeNames[iIndex] = kvp.Key;
            arAttributeValues[iIndex] = kvp.Value;
            iIndex++;
        }

        System.Diagnostics.Debug.WriteLine("Starting query...");

        string sStates = string.Empty;

        StringBuilder sbStates = new StringBuilder();

        foreach (int iState in listStates)
        {
            if (sbStates.Length > 0)
                sbStates.Append(",");
            sbStates.Append(iState);
        }

        sStates = sbStates.ToString();

        string sDelimiter = "^";

        try
        {
            if (Is64Bit())
                iCount = SearchForDocumentsWithStatesAndUpdateDatesAndReturnColumnsX64(iProjectId, bSearchSubFolders, sFullText, bWholePhrase,
                    bAnyWords, bSearchAttributes, sDocumentName, sFileName, sDocumentDescription, bOriginalsOnly,
                    iEnvironmentId,
                    arAttributeNames, arAttributeValues, slAttributes.Count,
                    sStates,
                    sDocUpdatedAfter,
                    sDocUpdatedBefore,
                    arColumns, arColumns.Length,
                    sDelimiter,
                    out ppProjects, out ppDocumentIds, out arDocumentGuidStrings, out arDocumentNames, out arDocumentFileNames,
                    out arDocumentDescriptions, out arDocumentUpdateDates, out arDocumentVersions, out arAdditionalAttributeValues);
            else
                iCount = SearchForDocumentsWithStatesAndUpdateDatesAndReturnColumns(iProjectId, bSearchSubFolders, sFullText, bWholePhrase,
                    bAnyWords, bSearchAttributes, sDocumentName, sFileName, sDocumentDescription, bOriginalsOnly,
                    iEnvironmentId,
                    arAttributeNames, arAttributeValues, slAttributes.Count,
                    sStates,
                    sDocUpdatedAfter,
                    sDocUpdatedBefore,
                    arColumns, arColumns.Length,
                    sDelimiter,
                    out ppProjects, out ppDocumentIds, out arDocumentGuidStrings, out arDocumentNames, out arDocumentFileNames,
                    out arDocumentDescriptions, out arDocumentUpdateDates, out arDocumentVersions, out arAdditionalAttributeValues);
        }
        catch (Exception ex)
        {
            BPSUtilities.WriteLog("Error: {0}\n{1}", ex.Message, ex.StackTrace);
        }

        System.Diagnostics.Debug.WriteLine(string.Format("Returned {0} matches", iCount));

        DataTable dt = new DataTable("Documents");

        dt.Columns.Add("DocumentGUID", Type.GetType("System.String"));
        dt.Columns.Add("DocumentName", Type.GetType("System.String"));
        dt.Columns.Add("DocumentFileName", Type.GetType("System.String"));
        dt.Columns.Add("DocumentDescription", Type.GetType("System.String"));
        dt.Columns.Add("DocumentUpdateDate", Type.GetType("System.String"));
        dt.Columns.Add("DocumentVersion", Type.GetType("System.String"));
        dt.Columns.Add("ProjectId", Type.GetType("System.Int32"));
        dt.Columns.Add("DocumentId", Type.GetType("System.Int32"));
        dt.Columns.Add("ProjectPath", Type.GetType("System.String"));

        foreach (string sColumnName in listColumns)
        {
            if (!dt.Columns.Contains(sColumnName))
                dt.Columns.Add(sColumnName, Type.GetType("System.String"));
        }

        DataColumn[] pk = new DataColumn[1];
        pk[0] = dt.Columns["DocumentGUID"];
        dt.PrimaryKey = pk;

        SortedList<int, string> slProjectPaths = new SortedList<int, string>();

        if (iCount > 0)
        {
            int[] arProjects = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppProjects, iCount, out arProjects);
            int[] arDocumentIds = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppDocumentIds, iCount, out arDocumentIds);

            for (int i = 0; i < iCount; i++)
            {
                DataRow dr = dt.NewRow();

                dr["DocumentId"] = arDocumentIds[i];

                dr["ProjectID"] = arProjects[i];
                dr["DocumentGUID"] = arDocumentGuidStrings[i];
                dr["DocumentName"] = arDocumentNames[i];
                dr["DocumentFileName"] = arDocumentFileNames[i];
                dr["DocumentDescription"] = arDocumentDescriptions[i];
                dr["DocumentUpdateDate"] = arDocumentUpdateDates[i];
                dr["DocumentVersion"] = arDocumentVersions[i];

                if (bGetPath)
                {
                    string sProjectPath = string.Empty;
                    if (!slProjectPaths.TryGetValue(arProjects[i], out sProjectPath))
                    {
                        sProjectPath = PWWrapper.GetProjectNamePath(arProjects[i]);
                    }

                    dr["ProjectPath"] = sProjectPath;
                }

                if (arAdditionalAttributeValues.Length >= iCount)
                {
                    if (!string.IsNullOrEmpty(arAdditionalAttributeValues[i]))
                    {
                        string[] sAdditionalValues = arAdditionalAttributeValues[i].Split(sDelimiter.ToCharArray());

                        if (sAdditionalValues.Length == arColumns.Length)
                        {
                            for (int k = 0; k < arColumns.Length; k++)
                            {
                                if (dt.Columns.Contains(arColumns[k]))
                                {
                                    dr[arColumns[k]] = sAdditionalValues[k].Trim();
                                }
                            }
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine(string.Format("Additional values length was {0} while list of columns contained {1} values.",
                                sAdditionalValues.Length, arColumns.Length));
                        }
                    }
                }

                try
                {
                    dt.Rows.Add(dr);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(string.Format("Error: {0}\n{1}", ex.Message, ex.StackTrace));
                }
            }

            // because I always forget the column names
            StringBuilder sbColumns = new StringBuilder();

            foreach (DataColumn dc in dt.Columns)
            {
                sbColumns.Append(dc.ColumnName + ";");
            }

            System.Diagnostics.Debug.WriteLine(string.Format("Columns: {0}", sbColumns.ToString()));
        }

        GC.Collect();

        return dt;
    }

    public static DataTable SearchForDocuments(int iProjectId, bool bSearchSubFolders,
        string sFullText, bool bWholePhrase, bool bAnyWords, bool bSearchAttributes,
        string sDocumentName, string sFileName, string sDocumentDescription, bool bOriginalsOnly,
        int iEnvironmentId,
        SortedList<string, string> slAttributes, bool bGetPath)
    {
        IntPtr ppProjects = IntPtr.Zero;
        IntPtr ppDocumentIds = IntPtr.Zero;
        IntPtr ppVersionSeqNumbers = IntPtr.Zero;

        int iCount = 0;

        string[] arDocumentGuidStrings = null;
        string[] arDocumentNames = null;
        string[] arDocumentFileNames = null;
        string[] arDocumentDescriptions = null;
        string[] arDocumentUpdateDates = null;
        string[] arVersions = null;

        string[] arAttributeNames = new string[slAttributes.Count];
        string[] arAttributeValues = new string[slAttributes.Count];

        int iIndex = 0;

        foreach (KeyValuePair<string, string> kvp in slAttributes)
        {
            arAttributeNames[iIndex] = kvp.Key;
            arAttributeValues[iIndex] = kvp.Value;
            iIndex++;
        }

        System.Diagnostics.Debug.WriteLine("Starting query...");

        try
        {
            if (Is64Bit())
                iCount = SearchForDocumentsX64(iProjectId, bSearchSubFolders, sFullText, bWholePhrase,
                    bAnyWords, bSearchAttributes, sDocumentName, sFileName, sDocumentDescription, bOriginalsOnly, iEnvironmentId,
                    arAttributeNames, arAttributeValues, slAttributes.Count,
                    out ppProjects, out ppDocumentIds, out ppVersionSeqNumbers, out arDocumentGuidStrings, out arDocumentNames, out arDocumentFileNames,
                    out arDocumentDescriptions, out arDocumentUpdateDates, out arVersions);
            else
                iCount = SearchForDocuments(iProjectId, bSearchSubFolders, sFullText, bWholePhrase,
                    bAnyWords, bSearchAttributes, sDocumentName, sFileName, sDocumentDescription, bOriginalsOnly, iEnvironmentId,
                    arAttributeNames, arAttributeValues, slAttributes.Count,
                    out ppProjects, out ppDocumentIds, out ppVersionSeqNumbers, out arDocumentGuidStrings, out arDocumentNames, out arDocumentFileNames,
                    out arDocumentDescriptions, out arDocumentUpdateDates, out arVersions);

        }
        catch (Exception ex)
        {
            BPSUtilities.WriteLog("Error: {0}\n{1}", ex.Message, ex.StackTrace);

            if (ex.InnerException != null)
                BPSUtilities.WriteLog("Inner exception: {0}", ex.InnerException);
        }

        BPSUtilities.WriteLog("Returned {0} matches", iCount);

        DataTable dt = new DataTable("Documents");

        dt.Columns.Add("DocumentGUID", Type.GetType("System.String"));
        dt.Columns.Add("DocumentName", Type.GetType("System.String"));
        dt.Columns.Add("DocumentFileName", Type.GetType("System.String"));
        dt.Columns.Add("DocumentDescription", Type.GetType("System.String"));
        dt.Columns.Add("DocumentUpdateDate", Type.GetType("System.String"));
        dt.Columns.Add("DocumentVersion", Type.GetType("System.String"));
        dt.Columns.Add("DocumentVersionSequence", Type.GetType("System.Int32"));
        dt.Columns.Add("ProjectId", Type.GetType("System.Int32"));
        dt.Columns.Add("DocumentId", Type.GetType("System.Int32"));
        dt.Columns.Add("ProjectPath", Type.GetType("System.String"));

        DataColumn[] pk = new DataColumn[1];
        pk[0] = dt.Columns["DocumentGUID"];
        dt.PrimaryKey = pk;

        SortedList<int, string> slProjectPaths = new SortedList<int, string>();

        if (iCount > 0)
        {
            int[] arProjects = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppProjects, iCount, out arProjects);
            int[] arDocumentIds = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppDocumentIds, iCount, out arDocumentIds);
            int[] arDocumentVersionSequenceNumbers = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppVersionSeqNumbers, iCount, out arDocumentVersionSequenceNumbers);

            for (int i = 0; i < iCount; i++)
            {
                DataRow dr = dt.NewRow();

                dr["ProjectID"] = arProjects[i];
                dr["DocumentID"] = arDocumentIds[i];
                dr["DocumentGUID"] = arDocumentGuidStrings[i];
                dr["DocumentName"] = arDocumentNames[i];
                dr["DocumentFileName"] = arDocumentFileNames[i];
                dr["DocumentDescription"] = arDocumentDescriptions[i];
                dr["DocumentUpdateDate"] = arDocumentUpdateDates[i];
                dr["DocumentVersion"] = arVersions[i];
                dr["DocumentVersionSequence"] = arDocumentVersionSequenceNumbers[i];

                if (bGetPath)
                {
                    string sProjectPath = string.Empty;
                    if (!slProjectPaths.TryGetValue(arProjects[i], out sProjectPath))
                    {
                        sProjectPath = PWWrapper.GetProjectNamePath2(arProjects[i]);
                    }

                    dr["ProjectPath"] = sProjectPath;
                }

                try
                {
                    dt.Rows.Add(dr);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(string.Format("Error: {0}\n{1}", ex.Message, ex.StackTrace));
                }
            }

            // because I always forget the column names
            StringBuilder sbColumns = new StringBuilder();

            foreach (DataColumn dc in dt.Columns)
            {
                sbColumns.Append(dc.ColumnName + ";");
            }

            System.Diagnostics.Debug.WriteLine(string.Format("Columns: {0}", sbColumns.ToString()));
        }

        GC.Collect();

        return dt;
    }
    public static DataTable SearchForDocumentsByItemType(int iProjectId, bool bSearchSubFolders,
        string sDocumentName, string sFileName, string sDocumentDescription, bool bOriginalsOnly,
        int iEnvironmentId,
        SortedList<string, string> slAttributes, bool bGetPath, int iItemType = 12)
    {
        IntPtr ppProjects = IntPtr.Zero;
        IntPtr ppDocumentIds = IntPtr.Zero;
        IntPtr ppVersionSeqNumbers = IntPtr.Zero;

        int iCount = 0;

        string[] arDocumentGuidStrings = null;
        string[] arDocumentNames = null;
        string[] arDocumentFileNames = null;
        string[] arDocumentDescriptions = null;
        string[] arDocumentUpdateDates = null;
        string[] arVersions = null;

        string[] arAttributeNames = new string[slAttributes.Count];
        string[] arAttributeValues = new string[slAttributes.Count];

        int iIndex = 0;

        foreach (KeyValuePair<string, string> kvp in slAttributes)
        {
            arAttributeNames[iIndex] = kvp.Key;
            arAttributeValues[iIndex] = kvp.Value;
            iIndex++;
        }

        System.Diagnostics.Debug.WriteLine("Starting query...");

        try
        {
            if (Is64Bit())
                iCount = SearchForDocumentsByItemTypeX64(iProjectId, bSearchSubFolders, iItemType, sDocumentName, 
                    sFileName, sDocumentDescription, bOriginalsOnly, iEnvironmentId,
                    arAttributeNames, arAttributeValues, slAttributes.Count,
                    out ppProjects, out ppDocumentIds, out ppVersionSeqNumbers, out arDocumentGuidStrings, 
                    out arDocumentNames, out arDocumentFileNames,
                    out arDocumentDescriptions, out arDocumentUpdateDates, out arVersions);
            else
                iCount = SearchForDocumentsByItemType(iProjectId, bSearchSubFolders, iItemType, sDocumentName, 
                    sFileName, sDocumentDescription, bOriginalsOnly, iEnvironmentId,
                    arAttributeNames, arAttributeValues, slAttributes.Count,
                    out ppProjects, out ppDocumentIds, out ppVersionSeqNumbers, out arDocumentGuidStrings, 
                    out arDocumentNames, out arDocumentFileNames,
                    out arDocumentDescriptions, out arDocumentUpdateDates, out arVersions);
        }
        catch (Exception ex)
        {
            BPSUtilities.WriteLog("Error: {0}\n{1}", ex.Message, ex.StackTrace);

            if (ex.InnerException != null)
                BPSUtilities.WriteLog("Inner exception: {0}", ex.InnerException);
        }

        BPSUtilities.WriteLog("Returned {0} matches", iCount);

        DataTable dt = new DataTable("Documents");

        dt.Columns.Add("DocumentGUID", Type.GetType("System.String"));
        dt.Columns.Add("DocumentName", Type.GetType("System.String"));
        dt.Columns.Add("DocumentFileName", Type.GetType("System.String"));
        dt.Columns.Add("DocumentDescription", Type.GetType("System.String"));
        dt.Columns.Add("DocumentUpdateDate", Type.GetType("System.String"));
        dt.Columns.Add("DocumentVersion", Type.GetType("System.String"));
        dt.Columns.Add("DocumentVersionSequence", Type.GetType("System.Int32"));
        dt.Columns.Add("ProjectId", Type.GetType("System.Int32"));
        dt.Columns.Add("DocumentId", Type.GetType("System.Int32"));
        dt.Columns.Add("ProjectPath", Type.GetType("System.String"));

        DataColumn[] pk = new DataColumn[1];
        pk[0] = dt.Columns["DocumentGUID"];
        dt.PrimaryKey = pk;

        SortedList<int, string> slProjectPaths = new SortedList<int, string>();

        if (iCount > 0)
        {
            int[] arProjects = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppProjects, iCount, out arProjects);
            int[] arDocumentIds = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppDocumentIds, iCount, out arDocumentIds);
            int[] arDocumentVersionSequenceNumbers = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppVersionSeqNumbers, iCount, out arDocumentVersionSequenceNumbers);

            for (int i = 0; i < iCount; i++)
            {
                DataRow dr = dt.NewRow();

                dr["ProjectID"] = arProjects[i];
                dr["DocumentID"] = arDocumentIds[i];
                dr["DocumentGUID"] = arDocumentGuidStrings[i];
                dr["DocumentName"] = arDocumentNames[i];
                dr["DocumentFileName"] = arDocumentFileNames[i];
                dr["DocumentDescription"] = arDocumentDescriptions[i];
                dr["DocumentUpdateDate"] = arDocumentUpdateDates[i];
                dr["DocumentVersion"] = arVersions[i];
                dr["DocumentVersionSequence"] = arDocumentVersionSequenceNumbers[i];

                if (bGetPath)
                {
                    string sProjectPath = string.Empty;
                    if (!slProjectPaths.TryGetValue(arProjects[i], out sProjectPath))
                    {
                        sProjectPath = PWWrapper.GetProjectNamePath2(arProjects[i]);
                    }

                    dr["ProjectPath"] = sProjectPath;
                }

                try
                {
                    dt.Rows.Add(dr);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(string.Format("Error: {0}\n{1}", ex.Message, ex.StackTrace));
                }
            }

            // because I always forget the column names
            StringBuilder sbColumns = new StringBuilder();

            foreach (DataColumn dc in dt.Columns)
            {
                sbColumns.Append(dc.ColumnName + ";");
            }

            System.Diagnostics.Debug.WriteLine(string.Format("Columns: {0}", sbColumns.ToString()));
        }

        GC.Collect();

        return dt;
    }


    // SearchForDocumentsReturnFileNamesOriginalNos

    [DllImport("PWSearchWrapper.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall)]
    private extern static int SearchForDocumentsReturnFileNamesOriginalNos(
        int iProjectId,
        bool bIncludeSubFolders,
        string sFullTextSearchString,
        bool bWholePhrase,
        bool bAnyWord,
        bool bSearchAttributes, // otherwise full text
        string sDocumentName,
        string sFileName,
        string sDocumentDescription,
        bool bOriginalsOnly,
        int iEnvironmentId,
        [In][MarshalAsAttribute(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] arAttributeNames,
        [In][MarshalAsAttribute(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] arAttributeValues,
        int size,
        [Out] out IntPtr ppProjects,
        [Out] out IntPtr ppDocumentIds,
        [Out] out IntPtr ppVersionSeqNumbers,
        [Out] out IntPtr ppOriginalNumbers,
        [Out] out IntPtr ppStorageIds,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentFileNames,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arMimeTypes
        );

    [DllImport("PWSearchWrapperX64.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall, EntryPoint = "SearchForDocumentsReturnFileNamesOriginalNos")]
    private extern static int SearchForDocumentsReturnFileNamesOriginalNosX64(
        int iProjectId,
        bool bIncludeSubFolders,
        string sFullTextSearchString,
        bool bWholePhrase,
        bool bAnyWord,
        bool bSearchAttributes, // otherwise full text
        string sDocumentName,
        string sFileName,
        string sDocumentDescription,
        bool bOriginalsOnly,
        int iEnvironmentId,
        [In][MarshalAsAttribute(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] arAttributeNames,
        [In][MarshalAsAttribute(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] arAttributeValues,
        int size,
        [Out] out IntPtr ppProjects,
        [Out] out IntPtr ppDocumentIds,
        [Out] out IntPtr ppVersionSeqNumbers,
        [Out] out IntPtr ppOriginalNumbers,
        [Out] out IntPtr ppStorageIds,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentFileNames,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arMimeTypes
        );

    // added storageIds and MimeTypes to return 2019-04-16
    public static DataTable SearchForDocumentsReturnFileNamesAndOriginalNos(int iProjectId, bool bSearchSubFolders,
        string sFullText, bool bWholePhrase, bool bAnyWords, bool bSearchAttributes,
        string sDocumentName, string sFileName, string sDocumentDescription, bool bOriginalsOnly,
        int iEnvironmentId,
        SortedList<string, string> slAttributes, bool bGetPath)
    {
        IntPtr ppProjects = IntPtr.Zero;
        IntPtr ppDocumentIds = IntPtr.Zero;
        IntPtr ppVersionSeqNumbers = IntPtr.Zero;
        IntPtr ppOriginalNumbers = IntPtr.Zero;
        IntPtr ppStorageIds = IntPtr.Zero;

        int iCount = 0;

        string[] arDocumentFileNames = null;
        string[] arMimeTypes = null;

        string[] arAttributeNames = new string[slAttributes.Count];
        string[] arAttributeValues = new string[slAttributes.Count];

        int iIndex = 0;

        foreach (KeyValuePair<string, string> kvp in slAttributes)
        {
            arAttributeNames[iIndex] = kvp.Key;
            arAttributeValues[iIndex] = kvp.Value;
            iIndex++;
        }

        System.Diagnostics.Debug.WriteLine("Starting query...");

        try
        {
            if (Is64Bit())
                iCount = SearchForDocumentsReturnFileNamesOriginalNosX64(iProjectId, bSearchSubFolders, sFullText, bWholePhrase,
                    bAnyWords, bSearchAttributes, sDocumentName, sFileName, sDocumentDescription, bOriginalsOnly, iEnvironmentId,
                    arAttributeNames, arAttributeValues, slAttributes.Count,
                    out ppProjects, out ppDocumentIds, out ppVersionSeqNumbers, out ppOriginalNumbers, out ppStorageIds, 
                    out arDocumentFileNames, out arMimeTypes);
            else
                iCount = SearchForDocumentsReturnFileNamesOriginalNos(iProjectId, bSearchSubFolders, sFullText, bWholePhrase,
                    bAnyWords, bSearchAttributes, sDocumentName, sFileName, sDocumentDescription, bOriginalsOnly, iEnvironmentId,
                    arAttributeNames, arAttributeValues, slAttributes.Count,
                    out ppProjects, out ppDocumentIds, out ppVersionSeqNumbers, out ppOriginalNumbers, out ppStorageIds, 
                    out arDocumentFileNames, out arMimeTypes);

        }
        catch (Exception ex)
        {
            BPSUtilities.WriteLog("Error: {0}\n{1}", ex.Message, ex.StackTrace);

            if (ex.InnerException != null)
                BPSUtilities.WriteLog("Inner exception: {0}", ex.InnerException);
        }

        DataTable dt = new DataTable("Documents");

        dt.Columns.Add("DocumentGUID", Type.GetType("System.String"));
        dt.Columns.Add("DocumentName", Type.GetType("System.String"));
        dt.Columns.Add("DocumentFileName", Type.GetType("System.String"));
        dt.Columns.Add("DocumentDescription", Type.GetType("System.String"));
        dt.Columns.Add("DocumentUpdateDate", Type.GetType("System.String"));
        dt.Columns.Add("DocumentVersion", Type.GetType("System.String"));
        dt.Columns.Add("DocumentVersionSequence", Type.GetType("System.Int32"));
        dt.Columns.Add("ProjectId", Type.GetType("System.Int32"));
        dt.Columns.Add("DocumentId", Type.GetType("System.Int32"));
        dt.Columns.Add("ProjectPath", Type.GetType("System.String"));
        dt.Columns.Add("DocumentOriginalNumber", Type.GetType("System.Int32"));
        dt.Columns.Add("DocumentStorageId", Type.GetType("System.Int32"));
        dt.Columns.Add("DocumentMimeType", Type.GetType("System.String"));

        DataColumn[] pk = new DataColumn[2];
        pk[0] = dt.Columns["ProjectId"];
        pk[1] = dt.Columns["DocumentId"];
        dt.PrimaryKey = pk;

        SortedList<int, string> slProjectPaths = new SortedList<int, string>();

        if (iCount > 0)
        {
            int[] arProjects = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppProjects, iCount, out arProjects);
            int[] arDocumentIds = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppDocumentIds, iCount, out arDocumentIds);
            int[] arDocumentVersionSequenceNumbers = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppVersionSeqNumbers, iCount, out arDocumentVersionSequenceNumbers);
            int[] arDocumentOriginalNumbers = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppOriginalNumbers, iCount, out arDocumentOriginalNumbers);
            int[] arStorageIds = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppStorageIds, iCount, out arStorageIds);

            for (int i = 0; i < iCount; i++)
            {
                DataRow dr = dt.NewRow();

                dr["ProjectID"] = arProjects[i];
                dr["DocumentID"] = arDocumentIds[i];
                dr["DocumentFileName"] = arDocumentFileNames[i];
                dr["DocumentVersionSequence"] = arDocumentVersionSequenceNumbers[i];
                dr["DocumentOriginalNumber"] = arDocumentOriginalNumbers[i];
                dr["DocumentStorageId"] = arStorageIds[i];
                dr["DocumentMimeType"] = arMimeTypes[i];

                if (bGetPath)
                {
                    string sProjectPath = string.Empty;
                    if (!slProjectPaths.TryGetValue(arProjects[i], out sProjectPath))
                    {
                        sProjectPath = PWWrapper.GetProjectNamePath2(arProjects[i]);
                    }

                    dr["ProjectPath"] = sProjectPath;
                }

                try
                {
                    dt.Rows.Add(dr);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(string.Format("Error: {0}\n{1}", ex.Message, ex.StackTrace));
                }
            }

            // because I always forget the column names
            StringBuilder sbColumns = new StringBuilder();

            foreach (DataColumn dc in dt.Columns)
            {
                sbColumns.Append(dc.ColumnName + ";");
            }

            System.Diagnostics.Debug.WriteLine(string.Format("Columns: {0}", sbColumns.ToString()));
        }

        GC.Collect();

        return dt;
    }

    public static DataTable SearchForDocumentsMinMemory(int iProjectId, bool bSearchSubFolders,
        string sFullText, bool bWholePhrase, bool bAnyWords, bool bSearchAttributes,
        string sDocumentName, string sFileName, string sDocumentDescription, bool bOriginalsOnly,
        int iEnvironmentId,
        SortedList<string, string> slAttributes, bool bGetPath)
    {
        IntPtr ppProjects = IntPtr.Zero;
        IntPtr ppDocumentIds = IntPtr.Zero;

        int iCount = 0;

        string[] arAttributeNames = new string[slAttributes.Count];
        string[] arAttributeValues = new string[slAttributes.Count];

        int iIndex = 0;

        foreach (KeyValuePair<string, string> kvp in slAttributes)
        {
            arAttributeNames[iIndex] = kvp.Key;
            arAttributeValues[iIndex] = kvp.Value;
            iIndex++;
        }

        System.Diagnostics.Debug.WriteLine("Starting query...");

        try
        {
            if (Is64Bit())
                iCount = SearchForDocumentsMinMemoryX64(iProjectId, bSearchSubFolders, sFullText, bWholePhrase,
                    bAnyWords, bSearchAttributes, sDocumentName, sFileName, sDocumentDescription, bOriginalsOnly, iEnvironmentId,
                    arAttributeNames, arAttributeValues, slAttributes.Count,
                    out ppProjects, out ppDocumentIds);
            else
                iCount = SearchForDocumentsMinMemory(iProjectId, bSearchSubFolders, sFullText, bWholePhrase,
                    bAnyWords, bSearchAttributes, sDocumentName, sFileName, sDocumentDescription, bOriginalsOnly, iEnvironmentId,
                    arAttributeNames, arAttributeValues, slAttributes.Count,
                    out ppProjects, out ppDocumentIds);
        }
        catch (Exception ex)
        {
            BPSUtilities.WriteLog("Error: {0}\n{1}", ex.Message, ex.StackTrace);

            if (ex.InnerException != null)
                BPSUtilities.WriteLog("Inner exception: {0}", ex.InnerException);
        }

        // BPSUtilities.WriteLog("Returned {0} matches", iCount);

        DataTable dt = new DataTable("Documents");

        dt.Columns.Add("DocumentGUID", Type.GetType("System.String"));
        dt.Columns.Add("DocumentName", Type.GetType("System.String"));
        dt.Columns.Add("DocumentFileName", Type.GetType("System.String"));
        dt.Columns.Add("DocumentDescription", Type.GetType("System.String"));
        dt.Columns.Add("DocumentUpdateDate", Type.GetType("System.String"));
        dt.Columns.Add("DocumentVersion", Type.GetType("System.String"));
        dt.Columns.Add("DocumentVersionSequence", Type.GetType("System.Int32"));
        dt.Columns.Add("ProjectId", Type.GetType("System.Int32"));
        dt.Columns.Add("DocumentId", Type.GetType("System.Int32"));
        dt.Columns.Add("ProjectPath", Type.GetType("System.String"));
        dt.Columns.Add("DocumentWorkflowId", Type.GetType("System.Int32"));
        dt.Columns.Add("DocumentStateId", Type.GetType("System.Int32"));
        dt.Columns.Add("DocumentFileSize", typeof(UInt64));

        //DataColumn[] pk = new DataColumn[1];
        //pk[0] = dt.Columns["DocumentGUID"];
        //dt.PrimaryKey = pk;

        SortedList<int, string> slProjectPaths = new SortedList<int, string>();

        if (iCount > 0)
        {
            int[] arProjects = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppProjects, iCount, out arProjects);
            int[] arDocumentIds = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppDocumentIds, iCount, out arDocumentIds);

            for (int i = 0; i < iCount; i++)
            {
                DataRow dr = dt.NewRow();

                dr["ProjectID"] = arProjects[i];
                dr["DocumentID"] = arDocumentIds[i];


                //dr["DocumentGUID"] = arDocumentGuidStrings[i];
                //dr["DocumentName"] = arDocumentNames[i];
                //dr["DocumentFileName"] = arDocumentFileNames[i];
                //dr["DocumentDescription"] = arDocumentDescriptions[i];
                //dr["DocumentUpdateDate"] = arDocumentUpdateDates[i];
                //dr["DocumentVersion"] = arVersions[i];
                //dr["DocumentVersionSequence"] = arDocumentVersionSequenceNumbers[i];

                if (bGetPath)
                {
                    string sProjectPath = string.Empty;
                    if (!slProjectPaths.TryGetValue(arProjects[i], out sProjectPath))
                    {
                        sProjectPath = PWWrapper.GetProjectNamePath2(arProjects[i]);
                    }

                    dr["ProjectPath"] = sProjectPath;
                }

                try
                {
                    dt.Rows.Add(dr);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(string.Format("Error: {0}\n{1}", ex.Message, ex.StackTrace));
                }
            }

            // because I always forget the column names
            StringBuilder sbColumns = new StringBuilder();

            foreach (DataColumn dc in dt.Columns)
            {
                sbColumns.Append(dc.ColumnName + ";");
            }

            System.Diagnostics.Debug.WriteLine(string.Format("Columns: {0}", sbColumns.ToString()));
        }

        GC.Collect();

        return dt;
    }

    public static DataTable SearchForDocumentsByQueryName(string sQueryName, bool bGetPath)
    {
        int iQueryId = Is64Bit() ? GetSavedSearchIdX64(sQueryName) : GetSavedSearchId(sQueryName);

        if (iQueryId > 0)
        {
            return SearchForDocumentsByQueryId(iQueryId, bGetPath);
        }

        System.Diagnostics.Debug.WriteLine(string.Format("Saved Search '{0}' not found", sQueryName));

        return new DataTable();
    }

    public static DataTable SearchForDocumentsByQueryId(int iQueryId, bool bGetPath)
    {
        IntPtr ppProjects = IntPtr.Zero;
        IntPtr ppDocumentIds = IntPtr.Zero;

        int iCount = 0;

        string[] arDocumentGuidStrings = null;
        string[] arDocumentNames = null;
        string[] arDocumentFileNames = null;
        string[] arDocumentDescriptions = null;
        string[] arDocumentUpdateDates = null;
        string[] arVersions = null;

        System.Diagnostics.Debug.WriteLine("Starting query...");

        try
        {
            if (Is64Bit())
                iCount = SearchForDocumentsByQueryIdX64(iQueryId,
                    out ppProjects, out ppDocumentIds, out arDocumentGuidStrings, out arDocumentNames, out arDocumentFileNames,
                    out arDocumentDescriptions, out arDocumentUpdateDates, out arVersions);
            else
                iCount = SearchForDocumentsByQueryId(iQueryId,
                    out ppProjects, out ppDocumentIds, out arDocumentGuidStrings, out arDocumentNames, out arDocumentFileNames,
                    out arDocumentDescriptions, out arDocumentUpdateDates, out arVersions);
        }
        catch (Exception ex)
        {
            BPSUtilities.WriteLog("Error: {0}\n{1}", ex.Message, ex.StackTrace);
        }

        System.Diagnostics.Debug.WriteLine(string.Format("Returned {0} matches", iCount));

        DataTable dt = new DataTable("Documents");

        dt.Columns.Add("DocumentGUID", Type.GetType("System.String"));
        dt.Columns.Add("DocumentName", Type.GetType("System.String"));
        dt.Columns.Add("DocumentFileName", Type.GetType("System.String"));
        dt.Columns.Add("DocumentDescription", Type.GetType("System.String"));
        dt.Columns.Add("DocumentUpdateDate", Type.GetType("System.String"));
        dt.Columns.Add("DocumentVersion", Type.GetType("System.String"));
        dt.Columns.Add("ProjectId", Type.GetType("System.Int32"));
        dt.Columns.Add("DocumentId", Type.GetType("System.Int32"));
        dt.Columns.Add("ProjectPath", Type.GetType("System.String"));

        DataColumn[] pk = new DataColumn[1];
        pk[0] = dt.Columns["DocumentGUID"];
        dt.PrimaryKey = pk;

        SortedList<int, string> slProjectPaths = new SortedList<int, string>();

        if (iCount > 0)
        {
            int[] arProjects = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppProjects, iCount, out arProjects);
            int[] arDocumentIds = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppDocumentIds, iCount, out arDocumentIds);

            for (int i = 0; i < iCount; i++)
            {
                DataRow dr = dt.NewRow();

                dr["ProjectID"] = arProjects[i];
                dr["DocumentID"] = arDocumentIds[i];
                dr["DocumentGUID"] = arDocumentGuidStrings[i];
                dr["DocumentName"] = arDocumentNames[i];
                dr["DocumentFileName"] = arDocumentFileNames[i];
                dr["DocumentDescription"] = arDocumentDescriptions[i];
                dr["DocumentUpdateDate"] = arDocumentUpdateDates[i];
                dr["DocumentVersion"] = arVersions[i];

                if (bGetPath)
                {
                    string sProjectPath = string.Empty;
                    if (!slProjectPaths.TryGetValue(arProjects[i], out sProjectPath))
                    {
                        sProjectPath = PWWrapper.GetProjectNamePath2(arProjects[i]);
                    }

                    dr["ProjectPath"] = sProjectPath;
                }

                try
                {
                    dt.Rows.Add(dr);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(string.Format("Error: {0}\n{1}", ex.Message, ex.StackTrace));
                }
            }

            // because I always forget the column names
            StringBuilder sbColumns = new StringBuilder();

            foreach (DataColumn dc in dt.Columns)
            {
                sbColumns.Append(dc.ColumnName + ";");
            }

            System.Diagnostics.Debug.WriteLine(string.Format("Columns: {0}", sbColumns.ToString()));
        }

        GC.Collect();

        return dt;
    }

    public static DataTable SearchForDocumentsByQueryNameMinMemory(string sQueryName, bool bGetPath)
    {
        int iQueryId = Is64Bit() ? GetSavedSearchIdX64(sQueryName) : GetSavedSearchId(sQueryName);

        if (iQueryId > 0)
        {
            return SearchForDocumentsByQueryIdMinMemory(iQueryId, bGetPath);
        }

        System.Diagnostics.Debug.WriteLine(string.Format("Saved Search '{0}' not found", sQueryName));

        return new DataTable();
    }

    public static DataTable SearchForDocumentsByQueryIdMinMemory(int iQueryId, bool bGetPath)
    {
        IntPtr ppProjects = IntPtr.Zero;
        IntPtr ppDocumentIds = IntPtr.Zero;

        int iCount = 0;

        System.Diagnostics.Debug.WriteLine("Starting query...");

        try
        {
            if (Is64Bit())
                iCount = SearchForDocumentsByQueryIdMinMemoryX64(iQueryId,
                    out ppProjects, out ppDocumentIds);
            else
                iCount = SearchForDocumentsByQueryIdMinMemory(iQueryId,
                    out ppProjects, out ppDocumentIds);
        }
        catch (Exception ex)
        {
            BPSUtilities.WriteLog("Error: {0}\n{1}", ex.Message, ex.StackTrace);
        }

        System.Diagnostics.Debug.WriteLine(string.Format("Returned {0} matches", iCount));

        DataTable dt = new DataTable("Documents");

        dt.Columns.Add("DocumentGUID", Type.GetType("System.String"));
        dt.Columns.Add("DocumentName", Type.GetType("System.String"));
        dt.Columns.Add("DocumentFileName", Type.GetType("System.String"));
        dt.Columns.Add("DocumentDescription", Type.GetType("System.String"));
        dt.Columns.Add("DocumentUpdateDate", Type.GetType("System.String"));
        dt.Columns.Add("DocumentVersion", Type.GetType("System.String"));
        dt.Columns.Add("ProjectId", Type.GetType("System.Int32"));
        dt.Columns.Add("DocumentId", Type.GetType("System.Int32"));
        dt.Columns.Add("ProjectPath", Type.GetType("System.String"));

        //DataColumn[] pk = new DataColumn[1];
        //pk[0] = dt.Columns["DocumentGUID"];
        //dt.PrimaryKey = pk;

        SortedList<int, string> slProjectPaths = new SortedList<int, string>();

        if (iCount > 0)
        {
            int[] arProjects = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppProjects, iCount, out arProjects);
            int[] arDocumentIds = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppDocumentIds, iCount, out arDocumentIds);

            for (int i = 0; i < iCount; i++)
            {
                DataRow dr = dt.NewRow();

                dr["ProjectID"] = arProjects[i];
                dr["DocumentID"] = arDocumentIds[i];
                // dr["DocumentGUID"] = arDocumentGuidStrings[i];
                // dr["DocumentName"] = arDocumentNames[i];
                // dr["DocumentFileName"] = arDocumentFileNames[i];
                // dr["DocumentDescription"] = arDocumentDescriptions[i];
                // dr["DocumentUpdateDate"] = arDocumentUpdateDates[i];
                // dr["DocumentVersion"] = arVersions[i];

                if (bGetPath)
                {
                    string sProjectPath = string.Empty;
                    if (!slProjectPaths.TryGetValue(arProjects[i], out sProjectPath))
                    {
                        sProjectPath = PWWrapper.GetProjectNamePath2(arProjects[i]);
                    }

                    dr["ProjectPath"] = sProjectPath;
                }

                try
                {
                    dt.Rows.Add(dr);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(string.Format("Error: {0}\n{1}", ex.Message, ex.StackTrace));
                }
            }

            // because I always forget the column names
            StringBuilder sbColumns = new StringBuilder();

            foreach (DataColumn dc in dt.Columns)
            {
                sbColumns.Append(dc.ColumnName + ";");
            }

            System.Diagnostics.Debug.WriteLine(string.Format("Columns: {0}", sbColumns.ToString()));
        }

        GC.Collect();

        return dt;
    }

    public static DataTable SearchForDocumentsByQueryNameReturningWorkflowsStatesFileSizes(string sQueryName, bool bGetPath)
    {
        int iQueryId = Is64Bit() ? GetSavedSearchIdX64(sQueryName) : GetSavedSearchId(sQueryName);

        if (iQueryId > 0)
        {
            return SearchForDocumentsByQueryIdReturningWorkflowsStatesFileSizes(iQueryId, bGetPath);
        }

        System.Diagnostics.Debug.WriteLine(string.Format("Saved Search '{0}' not found", sQueryName));

        return new DataTable();
    }

    public static DataTable SearchForDocumentsReturningWorkflowsStatesFileSizes(int iProjectId, bool bSearchSubFolders,
        string sFullText, bool bWholePhrase, bool bAnyWords, bool bSearchAttributes,
        string sDocumentName, string sFileName, string sDocumentDescription, bool bOriginalsOnly,
        int iEnvironmentId,
        SortedList<string, string> slAttributes, bool bGetPath)
    {
        IntPtr ppProjects = IntPtr.Zero;
        IntPtr ppDocumentIds = IntPtr.Zero;
        IntPtr ppVersionSeqNumbers = IntPtr.Zero;

        IntPtr ppWorkflowIds = IntPtr.Zero;
        IntPtr ppStateIds = IntPtr.Zero;
        IntPtr ppStorageIds = IntPtr.Zero;

        int iCount = 0;

        string[] arDocumentGuidStrings = null;
        string[] arDocumentNames = null;
        string[] arDocumentFileNames = null;
        string[] arDocumentDescriptions = null;
        string[] arDocumentUpdateDates = null;
        string[] arVersions = null;
        string[] arFileSizes = null;
        string[] arMimeTypes = null;

        string[] arAttributeNames = new string[slAttributes.Count];
        string[] arAttributeValues = new string[slAttributes.Count];

        int iIndex = 0;

        foreach (KeyValuePair<string, string> kvp in slAttributes)
        {
            arAttributeNames[iIndex] = kvp.Key;
            arAttributeValues[iIndex] = kvp.Value;
            iIndex++;
        }

        System.Diagnostics.Debug.WriteLine("Starting query...");

        try
        {
            if (Is64Bit())
                iCount = SearchForDocumentsReturningWorkflowStateSizesAsStringsX64(iProjectId, bSearchSubFolders, sFullText, bWholePhrase,
                    bAnyWords, bSearchAttributes, sDocumentName, sFileName, sDocumentDescription, bOriginalsOnly, iEnvironmentId,
                    arAttributeNames, arAttributeValues, slAttributes.Count,
                    out ppProjects, out ppDocumentIds, out ppVersionSeqNumbers, out ppWorkflowIds, out ppStateIds, out ppStorageIds, out arDocumentGuidStrings, out arDocumentNames, out arDocumentFileNames,
                    out arDocumentDescriptions, out arDocumentUpdateDates, out arVersions, out arFileSizes, out arMimeTypes);
            else
                iCount = SearchForDocumentsReturningWorkflowStateSizesAsStrings(iProjectId, bSearchSubFolders, sFullText, bWholePhrase,
                    bAnyWords, bSearchAttributes, sDocumentName, sFileName, sDocumentDescription, bOriginalsOnly, iEnvironmentId,
                    arAttributeNames, arAttributeValues, slAttributes.Count,
                    out ppProjects, out ppDocumentIds, out ppVersionSeqNumbers, out ppWorkflowIds, out ppStateIds, out ppStorageIds, out arDocumentGuidStrings, out arDocumentNames, out arDocumentFileNames,
                    out arDocumentDescriptions, out arDocumentUpdateDates, out arVersions, out arFileSizes, out arMimeTypes);
        }
        catch (Exception ex)
        {
            BPSUtilities.WriteLog("Error: {0}\n{1}", ex.Message, ex.StackTrace);
        }

        System.Diagnostics.Debug.WriteLine(string.Format("Returned {0} matches", iCount));

        DataTable dt = new DataTable("Documents");

        dt.Columns.Add("DocumentGUID", Type.GetType("System.String"));
        dt.Columns.Add("DocumentName", Type.GetType("System.String"));
        dt.Columns.Add("DocumentFileName", Type.GetType("System.String"));
        dt.Columns.Add("DocumentDescription", Type.GetType("System.String"));
        dt.Columns.Add("DocumentUpdateDate", Type.GetType("System.String"));
        dt.Columns.Add("DocumentVersion", Type.GetType("System.String"));
        dt.Columns.Add("DocumentVersionSequence", Type.GetType("System.Int32"));
        dt.Columns.Add("ProjectId", Type.GetType("System.Int32"));
        dt.Columns.Add("DocumentId", Type.GetType("System.Int32"));
        dt.Columns.Add("ProjectPath", Type.GetType("System.String"));
        dt.Columns.Add("DocumentWorkflowId", Type.GetType("System.Int32"));
        dt.Columns.Add("DocumentStateId", Type.GetType("System.Int32"));
        dt.Columns.Add("DocumentFileSize", typeof(UInt64));
        dt.Columns.Add("DocumentStorageId", Type.GetType("System.Int32"));
        dt.Columns.Add("DocumentMimeType", Type.GetType("System.String"));

        DataColumn[] pk = new DataColumn[1];
        pk[0] = dt.Columns["DocumentGUID"];
        dt.PrimaryKey = pk;

        SortedList<int, string> slProjectPaths = new SortedList<int, string>();

        if (iCount > 0)
        {
            int[] arProjects = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppProjects, iCount, out arProjects);
            int[] arDocumentIds = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppDocumentIds, iCount, out arDocumentIds);
            int[] arDocumentVersionSequenceNumbers = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppVersionSeqNumbers, iCount, out arDocumentVersionSequenceNumbers);

            int[] arWorkflowIds = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppWorkflowIds, iCount, out arWorkflowIds);
            int[] arStateIds = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppStateIds, iCount, out arStateIds);

            int[] arStorageIds = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppStorageIds, iCount, out arStorageIds);

            for (int i = 0; i < iCount; i++)
            {
                DataRow dr = dt.NewRow();

                dr["ProjectID"] = arProjects[i];
                dr["DocumentID"] = arDocumentIds[i];
                dr["DocumentGUID"] = arDocumentGuidStrings[i];
                dr["DocumentName"] = arDocumentNames[i];
                dr["DocumentFileName"] = arDocumentFileNames[i];
                dr["DocumentDescription"] = arDocumentDescriptions[i];
                dr["DocumentUpdateDate"] = arDocumentUpdateDates[i];
                dr["DocumentVersion"] = arVersions[i];
                dr["DocumentVersionSequence"] = arDocumentVersionSequenceNumbers[i];

                dr["DocumentWorkflowId"] = arWorkflowIds[i];
                dr["DocumentStateId"] = arStateIds[i];

                dr["DocumentStorageId"] = arStorageIds[i];
                dr["DocumentMimeType"] = arMimeTypes[i];

                UInt64 uiFileSize = 0;

                UInt64.TryParse(arFileSizes[i], out uiFileSize);

                dr["DocumentFileSize"] = uiFileSize;

                if (bGetPath)
                {
                    string sProjectPath = string.Empty;
                    if (!slProjectPaths.TryGetValue(arProjects[i], out sProjectPath))
                    {
                        sProjectPath = PWWrapper.GetProjectNamePath2(arProjects[i]);
                    }

                    dr["ProjectPath"] = sProjectPath;
                }

                try
                {
                    dt.Rows.Add(dr);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(string.Format("Error: {0}\n{1}", ex.Message, ex.StackTrace));
                }
            }

            // because I always forget the column names
            StringBuilder sbColumns = new StringBuilder();

            foreach (DataColumn dc in dt.Columns)
            {
                sbColumns.Append(dc.ColumnName + ";");
            }

            System.Diagnostics.Debug.WriteLine(string.Format("Columns: {0}", sbColumns.ToString()));
        }

        GC.Collect();

        return dt;
    }

    public static DataTable SearchForDocumentsMultiValuesReturningWorkflowsStatesFileSizes(int iProjectId, bool bSearchSubFolders,
        string sDocumentName, string sFileName, string sDocumentDescription, bool bOriginalsOnly,
        int iEnvironmentId, string sAttributeName,
        List<string> listValues, bool bGetPath)
    {
        IntPtr ppProjects = IntPtr.Zero;
        IntPtr ppDocumentIds = IntPtr.Zero;
        IntPtr ppVersionSeqNumbers = IntPtr.Zero;

        IntPtr ppWorkflowIds = IntPtr.Zero;
        IntPtr ppStateIds = IntPtr.Zero;
        IntPtr ppStorageIds = IntPtr.Zero;

        int iCount = 0;

        string[] arDocumentGuidStrings = null;
        string[] arDocumentNames = null;
        string[] arDocumentFileNames = null;
        string[] arDocumentDescriptions = null;
        string[] arDocumentUpdateDates = null;
        string[] arVersions = null;
        string[] arFileSizes = null;
        string[] arMimeTypes = null;

        string[] arAttributeValues = new string[listValues.Count];

        int iIndex = 0;

        foreach (string sValue in listValues)
        {
            arAttributeValues[iIndex] = sValue;
            iIndex++;
        }

        System.Diagnostics.Debug.WriteLine("Starting query...");

        try
        {
            if (Is64Bit())
                iCount = SearchForDocumentsMultiValuesReturningWorkflowStateSizesAsStringsX64(iProjectId, bSearchSubFolders, sDocumentName, sFileName, sDocumentDescription, bOriginalsOnly, iEnvironmentId,
                    sAttributeName, arAttributeValues, listValues.Count,
                    out ppProjects, out ppDocumentIds, out ppVersionSeqNumbers, out ppWorkflowIds, out ppStateIds, out ppStorageIds, out arDocumentGuidStrings, out arDocumentNames, out arDocumentFileNames,
                    out arDocumentDescriptions, out arDocumentUpdateDates, out arVersions, out arFileSizes, out arMimeTypes);
            else
                iCount = SearchForDocumentsMultiValuesReturningWorkflowStateSizesAsStrings(iProjectId, bSearchSubFolders, sDocumentName, sFileName, sDocumentDescription, bOriginalsOnly, iEnvironmentId,
                    sAttributeName, arAttributeValues, listValues.Count,
                    out ppProjects, out ppDocumentIds, out ppVersionSeqNumbers, out ppWorkflowIds, out ppStateIds, out ppStorageIds, out arDocumentGuidStrings, out arDocumentNames, out arDocumentFileNames,
                    out arDocumentDescriptions, out arDocumentUpdateDates, out arVersions, out arFileSizes, out arMimeTypes);
        }
        catch (Exception ex)
        {
            BPSUtilities.WriteLog("Error: {0}\n{1}", ex.Message, ex.StackTrace);
        }

        System.Diagnostics.Debug.WriteLine(string.Format("Returned {0} matches", iCount));

        DataTable dt = new DataTable("Documents");

        dt.Columns.Add("DocumentGUID", Type.GetType("System.String"));
        dt.Columns.Add("DocumentName", Type.GetType("System.String"));
        dt.Columns.Add("DocumentFileName", Type.GetType("System.String"));
        dt.Columns.Add("DocumentDescription", Type.GetType("System.String"));
        dt.Columns.Add("DocumentUpdateDate", Type.GetType("System.String"));
        dt.Columns.Add("DocumentVersion", Type.GetType("System.String"));
        dt.Columns.Add("DocumentVersionSequence", Type.GetType("System.Int32"));
        dt.Columns.Add("ProjectId", Type.GetType("System.Int32"));
        dt.Columns.Add("DocumentId", Type.GetType("System.Int32"));
        dt.Columns.Add("ProjectPath", Type.GetType("System.String"));
        dt.Columns.Add("DocumentWorkflowId", Type.GetType("System.Int32"));
        dt.Columns.Add("DocumentStateId", Type.GetType("System.Int32"));
        dt.Columns.Add("DocumentFileSize", typeof(UInt64));
        dt.Columns.Add("DocumentStorageId", Type.GetType("System.Int32"));
        dt.Columns.Add("DocumentMimeType", Type.GetType("System.String"));

        DataColumn[] pk = new DataColumn[1];
        pk[0] = dt.Columns["DocumentGUID"];
        dt.PrimaryKey = pk;

        SortedList<int, string> slProjectPaths = new SortedList<int, string>();

        if (iCount > 0)
        {
            int[] arProjects = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppProjects, iCount, out arProjects);
            int[] arDocumentIds = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppDocumentIds, iCount, out arDocumentIds);
            int[] arDocumentVersionSequenceNumbers = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppVersionSeqNumbers, iCount, out arDocumentVersionSequenceNumbers);

            int[] arWorkflowIds = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppWorkflowIds, iCount, out arWorkflowIds);
            int[] arStateIds = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppStateIds, iCount, out arStateIds);

            int[] arStorageIds = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppStorageIds, iCount, out arStorageIds);

            for (int i = 0; i < iCount; i++)
            {
                DataRow dr = dt.NewRow();

                dr["ProjectID"] = arProjects[i];
                dr["DocumentID"] = arDocumentIds[i];
                dr["DocumentGUID"] = arDocumentGuidStrings[i];
                dr["DocumentName"] = arDocumentNames[i];
                dr["DocumentFileName"] = arDocumentFileNames[i];
                dr["DocumentDescription"] = arDocumentDescriptions[i];
                dr["DocumentUpdateDate"] = arDocumentUpdateDates[i];
                dr["DocumentVersion"] = arVersions[i];
                dr["DocumentVersionSequence"] = arDocumentVersionSequenceNumbers[i];

                dr["DocumentWorkflowId"] = arWorkflowIds[i];
                dr["DocumentStateId"] = arStateIds[i];

                dr["DocumentStorageId"] = arStorageIds[i];
                dr["DocumentMimeType"] = arMimeTypes[i];

                UInt64 uiFileSize = 0;

                UInt64.TryParse(arFileSizes[i], out uiFileSize);

                dr["DocumentFileSize"] = uiFileSize;

                if (bGetPath)
                {
                    string sProjectPath = string.Empty;
                    if (!slProjectPaths.TryGetValue(arProjects[i], out sProjectPath))
                    {
                        sProjectPath = PWWrapper.GetProjectNamePath2(arProjects[i]);
                    }

                    dr["ProjectPath"] = sProjectPath;
                }

                try
                {
                    dt.Rows.Add(dr);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(string.Format("Error: {0}\n{1}", ex.Message, ex.StackTrace));
                }
            }

            // because I always forget the column names
            StringBuilder sbColumns = new StringBuilder();

            foreach (DataColumn dc in dt.Columns)
            {
                sbColumns.Append(dc.ColumnName + ";");
            }

            System.Diagnostics.Debug.WriteLine(string.Format("Columns: {0}", sbColumns.ToString()));
        }

        GC.Collect();

        return dt;
    }

    public static DataTable SearchForDocumentsByQueryIdReturningWorkflowsStatesFileSizes(int iQueryId, bool bGetPath)
    {
        IntPtr ppProjects = IntPtr.Zero;
        IntPtr ppDocumentIds = IntPtr.Zero;
        IntPtr ppVersionSeqNumbers = IntPtr.Zero;

        IntPtr ppWorkflowIds = IntPtr.Zero;
        IntPtr ppStateIds = IntPtr.Zero;
        IntPtr ppStorageIds = IntPtr.Zero;

        int iCount = 0;

        string[] arDocumentGuidStrings = null;
        string[] arDocumentNames = null;
        string[] arDocumentFileNames = null;
        string[] arDocumentDescriptions = null;
        string[] arDocumentUpdateDates = null;
        string[] arVersions = null;
        string[] arFileSizes = null;
        string[] arMimeTypes = null;

        System.Diagnostics.Debug.WriteLine("Starting query...");

        try
        {
            if (Is64Bit())
                iCount = SearchForDocumentsByQueryIdReturningWorkflowStateSizesAsStringsX64(iQueryId,
                    out ppProjects, out ppDocumentIds, out ppVersionSeqNumbers, out ppWorkflowIds, out ppStateIds, out ppStorageIds,
                    out arDocumentGuidStrings, out arDocumentNames, out arDocumentFileNames,
                    out arDocumentDescriptions, out arDocumentUpdateDates, out arVersions, out arFileSizes, out arMimeTypes);
            else
                iCount = SearchForDocumentsByQueryIdReturningWorkflowStateSizesAsStrings(iQueryId,
                    out ppProjects, out ppDocumentIds, out ppVersionSeqNumbers, out ppWorkflowIds, out ppStateIds, out ppStorageIds,
                    out arDocumentGuidStrings, out arDocumentNames, out arDocumentFileNames,
                    out arDocumentDescriptions, out arDocumentUpdateDates, out arVersions, out arFileSizes, out arMimeTypes);

        }
        catch (Exception ex)
        {
            BPSUtilities.WriteLog("Error: {0}\n{1}", ex.Message, ex.StackTrace);
        }

        System.Diagnostics.Debug.WriteLine(string.Format("Returned {0} matches", iCount));

        DataTable dt = new DataTable("Documents");

        dt.Columns.Add("DocumentGUID", Type.GetType("System.String"));
        dt.Columns.Add("DocumentName", Type.GetType("System.String"));
        dt.Columns.Add("DocumentFileName", Type.GetType("System.String"));
        dt.Columns.Add("DocumentDescription", Type.GetType("System.String"));
        dt.Columns.Add("DocumentUpdateDate", Type.GetType("System.String"));
        dt.Columns.Add("DocumentVersion", Type.GetType("System.String"));
        dt.Columns.Add("DocumentVersionSequence", Type.GetType("System.Int32"));
        dt.Columns.Add("ProjectId", Type.GetType("System.Int32"));
        dt.Columns.Add("DocumentId", Type.GetType("System.Int32"));
        dt.Columns.Add("ProjectPath", Type.GetType("System.String"));
        dt.Columns.Add("DocumentWorkflowId", Type.GetType("System.Int32"));
        dt.Columns.Add("DocumentStateId", Type.GetType("System.Int32"));
        dt.Columns.Add("DocumentFileSize", typeof(UInt64));
        dt.Columns.Add("DocumentStorageId", Type.GetType("System.Int32"));
        dt.Columns.Add("DocumentMimeType", Type.GetType("System.String"));

        DataColumn[] pk = new DataColumn[1];
        pk[0] = dt.Columns["DocumentGUID"];
        dt.PrimaryKey = pk;

        SortedList<int, string> slProjectPaths = new SortedList<int, string>();

        if (iCount > 0)
        {
            int[] arProjects = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppProjects, iCount, out arProjects);
            int[] arDocumentIds = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppDocumentIds, iCount, out arDocumentIds);
            int[] arDocumentVersionSequenceNumbers = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppVersionSeqNumbers, iCount, out arDocumentVersionSequenceNumbers);

            int[] arWorkflowIds = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppWorkflowIds, iCount, out arWorkflowIds);
            int[] arStateIds = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppStateIds, iCount, out arStateIds);
            int[] arStorageIds = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppStorageIds, iCount, out arStorageIds);

            for (int i = 0; i < iCount; i++)
            {
                DataRow dr = dt.NewRow();

                dr["ProjectID"] = arProjects[i];
                dr["DocumentID"] = arDocumentIds[i];
                dr["DocumentGUID"] = arDocumentGuidStrings[i];
                dr["DocumentName"] = arDocumentNames[i];
                dr["DocumentFileName"] = arDocumentFileNames[i];
                dr["DocumentDescription"] = arDocumentDescriptions[i];
                dr["DocumentUpdateDate"] = arDocumentUpdateDates[i];
                dr["DocumentVersion"] = arVersions[i];
                dr["DocumentVersionSequence"] = arDocumentVersionSequenceNumbers[i];

                dr["DocumentWorkflowId"] = arWorkflowIds[i];
                dr["DocumentStateId"] = arStateIds[i];

                dr["DocumentStorageId"] = arStorageIds[i];
                dr["DocumentMimeType"] = arMimeTypes[i];

                UInt64 uiFileSize = 0;
                UInt64.TryParse(arFileSizes[i], out uiFileSize);

                dr["DocumentFileSize"] = uiFileSize;

                if (bGetPath)
                {
                    string sProjectPath = string.Empty;
                    if (!slProjectPaths.TryGetValue(arProjects[i], out sProjectPath))
                    {
                        sProjectPath = PWWrapper.GetProjectNamePath2(arProjects[i]);
                    }

                    dr["ProjectPath"] = sProjectPath;
                }

                try
                {
                    dt.Rows.Add(dr);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(string.Format("Error: {0}\n{1}", ex.Message, ex.StackTrace));
                }
            }

            // because I always forget the column names
            StringBuilder sbColumns = new StringBuilder();

            foreach (DataColumn dc in dt.Columns)
            {
                sbColumns.Append(dc.ColumnName + ";");
            }

            System.Diagnostics.Debug.WriteLine(string.Format("Columns: {0}", sbColumns.ToString()));
        }

        GC.Collect();

        return dt;
    }

    [DllImport("PWSearchWrapperX64.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall, EntryPoint = "SearchForDocsWithSpatial")]
    private extern static int SearchForDocsWithSpatialX64(
        int iProjectId,
        bool bIncludeSubFolders,
        string sFullTextSearchString,
        bool bWholePhrase,
        bool bAnyWord,
        bool bSearchAttributes, // otherwise full text
        string sDocumentName,
        string sFileName,
        string sDocumentDescription,
        bool bOriginalsOnly,
        int iEnvironmentId,
        [In][MarshalAsAttribute(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] arAttributeNames,
        [In][MarshalAsAttribute(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] arAttributeValues,
        int size,
        double dLatitudeMin,
        double dLongitudeMin,
        double dLatitudeMax,
        double dLongitudeMax,
        bool bSpatialSecondPass,
        [Out] out IntPtr ppProjects,
        [Out] out IntPtr ppDocumentIds,
        [Out] out IntPtr ppVersionSeqNumbers,
        [Out] out IntPtr ppWorkflowIds,
        [Out] out IntPtr ppStateIds,
        [Out] out IntPtr ppStorageIds,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentGuidStrings,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentNames,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentFileNames,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentDescriptions,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentUpdateDates,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arVersions,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arFileSizes,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arMimeTypes
        );

    [DllImport("PWSearchWrapper.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall, EntryPoint = "SearchForDocsWithSpatial")]
    private extern static int SearchForDocsWithSpatial(
        int iProjectId,
        bool bIncludeSubFolders,
        string sFullTextSearchString,
        bool bWholePhrase,
        bool bAnyWord,
        bool bSearchAttributes, // otherwise full text
        string sDocumentName,
        string sFileName,
        string sDocumentDescription,
        bool bOriginalsOnly,
        int iEnvironmentId,
        [In][MarshalAsAttribute(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] arAttributeNames,
        [In][MarshalAsAttribute(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] arAttributeValues,
        int size,
        double dLatitudeMin,
        double dLongitudeMin,
        double dLatitudeMax,
        double dLongitudeMax,
        bool bSpatialSecondPass,
        [Out] out IntPtr ppProjects,
        [Out] out IntPtr ppDocumentIds,
        [Out] out IntPtr ppVersionSeqNumbers,
        [Out] out IntPtr ppWorkflowIds,
        [Out] out IntPtr ppStateIds,
        [Out] out IntPtr ppStorageIds,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentGuidStrings,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentNames,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentFileNames,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentDescriptions,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arDocumentUpdateDates,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arVersions,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arFileSizes,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)]
            out string[] arMimeTypes
        );

    public static DataTable SearchForDocsWithSpatial(int iProjectId, bool bSearchSubFolders,
        string sFullText, bool bWholePhrase, bool bAnyWords, bool bSearchAttributes,
        string sDocumentName, string sFileName, string sDocumentDescription, bool bOriginalsOnly,
        int iEnvironmentId,
        SortedList<string, string> slAttributes,
        double dLatitudeMin,
        double dLongitudeMin,
        double dLatitudeMax,
        double dLongitudeMax,
        bool bSpatialSecondPass,
        bool bGetPath)
    {
        IntPtr ppProjects = IntPtr.Zero;
        IntPtr ppDocumentIds = IntPtr.Zero;
        IntPtr ppVersionSeqNumbers = IntPtr.Zero;

        IntPtr ppWorkflowIds = IntPtr.Zero;
        IntPtr ppStateIds = IntPtr.Zero;
        IntPtr ppStorageIds = IntPtr.Zero;

        int iCount = 0;

        string[] arDocumentGuidStrings = null;
        string[] arDocumentNames = null;
        string[] arDocumentFileNames = null;
        string[] arDocumentDescriptions = null;
        string[] arDocumentUpdateDates = null;
        string[] arVersions = null;
        string[] arFileSizes = null;
        string[] arMimeTypes = null;

        string[] arAttributeNames = new string[slAttributes.Count];
        string[] arAttributeValues = new string[slAttributes.Count];

        int iIndex = 0;

        foreach (KeyValuePair<string, string> kvp in slAttributes)
        {
            arAttributeNames[iIndex] = kvp.Key;
            arAttributeValues[iIndex] = kvp.Value;
            iIndex++;
        }

        System.Diagnostics.Debug.WriteLine("Starting query...");

        try
        {
            if (Is64Bit())
                iCount = SearchForDocsWithSpatialX64(iProjectId, bSearchSubFolders, sFullText, bWholePhrase,
                    bAnyWords, bSearchAttributes, sDocumentName, sFileName, sDocumentDescription, bOriginalsOnly, iEnvironmentId,
                    arAttributeNames, arAttributeValues, slAttributes.Count,
                    dLatitudeMin, dLongitudeMin, dLatitudeMax, dLongitudeMax, bSpatialSecondPass,
                    out ppProjects, out ppDocumentIds, out ppVersionSeqNumbers, out ppWorkflowIds, out ppStateIds, out ppStorageIds, out arDocumentGuidStrings, out arDocumentNames, out arDocumentFileNames,
                    out arDocumentDescriptions, out arDocumentUpdateDates, out arVersions, out arFileSizes, out arMimeTypes);
            else
                iCount = SearchForDocsWithSpatial(iProjectId, bSearchSubFolders, sFullText, bWholePhrase,
                    bAnyWords, bSearchAttributes, sDocumentName, sFileName, sDocumentDescription, bOriginalsOnly, iEnvironmentId,
                    arAttributeNames, arAttributeValues, slAttributes.Count,
                    dLatitudeMin, dLongitudeMin, dLatitudeMax, dLongitudeMax, bSpatialSecondPass,
                    out ppProjects, out ppDocumentIds, out ppVersionSeqNumbers, out ppWorkflowIds, out ppStateIds, out ppStorageIds, out arDocumentGuidStrings, out arDocumentNames, out arDocumentFileNames,
                    out arDocumentDescriptions, out arDocumentUpdateDates, out arVersions, out arFileSizes, out arMimeTypes);
        }
        catch (Exception ex)
        {
            BPSUtilities.WriteLog("Error: {0}\n{1}", ex.Message, ex.StackTrace);
        }

        System.Diagnostics.Debug.WriteLine(string.Format("Returned {0} matches", iCount));

        DataTable dt = new DataTable("Documents");

        dt.Columns.Add("DocumentGUID", Type.GetType("System.String"));
        dt.Columns.Add("DocumentName", Type.GetType("System.String"));
        dt.Columns.Add("DocumentFileName", Type.GetType("System.String"));
        dt.Columns.Add("DocumentDescription", Type.GetType("System.String"));
        dt.Columns.Add("DocumentUpdateDate", Type.GetType("System.String"));
        dt.Columns.Add("DocumentVersion", Type.GetType("System.String"));
        dt.Columns.Add("DocumentVersionSequence", Type.GetType("System.Int32"));
        dt.Columns.Add("ProjectId", Type.GetType("System.Int32"));
        dt.Columns.Add("DocumentId", Type.GetType("System.Int32"));
        dt.Columns.Add("ProjectPath", Type.GetType("System.String"));
        dt.Columns.Add("DocumentWorkflowId", Type.GetType("System.Int32"));
        dt.Columns.Add("DocumentStateId", Type.GetType("System.Int32"));
        dt.Columns.Add("DocumentFileSize", typeof(UInt64));
        dt.Columns.Add("DocumentStorageId", Type.GetType("System.Int32"));
        dt.Columns.Add("DocumentMimeType", Type.GetType("System.String"));

        DataColumn[] pk = new DataColumn[1];
        pk[0] = dt.Columns["DocumentGUID"];
        dt.PrimaryKey = pk;

        SortedList<int, string> slProjectPaths = new SortedList<int, string>();

        if (iCount > 0)
        {
            int[] arProjects = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppProjects, iCount, out arProjects);
            int[] arDocumentIds = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppDocumentIds, iCount, out arDocumentIds);
            int[] arDocumentVersionSequenceNumbers = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppVersionSeqNumbers, iCount, out arDocumentVersionSequenceNumbers);

            int[] arWorkflowIds = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppWorkflowIds, iCount, out arWorkflowIds);
            int[] arStateIds = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppStateIds, iCount, out arStateIds);

            int[] arStorageIds = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppStorageIds, iCount, out arStorageIds);

            for (int i = 0; i < iCount; i++)
            {
                DataRow dr = dt.NewRow();

                dr["ProjectID"] = arProjects[i];
                dr["DocumentID"] = arDocumentIds[i];
                dr["DocumentGUID"] = arDocumentGuidStrings[i];
                dr["DocumentName"] = arDocumentNames[i];
                dr["DocumentFileName"] = arDocumentFileNames[i];
                dr["DocumentDescription"] = arDocumentDescriptions[i];
                dr["DocumentUpdateDate"] = arDocumentUpdateDates[i];
                dr["DocumentVersion"] = arVersions[i];
                dr["DocumentVersionSequence"] = arDocumentVersionSequenceNumbers[i];

                dr["DocumentWorkflowId"] = arWorkflowIds[i];
                dr["DocumentStateId"] = arStateIds[i];

                dr["DocumentStorageId"] = arStorageIds[i];
                dr["DocumentMimeType"] = arMimeTypes[i];

                UInt64 uiFileSize = 0;

                UInt64.TryParse(arFileSizes[i], out uiFileSize);

                dr["DocumentFileSize"] = uiFileSize;

                if (bGetPath)
                {
                    string sProjectPath = string.Empty;
                    if (!slProjectPaths.TryGetValue(arProjects[i], out sProjectPath))
                    {
                        sProjectPath = PWWrapper.GetProjectNamePath2(arProjects[i]);
                    }

                    dr["ProjectPath"] = sProjectPath;
                }

                try
                {
                    dt.Rows.Add(dr);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(string.Format("Error: {0}\n{1}", ex.Message, ex.StackTrace));
                }
            }

            // because I always forget the column names
            StringBuilder sbColumns = new StringBuilder();

            foreach (DataColumn dc in dt.Columns)
            {
                sbColumns.Append(dc.ColumnName + ";");
            }

            System.Diagnostics.Debug.WriteLine(string.Format("Columns: {0}", sbColumns.ToString()));
        }

        GC.Collect();

        return dt;
    }

    public static DataTable SearchForDocumentsUltimate(int iProjectId, // project id
        bool bSearchSubFolders, // search subfolders
        string sFullText, // text string to search for
        bool bWholePhrase, // search for the whole phrase
        bool bAnyWords, // search for any words in the string
        bool bSearchAttributes, // also search in the attributes
        string sDocumentName, // document name to search for (supports % wildcard)
        string sFileName, // file name to search for (supports % wildcard)
        string sDocumentDescription, // description to search for (supports % wildcard)
        bool bOriginalsOnly, // only originals, no versions
        SortedList<string, string> slAttributes, // attributes and values to search for
        double dLatitudeMin,
        double dLongitudeMin,
        double dLatitudeMax,
        double dLongitudeMax,
        bool bSpatialSecondPass,
        List<int> listEnvironments,
        List<int> listStates,
        List<int> listWorkflows,
        List<int> listStorages,
        List<int> listApplications,
        List<string> listStatuses,
        List<int> listItemTypes,
        List<int> listCreators,
        List<int> listUpdaters,
        List<int> listCheckedOutUsers,
        int iFinalStatus,
        string sFileUpdatedAfter, // early date 2009-10-22 01:00:00
        string sFileUpdatedBefore, // late date 2010-10-22 01:00:00
        string sDocUpdatedAfter, // early date 2009-10-22 01:00:00
        string sDocUpdatedBefore, // late date 2010-10-22 01:00:00
        string sDocCreatedAfter, // early date 2009-10-22 01:00:00
        string sDocCreatedBefore, // late date 2010-10-22 01:00:00
        string sDocCheckedOutAfter, // early date 2009-10-22 01:00:00
        string sDocCheckedOutBefore, // late date 2010-10-22 01:00:00
        bool bGetPath,
        List<string> listColumns, // return attribute columns
        string sQueryName, // if you want to save the query
        int iQueryParentProjectId, // if you want to hang it on a project
        int iParentQueryId // if you want to make a sub-query
        )
    {
        IntPtr ppProjects = IntPtr.Zero;
        IntPtr ppDocumentIds = IntPtr.Zero;
        IntPtr ppVersionSeqNumbers = IntPtr.Zero;
        IntPtr ppVersionOriginalNumbers = IntPtr.Zero;

        IntPtr ppWorkflowIds = IntPtr.Zero;
        IntPtr ppStateIds = IntPtr.Zero;
        IntPtr ppStorageIds = IntPtr.Zero;

        IntPtr ppUpdaterIds = IntPtr.Zero;
        IntPtr ppCreatorIds = IntPtr.Zero;
        IntPtr ppCheckedOutIds = IntPtr.Zero;
        IntPtr ppApplicationIds = IntPtr.Zero;
        IntPtr ppItemTypes = IntPtr.Zero;
        IntPtr ppFinalStatuses = IntPtr.Zero;

        int iCount = 0;

        string[] arDocumentGuidStrings = null;
        string[] arDocumentNames = null;
        string[] arDocumentFileNames = null;
        string[] arDocumentDescriptions = null;
        string[] arDocumentUpdateDates = null;
        string[] arFileUpdateDates = null;
        string[] arDocumentCreateDates = null;
        string[] arDocumentCheckedOutDates = null;
        string[] arVersions = null;
        string[] arFileSizes = null;
        string[] arMimeTypes = null;
        string[] arStatuses = null;
        string[] arAttributes = null;

        string[] arAttributeNames = new string[slAttributes.Count];
        string[] arAttributeValues = new string[slAttributes.Count];

        int iIndex = 0;

        foreach (KeyValuePair<string, string> kvp in slAttributes)
        {
            arAttributeNames[iIndex] = kvp.Key;
            arAttributeValues[iIndex] = kvp.Value;
            iIndex++;
        }

        string sStates = string.Join(",", listStates);
        string sWorkflows = string.Join(",", listWorkflows);
        string sStorages = string.Join(",", listStorages);
        string sStatuses = string.Join(",", listStatuses);
        string sItemTypes = string.Join(",", listItemTypes);
        string sCreators = string.Join(",", listCreators);
        string sUpdaters = string.Join(",", listUpdaters);
        string sCheckerOuters = string.Join(",", listCheckedOutUsers);
        string sApplications = string.Join(",", listApplications);
        string sEnvironments = string.Join(",", listEnvironments);

        System.Diagnostics.Debug.WriteLine("Starting query...");

        string sDelimiter = "^";

        try
        {
            if (Is64Bit())
                iCount = SearchForDocumentsUltimateX64(iProjectId, bSearchSubFolders, sFullText, bWholePhrase,
                    bAnyWords, bSearchAttributes, sDocumentName, sFileName, sDocumentDescription, bOriginalsOnly,
                    arAttributeNames, arAttributeValues, slAttributes.Count,
                    sEnvironments,
                    sStates,
                    sWorkflows,
                    sFileUpdatedAfter,
                    sFileUpdatedBefore,
                    sDocUpdatedAfter,
                    sDocUpdatedBefore,
                    sDocCreatedAfter,
                    sDocCreatedBefore,
                    sDocCheckedOutAfter,
                    sDocCheckedOutBefore,
                    dLatitudeMin, dLongitudeMin, dLatitudeMax, dLongitudeMax, bSpatialSecondPass,
                    sStorages,
                    sStatuses,
                    sItemTypes,
                    iFinalStatus,
                    sCreators,
                    sUpdaters,
                    sCheckerOuters,
                    sApplications,
                    listColumns.ToArray(),
                    listColumns.Count,
                    sDelimiter,
                    sQueryName,
                    iQueryParentProjectId,
                    iParentQueryId,
                    out ppProjects,
                    out ppDocumentIds,
                    out ppVersionSeqNumbers,
                    out ppVersionOriginalNumbers,
                    out ppWorkflowIds,
                    out ppStateIds,
                    out ppStorageIds,
                    out ppCreatorIds,
                    out ppUpdaterIds,
                    out ppCheckedOutIds,
                    out ppItemTypes,
                    out ppApplicationIds,
                    out arDocumentGuidStrings,
                    out arDocumentNames,
                    out arDocumentFileNames,
                    out arDocumentDescriptions,
                    out arDocumentUpdateDates,
                    out arFileUpdateDates,
                    out arDocumentCreateDates,
                    out arDocumentCheckedOutDates,
                    out arVersions,
                    out arFileSizes,
                    out arStatuses,
                    out arMimeTypes,
                    out arAttributes);
            else
                iCount = SearchForDocumentsUltimate(iProjectId, bSearchSubFolders, sFullText, bWholePhrase,
                    bAnyWords, bSearchAttributes, sDocumentName, sFileName, sDocumentDescription, bOriginalsOnly,
                    arAttributeNames, arAttributeValues, slAttributes.Count,
                    sEnvironments,
                    sStates,
                    sWorkflows,
                    sFileUpdatedAfter,
                    sFileUpdatedBefore,
                    sDocUpdatedAfter,
                    sDocUpdatedBefore,
                    sDocCreatedAfter,
                    sDocCreatedBefore,
                    sDocCheckedOutAfter,
                    sDocCheckedOutBefore,
                    dLatitudeMin, dLongitudeMin, dLatitudeMax, dLongitudeMax, bSpatialSecondPass,
                    sStorages,
                    sStatuses,
                    sItemTypes,
                    iFinalStatus,
                    sCreators,
                    sUpdaters,
                    sCheckerOuters,
                    sApplications,
                    listColumns.ToArray(),
                    listColumns.Count,
                    sDelimiter,
                    sQueryName,
                    iQueryParentProjectId,
                    iParentQueryId,
                    out ppProjects,
                    out ppDocumentIds,
                    out ppVersionSeqNumbers,
                    out ppVersionOriginalNumbers,
                    out ppWorkflowIds,
                    out ppStateIds,
                    out ppStorageIds,
                    out ppCreatorIds,
                    out ppUpdaterIds,
                    out ppCheckedOutIds,
                    out ppItemTypes,
                    out ppApplicationIds,
                    out arDocumentGuidStrings,
                    out arDocumentNames,
                    out arDocumentFileNames,
                    out arDocumentDescriptions,
                    out arDocumentUpdateDates,
                    out arFileUpdateDates,
                    out arDocumentCreateDates,
                    out arDocumentCheckedOutDates,
                    out arVersions,
                    out arFileSizes,
                    out arStatuses,
                    out arMimeTypes,
                    out arAttributes);
        }
        catch (Exception ex)
        {
            BPSUtilities.WriteLog("Error: {0}\n{1}", ex.Message, ex.StackTrace);
        }

        System.Diagnostics.Debug.WriteLine(string.Format("Returned {0} matches", iCount));

        DataTable dt = new DataTable("Documents");

        dt.Columns.Add("DocumentGUID", Type.GetType("System.String"));  // 0
        dt.Columns.Add("DocumentName", Type.GetType("System.String"));  // 1
        dt.Columns.Add("DocumentFileName", Type.GetType("System.String")); // 2
        dt.Columns.Add("DocumentDescription", Type.GetType("System.String")); // 3
        dt.Columns.Add("FileUpdateDate", Type.GetType("System.String")); // 4
        dt.Columns.Add("DocumentUpdateDate", Type.GetType("System.String")); // 5
        dt.Columns.Add("ProjectId", Type.GetType("System.Int32")); // 6
        dt.Columns.Add("DocumentId", Type.GetType("System.Int32")); // 7
        dt.Columns.Add("DocumentVersion", Type.GetType("System.String")); // 8
        dt.Columns.Add("DocumentVersionSequence", Type.GetType("System.Int32")); // 9
        dt.Columns.Add("DocumentOriginalNo", Type.GetType("System.Int32")); // 10
        dt.Columns.Add("DocumentWorkflowId", Type.GetType("System.Int32")); // 11
        dt.Columns.Add("DocumentStateId", Type.GetType("System.Int32")); // 12
        dt.Columns.Add("DocumentStorageId", Type.GetType("System.Int32")); // 13
        dt.Columns.Add("DocumentCreatorId", Type.GetType("System.Int32")); // 14
        dt.Columns.Add("DocumentUpdaterId", Type.GetType("System.Int32")); // 15
        dt.Columns.Add("DocumentItemType", Type.GetType("System.Int32"));  // 16
        dt.Columns.Add("DocumentApplicationId", Type.GetType("System.Int32")); // 17
        dt.Columns.Add("DocumentCreatedDate", Type.GetType("System.String")); // 18
        dt.Columns.Add("DocumentFileSize", typeof(UInt64));             // 19
        dt.Columns.Add("DocumentStatus", Type.GetType("System.String")); // 20
        dt.Columns.Add("DocumentMimeType", Type.GetType("System.String")); // 21
        dt.Columns.Add("DocumentCheckOutUserId", Type.GetType("System.Int32")); // 22
        dt.Columns.Add("DocumentCheckOutDate", Type.GetType("System.String")); // 23

        dt.Columns.Add("ProjectPath", Type.GetType("System.String"));

        foreach (string sColumnName in listColumns)
        {
            if (!dt.Columns.Contains(sColumnName))
                dt.Columns.Add(sColumnName, typeof(string));
        }

        DataColumn[] pk = new DataColumn[1];
        pk[0] = dt.Columns["DocumentGUID"];
        dt.PrimaryKey = pk;

        SortedList<int, string> slProjectPaths = new SortedList<int, string>();

        if (iCount > 0)
        {
            int[] arProjects = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppProjects, iCount, out arProjects);
            int[] arDocumentIds = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppDocumentIds, iCount, out arDocumentIds);
            int[] arDocumentVersionSequenceNumbers = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppVersionSeqNumbers, iCount, out arDocumentVersionSequenceNumbers);
            int[] arWorkflowIds = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppWorkflowIds, iCount, out arWorkflowIds);
            int[] arStateIds = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppStateIds, iCount, out arStateIds);
            int[] arStorageIds = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppStorageIds, iCount, out arStorageIds);
            int[] arVersionOriginalNos = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppVersionOriginalNumbers, iCount, out arVersionOriginalNos);
            int[] arItemTypes = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppItemTypes, iCount, out arItemTypes);
            int[] arCreatorIds = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppCreatorIds, iCount, out arCreatorIds);
            int[] arUpdaterIds = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppUpdaterIds, iCount, out arUpdaterIds);
            int[] arApplicationIds = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppApplicationIds, iCount, out arApplicationIds);
            int[] arCheckedOutIds = new int[iCount];
            MarshalUnmananagedIntArrayToManagedIntArray(ppCheckedOutIds, iCount, out arCheckedOutIds);

            for (int i = 0; i < iCount; i++)
            {
                DataRow dr = dt.NewRow();

                dr["ProjectID"] = arProjects[i];
                dr["DocumentID"] = arDocumentIds[i];
                dr["DocumentGUID"] = arDocumentGuidStrings[i];
                dr["DocumentName"] = arDocumentNames[i];
                dr["DocumentFileName"] = arDocumentFileNames[i];
                dr["DocumentDescription"] = arDocumentDescriptions[i];
                dr["DocumentUpdateDate"] = arDocumentUpdateDates[i];
                dr["DocumentVersion"] = arVersions[i];
                dr["DocumentVersionSequence"] = arDocumentVersionSequenceNumbers[i];
                dr["DocumentWorkflowId"] = arWorkflowIds[i];
                dr["DocumentStateId"] = arStateIds[i];
                dr["DocumentStorageId"] = arStorageIds[i];
                if (arMimeTypes != null)
                    dr["DocumentMimeType"] = arMimeTypes[i];
                UInt64 uiFileSize = 0;
                if (arFileSizes != null)
                    UInt64.TryParse(arFileSizes[i], out uiFileSize);
                dr["DocumentFileSize"] = uiFileSize;

                dr["FileUpdateDate"] = arFileUpdateDates[i];
                dr["DocumentOriginalNo"] = arVersionOriginalNos[i];
                dr["DocumentCreatorId"] = arCreatorIds[i];
                dr["DocumentUpdaterId"] = arUpdaterIds[i];
                dr["DocumentItemType"] = arItemTypes[i];
                dr["DocumentApplicationId"] = arApplicationIds[i];
                dr["DocumentCreatedDate"] = arDocumentCreateDates[i];
                dr["DocumentStatus"] = arStatuses[i];

                dr["DocumentCheckOutUserId"] = arCheckedOutIds[i];
                dr["DocumentCheckOutDate"] = arDocumentCheckedOutDates[i];

                if (bGetPath)
                {
                    string sProjectPath = string.Empty;
                    if (!slProjectPaths.TryGetValue(arProjects[i], out sProjectPath))
                    {
                        sProjectPath = PWWrapper.GetProjectNamePath2(arProjects[i]);
                        slProjectPaths.AddWithCheck(arProjects[i], sProjectPath);
                    }

                    dr["ProjectPath"] = sProjectPath;
                }

                if (arAttributes != null)
                {
                    if (arAttributes.Length >= iCount)
                    {
                        if (!string.IsNullOrEmpty(arAttributes[i]))
                        {
                            string[] sAdditionalValues = arAttributes[i].Split(sDelimiter.ToCharArray());

                            if (sAdditionalValues.Length == listColumns.Count)
                            {
                                for (int k = 0; k < listColumns.Count; k++)
                                {
                                    if (dt.Columns.Contains(listColumns[k]))
                                    {
                                        dr[listColumns[k]] = sAdditionalValues[k].Trim();
                                    }
                                }
                            }
                            else
                            {
                                System.Diagnostics.Debug.WriteLine(string.Format("Additional values length was {0} while list of columns contained {1} values.",
                                    sAdditionalValues.Length, listColumns.Count));
                            }
                        }
                    }
                }

                try
                {
                    dt.Rows.Add(dr);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(string.Format("Error: {0}\n{1}", ex.Message, ex.StackTrace));
                }
            }

            // because I always forget the column names
            StringBuilder sbColumns = new StringBuilder();

            foreach (DataColumn dc in dt.Columns)
            {
                sbColumns.Append(dc.ColumnName + ";");
            }

            System.Diagnostics.Debug.WriteLine(string.Format("Columns: {0}", sbColumns.ToString()));
        }

        GC.Collect();

        return dt;
    }
}

public class BPSUtilities
{

    public static string MIXED_CASE = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890";
    public static string LOWER_CASE = "abcdefghijklmnopqrstuvwxyz1234567890";

#if !IS_NET35
    public static string GetARandomString(int length, string characterSet)
    {
        StringBuilder sb = new StringBuilder();

        using (System.Security.Cryptography.RNGCryptoServiceProvider provider = new System.Security.Cryptography.RNGCryptoServiceProvider())
        {
            while (sb.Length != length)
            {
                byte[] oneByte = new byte[1];
                provider.GetBytes(oneByte);
                char character = (char)oneByte[0];
                if (characterSet.Contains(character.ToString()))
                {
                    // s += character;
                    sb.Append(character);
                }
            }
        }

        return sb.ToString();
    }
    public static string GetARandomString(int length)
    {
        StringBuilder sb = new StringBuilder();

        using (System.Security.Cryptography.RNGCryptoServiceProvider provider = new System.Security.Cryptography.RNGCryptoServiceProvider())
        {
            while (sb.Length != length)
            {
                byte[] oneByte = new byte[1];
                provider.GetBytes(oneByte);
                char character = (char)oneByte[0];
                if (LOWER_CASE.Contains(character.ToString()))
                {
                    // s += character;
                    sb.Append(character);
                }
            }
        }

        return sb.ToString();
    }
#endif

    public static string DecodeAndValidateToken(string sPossiblyEncodedToken)
    {
#if NETCOREAPP3_1
        // can't figure out how to redo this in .NET Core yet, basically will return unencoded token without validation

        if (sPossiblyEncodedToken.ToLower().StartsWith("token "))
            sPossiblyEncodedToken = sPossiblyEncodedToken.Substring("token ".Length);

        string sUnencodedToken = sPossiblyEncodedToken;

        try
        {
            sUnencodedToken = Encoding.UTF8.GetString(Convert.FromBase64String(sPossiblyEncodedToken));
        }
        catch
        {
            sUnencodedToken = sPossiblyEncodedToken;
        }

        return sUnencodedToken;

#else
        try
        {
            System.IdentityModel.Tokens.SecurityToken token = null;
            var handlers = System.IdentityModel.Tokens.SecurityTokenHandlerCollection.CreateDefaultSecurityTokenHandlerCollection();

            if (sPossiblyEncodedToken.ToLower().StartsWith("token "))
                sPossiblyEncodedToken = sPossiblyEncodedToken.Substring("token ".Length);

            string sUnencodedToken = sPossiblyEncodedToken;

            try
            {
                sUnencodedToken = Encoding.UTF8.GetString(Convert.FromBase64String(sPossiblyEncodedToken));
            }
            catch
            {
                sUnencodedToken = sPossiblyEncodedToken;
            }

            using (var reader = System.Xml.XmlReader.Create(new StringReader(sUnencodedToken)))
            {
                if (handlers.CanReadToken(reader))
                    token = handlers.ReadToken(reader);
            }

            if (null != token)
            {
                // means the unencoded token was valid SAML
                return sUnencodedToken;
            }
        }
        catch (Exception ex)
        {
            BPSUtilities.WriteLog(ex.Message);
        }

        return string.Empty;

#endif
    }

    public static bool GetEmbeddedResourceFile(string sResourceName, string sTargetFile)
    {
        System.Reflection.Assembly thisExe = System.Reflection.Assembly.GetExecutingAssembly();

        if (File.Exists(sTargetFile))
        {
            try
            {
                File.Delete(sTargetFile);
            }
            catch
            {
            }
        }

        try
        {
            using (var resourceStream = thisExe.GetManifestResourceStream(sResourceName))
            {
                byte[] buf = new byte[resourceStream.Length];
                resourceStream.Read(buf, 0, buf.Length);
                File.WriteAllBytes(sTargetFile, buf);
            }
        }
        catch (Exception ex)
        {
            WriteLog("Error: {0}\n{1}", ex.Message, ex.StackTrace);
        }

        return File.Exists(sTargetFile);
    }

    private static readonly object _syncObject = new object();

    public static void WriteLog(string sMessage)
    {
        System.Diagnostics.Debug.WriteLine(sMessage);
        Console.WriteLine(sMessage);

        string sLogFolder = GetLogFolder(); //
        string sAppPath = System.Reflection.Assembly.GetCallingAssembly().Location;

        string sOutputFileName = string.Format("{3}_{0:D4}{1:D2}{2:D2}.log",
                DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day,
                System.IO.Path.GetFileNameWithoutExtension(sAppPath));

        sOutputFileName = Path.Combine(sLogFolder, sOutputFileName);

        if (!string.IsNullOrEmpty(sOutputFileName))
        {
            try
            {
                // only one thread can own this lock, so other threads
                // entering this method will wait here until lock is
                // available.
                lock (_syncObject)
                {
                    using (System.IO.StreamWriter sw = new System.IO.StreamWriter(sOutputFileName, true))
                    {
                        // Add some text to the file.
                        sw.WriteLine(string.Format("{0:u}\t{1}", DateTime.Now, sMessage));
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(ex.Message);
                string sMsg = string.Format("{0:u}\t{1}", DateTime.Now, sMessage);
                System.Diagnostics.Debug.WriteLine("Message in error: " + sMsg);
            }
            finally
            {
            }
        }
    }

    //added MDS 11/1/2011
    private static void WriteLogToThisFile(string sMessage, string sFullPathToLogFile)
    {
        if (!string.IsNullOrEmpty(sFullPathToLogFile))
        {
            try
            {
                // only one thread can own this lock, so other threads
                // entering this method will wait here until lock is
                // available.
                lock (_syncObject)
                {
                    using (System.IO.StreamWriter sw = new System.IO.StreamWriter(sFullPathToLogFile, true))
                    {
                        // Add some text to the file.
                        sw.WriteLine(string.Format("{0:u}\t{1}", DateTime.Now, sMessage));
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(ex.Message);
            }
            finally
            {
            }
        }
    }

    //added MDS 11/1/2011
    public static void WriteLogToFile(string sMessage, string sLogFilePath)
    {
        System.Diagnostics.Debug.WriteLine(sMessage);
        Console.WriteLine(sMessage);

        string sAppPath = System.Reflection.Assembly.GetCallingAssembly().Location;

        string sOutputFileName = string.Format("{4}\\{3}_{0:D4}{1:D2}{2:D2}.log",
                DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day,
                System.IO.Path.GetFileNameWithoutExtension(sAppPath),
                sLogFilePath);

        if (!string.IsNullOrEmpty(sOutputFileName))
        {
            try
            {
                // only one thread can own this lock, so other threads
                // entering this method will wait here until lock is
                // available.
                lock (_syncObject)
                {
                    using (System.IO.StreamWriter sw = new System.IO.StreamWriter(sOutputFileName, true))
                    {
                        // Add some text to the file.
                        sw.WriteLine(string.Format("{0:u}\t{1}", DateTime.Now, sMessage));
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(ex.Message);
            }
            finally
            {
            }
        }
    }

    //added MDS 11/1/2011
    public static void WriteLog(string sMessage, params object[] args)
    {
        string sLogFolder = GetLogFolder(); //
        string sAppPath = System.Reflection.Assembly.GetCallingAssembly().Location;

        string sOutputFileName = string.Format("{3}_{0:D4}{1:D2}{2:D2}.log",
                DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day,
                System.IO.Path.GetFileNameWithoutExtension(sAppPath));

        sOutputFileName = Path.Combine(sLogFolder, sOutputFileName);

        try
        {
            // will error here if format string is bad.
            System.Diagnostics.Debug.WriteLine(string.Format(sMessage, args));
            Console.WriteLine(string.Format(sMessage, args));

            //if (!string.IsNullOrEmpty(GetSetting("LogFolder")))
            //    sOutputFileName = Path.Combine(GetSetting("LogFolder"), sOutputFileName);

            if (!string.IsNullOrEmpty(sOutputFileName))
            {
                try
                {
                    // only one thread can own this lock, so other threads
                    // entering this method will wait here until lock is
                    // available.
                    lock (_syncObject)
                    {
                        using (System.IO.StreamWriter sw = new System.IO.StreamWriter(sOutputFileName, true))
                        {
                            //Add some text to the file.
                            string sMsg = string.Format("{0:u}\t{1}", DateTime.Now, string.Format(sMessage, args));
                            sw.WriteLine(sMsg);
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(ex.Message);

                    try
                    {
                        string sMsg = string.Format("{0:u}\t{1}", DateTime.Now, string.Format(sMessage, args));
                        System.Diagnostics.Debug.WriteLine("Message in error: " + sMsg);
                    }
                    catch (Exception ex2)
                    {
                        System.Diagnostics.Debug.WriteLine(string.Format("Error: {0}\n{1}", ex2.Message, ex2.StackTrace));
                    }
                }
                finally
                {
                }
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(string.Format("Error: {0}\n{1}", ex.Message, ex.StackTrace));
            Console.WriteLine(string.Format("Error: {0}\n{1}", ex.Message, ex.StackTrace));
        }
    }

    public static void WriteLogError(string sMessage, params object[] args)
    {
        string sLogFolder = GetLogFolder(); //
        string sAppPath = System.Reflection.Assembly.GetCallingAssembly().Location;

        string sOutputFileName = string.Format("{3}_{0:D4}{1:D2}{2:D2}.error.log",
                DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day,
                System.IO.Path.GetFileNameWithoutExtension(sAppPath));

        sOutputFileName = Path.Combine(sLogFolder, sOutputFileName);

        try
        {
            // will error here if format string is bad.
            System.Diagnostics.Debug.WriteLine("ERROR: " + string.Format(sMessage, args));
            Console.WriteLine("ERROR: " + string.Format(sMessage, args));

            // if (!string.IsNullOrEmpty(GetSetting("LogFolder")))
            // sOutputFileName = Path.Combine(GetSetting("LogFolder"), sOutputFileName);

            if (!string.IsNullOrEmpty(sOutputFileName))
            {
                try
                {
                    // only one thread can own this lock, so other threads
                    // entering this method will wait here until lock is
                    // available.
                    lock (_syncObject)
                    {
                        using (System.IO.StreamWriter sw = new System.IO.StreamWriter(sOutputFileName, true))
                        {
                            //Add some text to the file.
                            string sMsg = string.Format("{0:u}\t{1}", DateTime.Now, string.Format(sMessage, args));
                            sw.WriteLine(sMsg);
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(ex.Message);

                    try
                    {
                        string sMsg = string.Format("{0:u}\t{1}", DateTime.Now, string.Format(sMessage, args));
                        System.Diagnostics.Debug.WriteLine("Message in error: " + sMsg);
                    }
                    catch (Exception ex2)
                    {
                        System.Diagnostics.Debug.WriteLine(string.Format("Error: {0}\n{1}", ex2.Message, ex2.StackTrace));
                    }
                }
                finally
                {
                }
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(string.Format("Error: {0}\n{1}", ex.Message, ex.StackTrace));
            Console.WriteLine(string.Format("Error: {0}\n{1}", ex.Message, ex.StackTrace));
        }
    }

    private static string GetLogFolder()
    {
        string sLogPath = Path.GetDirectoryName(System.Reflection.Assembly.GetCallingAssembly().Location);

        try
        {
            // should mean is Windows app and we want to put in special folder
            if (string.IsNullOrEmpty(Console.Title))
            {
                sLogPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    @"Bentley\Logs");
            }
        }
        catch // (Exception ex)
        {
            sLogPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                @"Bentley\Logs");
        }

#if (NET45)
        // LogFolder setting should override
        if (!string.IsNullOrEmpty(GetSetting("LogFolder")))
        {
            sLogPath = GetSetting("LogFolder");
        }
#endif

#if (USE_LOG_FOLDER_SETTING)
        // LogFolder setting should override
        if (!string.IsNullOrEmpty(GetSetting("LogFolder")))
        {
            sLogPath = GetSetting("LogFolder");
        }
#endif

        if (string.IsNullOrEmpty(sLogPath))
        {
            sLogPath = Path.GetDirectoryName(System.Reflection.Assembly.GetCallingAssembly().Location);
        }

        if (!string.IsNullOrEmpty(sLogPath))
        {
            if (!Directory.Exists(sLogPath))
            {
                try
                {
                    Directory.CreateDirectory(sLogPath);
                }
                catch //  (Exception ex)
                {
                    sLogPath = Path.GetDirectoryName(System.Reflection.Assembly.GetCallingAssembly().Location);
                }
            }
        }

        return sLogPath;
    }

    public static void LogInfo(string sMessage, params object[] args)
    {
        string sLogFolder = GetLogFolder(); //
        string sAppPath = System.Reflection.Assembly.GetCallingAssembly().Location;

        string sOutputFileName = string.Format("{3}_{0:D4}{1:D2}{2:D2}.log",
                DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day,
                System.IO.Path.GetFileNameWithoutExtension(sAppPath));

        sOutputFileName = Path.Combine(sLogFolder, sOutputFileName);

        System.Diagnostics.Debug.WriteLine(string.Format(sMessage, args));
        Console.WriteLine("[Info   ] " + string.Format(sMessage, args));

        //if (!string.IsNullOrEmpty(GetSetting("LogFolder")))
        //    sOutputFileName = Path.Combine(GetSetting("LogFolder"), sOutputFileName);

        WriteLogToThisFile("[Info   ] " + string.Format(sMessage, args), sOutputFileName);
    }

    public static void LogError(string sMessage, params object[] args)
    {
        string sLogFolder = GetLogFolder(); //
        string sAppPath = System.Reflection.Assembly.GetCallingAssembly().Location;

        string sOutputFileName = string.Format("{3}_{0:D4}{1:D2}{2:D2}.log",
                DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day,
                System.IO.Path.GetFileNameWithoutExtension(sAppPath));

        sOutputFileName = Path.Combine(sLogFolder, sOutputFileName);

        System.Diagnostics.Debug.WriteLine(string.Format(sMessage, args));
        Console.WriteLine("[Error  ] " + string.Format(sMessage, args));

        //if (!string.IsNullOrEmpty(GetSetting("LogFolder")))
        //    sOutputFileName = Path.Combine(GetSetting("LogFolder"), sOutputFileName);

        WriteLogToThisFile("[Error  ] " + string.Format(sMessage, args), sOutputFileName);
    }

    public static void LogWarning(string sMessage, params object[] args)
    {
        string sLogFolder = GetLogFolder(); //
        string sAppPath = System.Reflection.Assembly.GetCallingAssembly().Location;

        string sOutputFileName = string.Format("{3}_{0:D4}{1:D2}{2:D2}.log",
                DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day,
                System.IO.Path.GetFileNameWithoutExtension(sAppPath));

        sOutputFileName = Path.Combine(sLogFolder, sOutputFileName);

        System.Diagnostics.Debug.WriteLine(string.Format(sMessage, args));
        Console.WriteLine("[Warning] " + string.Format(sMessage, args));

        //if (!string.IsNullOrEmpty(GetSetting("LogFolder")))
        //    sOutputFileName = Path.Combine(GetSetting("LogFolder"), sOutputFileName);

        WriteLogToThisFile("[Warning] " + string.Format(sMessage, args), sOutputFileName);
    }

    public static SortedList<string, string> BuildListFromString(string sList, string sDelimiter)
    {
        SortedList<string, string> sl = new SortedList<string, string>(StringComparer.CurrentCultureIgnoreCase);

        string[] sParts = sList.Split(sDelimiter.ToCharArray(), StringSplitOptions.RemoveEmptyEntries);

        for (int i = 0; i < sParts.Length; i++)
        {
            if (!sl.ContainsKey(sParts[i]))
                sl.Add(sParts[i], sParts[i]);
        }

        return sl;
    }

    public static Hashtable BuildHashTableFromString(string sList, string sDelimiter)
    {
        Hashtable ht = new Hashtable();

        string[] sParts = sList.Split(sDelimiter.ToCharArray(), StringSplitOptions.RemoveEmptyEntries);

        for (int i = 0; i < sParts.Length; i++)
        {
            if (!ht.Contains(sParts[i].ToLower()))
                ht.Add(sParts[i].ToLower(), sParts[i].ToLower());
        }

        return ht;
    }

#if true
    /// <summary>Read integer setting from config file.</summary>
    /// <param name="sSettingName">string name of setting.</param>
    /// <returns>int configuration setting value.</returns>
    public static int GetIntSetting(string sSettingName)
    {
        int iValue = 0;

        if (ConfigurationManager.AppSettings[sSettingName] != null)
            int.TryParse(ConfigurationManager.AppSettings[sSettingName] as string, out iValue);

        return iValue;
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="sSettingName"></param>
    /// <returns></returns>
    public static bool GetBooleanSetting(string sSettingName)
    {
        bool bValue = false;

        if (ConfigurationManager.AppSettings[sSettingName] != null)
        {
            if (!bool.TryParse(ConfigurationManager.AppSettings[sSettingName], out bValue))
            {
                if (ConfigurationManager.AppSettings[sSettingName].ToLower() == "true")
                    bValue = true;
                else if (ConfigurationManager.AppSettings[sSettingName].ToLower() == "false")
                    bValue = false;
                else if (ConfigurationManager.AppSettings[sSettingName].ToLower() == "yes")
                    bValue = true;
                else if (ConfigurationManager.AppSettings[sSettingName].ToLower() == "no")
                    bValue = false;
                else if (ConfigurationManager.AppSettings[sSettingName] == "1")
                    bValue = true;
                else if (ConfigurationManager.AppSettings[sSettingName] == "0")
                    bValue = false;
            }
        }

        return bValue;
    }
#endif

    /// <summary>Check if file is locked.</summary>
    /// <param name="sFileName">string full path to file.</param>
    /// <returns>bool true if locked or false if not.</returns>
    public static bool IsFileLocked(string sFileName)
    {
        if (File.Exists(sFileName))
        {
            FileInfo fi = new FileInfo(sFileName);

            FileStream stream = null;

            try
            {
                stream = fi.Open(FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            }
            catch (IOException ex)
            {
                System.Diagnostics.Debug.WriteLine(ex.Message);
                System.Diagnostics.Debug.WriteLine(ex.StackTrace);

                return true;
            }
            finally
            {
                if (stream != null)
                    stream.Close();
            }
        }

        return false;
    }

#if true

    /// <summary>
    /// 
    /// </summary>
    /// <param name="sSettingName"></param>
    /// <returns></returns>
    public static string GetSetting(string sSettingName)
    {
        if (ConfigurationManager.AppSettings[sSettingName] != null)
            return ConfigurationManager.AppSettings[sSettingName];
        return string.Empty;
    }

    public static string GetSettingFromDLLConfig(string sSettingName)
    {
        Configuration appConfig = ConfigurationManager.OpenExeConfiguration
                        (
                        System.Reflection.Assembly.GetExecutingAssembly().Location
                        );

        if (appConfig.AppSettings.Settings[sSettingName] != null)
            return appConfig.AppSettings.Settings[sSettingName].Value;

        return string.Empty;
    }

    /// <summary>Read integer setting from DLL config file.</summary>
    /// <param name="sSettingName">string name of setting.</param>
    /// <returns>int configuration setting value.</returns>
    public static int GetIntSettingFromDLLConfig(string sSettingName)
    {
        int iValue = 0;

        Configuration appConfig = ConfigurationManager.OpenExeConfiguration
                        (
                        System.Reflection.Assembly.GetExecutingAssembly().Location
                        );

        if (appConfig.AppSettings.Settings[sSettingName] != null)
            int.TryParse(appConfig.AppSettings.Settings[sSettingName].Value, out iValue);

        return iValue;
    }

    /// <summary>
    /// Read Boolean setting from DLL Config file
    /// </summary>
    /// <param name="sSettingName"></param>
    /// <returns></returns>
    public static bool GetBooleanSettingFromDLLConfig(string sSettingName)
    {
        bool bValue = false;

        Configuration appConfig = ConfigurationManager.OpenExeConfiguration
                        (
                        System.Reflection.Assembly.GetExecutingAssembly().Location
                        );

        if (appConfig.AppSettings.Settings[sSettingName] != null)
        {
            if (!bool.TryParse(appConfig.AppSettings.Settings[sSettingName].Value, out bValue))
            {
                string sValue = appConfig.AppSettings.Settings[sSettingName].Value;

                if (sValue.ToLower() == "true")
                    bValue = true;
                else if (sValue.ToLower() == "false")
                    bValue = false;
                else if (sValue == "yes")
                    bValue = true;
                else if (sValue == "no")
                    bValue = false;
                else if (sValue == "1")
                    bValue = true;
                else if (sValue == "0")
                    bValue = false;
            }
        }

        return bValue;
    }
#endif
}

public sealed class PWSession : IDisposable
{
    public string Datasource { get; set; }
    public bool LoggedInOK { get; set; }
    public string User { get; set; }
    public bool IsAdmin { get; set; }

    private bool DoLogin(string sDatasourceName, string sUserName, string sPassword)
    {
        PWWrapper.aaApi_Initialize(512);

        return (PWWrapper.aaApi_Login(PWWrapper.DataSourceType.Unknown,
            sDatasourceName, sUserName, sPassword, "", false));
    }

    private bool DoAdminLogin(string sDatasourceName, string sUserName, string sPassword)
    {
        PWWrapper.aaApi_Initialize(512);

        return (PWWrapper.aaApi_AdminLogin(PWWrapper.DataSourceType.Unknown,
            sDatasourceName, sUserName, sPassword));
    }

    /// <summary>
    /// Constructor with all parameters passed, use unencrypted password.
    /// Throws exception if login unsuccessful.
    /// Throws exception if datasource name not set.
    /// </summary>
    /// <param name="sDatasourceName"></param>
    /// <param name="sUserName"></param>
    /// <param name="sPassword"></param>
    /// <param name="bLoginAsAdmin"></param>
    public PWSession(string sDatasourceName, string sUserName, string sPassword, bool bLoginAsAdmin)
    {
        if (string.IsNullOrEmpty(sDatasourceName))
        {
            LoggedInOK = false;
            throw new Exception("Datasource name not set");
        }

        if (bLoginAsAdmin)
        {
            if (!DoAdminLogin(sDatasourceName, sUserName, sPassword))
            {
                LoggedInOK = false;

                throw new Exception(string.Format("Error logging in as ADMIN to '{0}' as {1}, Error: {2}",
                    sDatasourceName, sUserName, PWWrapper.aaApi_GetLastErrorId()));
            }
            else
            {
                Datasource = sDatasourceName;
                User = sUserName;
                LoggedInOK = true;
            }
        }
        else
        {
            if (!DoLogin(sDatasourceName, sUserName, sPassword))
            {
                LoggedInOK = false;

                throw new Exception(string.Format("Error logging in to '{0}' as {1}, Error: {2}",
                    sDatasourceName, sUserName, PWWrapper.aaApi_GetLastErrorId()));
            }
            else
            {
                Datasource = sDatasourceName;
                User = sUserName;
                LoggedInOK = true;
            }
        }

        IsAdmin = PWWrapper.aaApi_HasAdminSetup();
    }
    
    /// <summary>
    /// Constructor with all parameters passed, use unencrypted password.
    /// Throws exception if login unsuccessful.
    /// Throws exception if datasource name not set.
    /// </summary>
    /// <param name="sDatasourceName"></param>
    /// <param name="sUserName"></param>
    /// <param name="sPassword"></param>
    public PWSession(string sDatasourceName, string sUserName, string sPassword)
    {
        if (string.IsNullOrEmpty(sDatasourceName))
        {
            LoggedInOK = false;
            throw new Exception("Datasource name not set");
        }

        if (!DoLogin(sDatasourceName, sUserName, sPassword))
        {
            LoggedInOK = false;

            throw new Exception(string.Format("Error logging in to '{0}' as {1}, Error: {2}",
                sDatasourceName, sUserName, PWWrapper.aaApi_GetLastErrorId()));
        }
        else
        {
            Datasource = sDatasourceName;
            User = sUserName;
            LoggedInOK = true;
        }

        IsAdmin = PWWrapper.aaApi_HasAdminSetup();
    }
    /// <summary>
    /// Constructor which just takes datasource name and performs single sign-on login
    /// Throws exception if login unsuccessful.
    /// Throws exception if datasource name not set.
    /// </summary>
    /// <param name="sDatasourceName"></param>
    public PWSession(string sDatasourceName)
    {
        if (string.IsNullOrEmpty(sDatasourceName))
        {
            LoggedInOK = false;
            throw new Exception("Datasource name not set");
        }

        if (!DoLogin(sDatasourceName, string.Empty, string.Empty))
        {
            LoggedInOK = false;

            throw new Exception(string.Format("Error logging in to '{0}' with {1}, Error: {2}",
                sDatasourceName, "SSO", PWWrapper.aaApi_GetLastErrorId()));
        }
        else
        {
            Datasource = sDatasourceName;
            LoggedInOK = true;
        }

        IsAdmin = PWWrapper.aaApi_HasAdminSetup();
    }

    /// <summary>
    /// Constructor which just takes a Token with datasource name and performs IMS login
    /// Throws exception if login unsuccessful.
    /// Throws exception if datasource name not set.
    /// Throws exception if token not set.
    /// </summary>
    /// <param name="sDatasourceName"></param>
    /// <param name="sPossiblyEncodedTokenWithOrWithoutPrefix"></param>
    /// <param name="bDoAdminLogin"></param>
    public PWSession(string sDatasourceName, string sPossiblyEncodedTokenWithOrWithoutPrefix, bool bDoAdminLogin)
    {
        if (string.IsNullOrEmpty(sPossiblyEncodedTokenWithOrWithoutPrefix))
        {
            LoggedInOK = false;
            throw new Exception("Token not set");
        }

        if (string.IsNullOrEmpty(sDatasourceName))
        {
            LoggedInOK = false;
            throw new Exception("Datasource not set");
        }

        string sValidatedToken = BPSUtilities.DecodeAndValidateToken(sPossiblyEncodedTokenWithOrWithoutPrefix);

        if (string.IsNullOrEmpty(sValidatedToken))
        {
            // token was not valid SAML
            if (sPossiblyEncodedTokenWithOrWithoutPrefix.Contains(":"))
            {
                string[] sParts = sPossiblyEncodedTokenWithOrWithoutPrefix.Split(":".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);

                if (sParts.Length == 2)
                {
                    if (bDoAdminLogin)
                    {
                        if (!PWWrapper.aaApi_AdminLogin(PWWrapper.DataSourceType.Unknown, sDatasourceName, sParts[0], sParts[1]))
                        {
                            LoggedInOK = false;

                            throw new Exception(string.Format("Error logging in to '{0}' with {1}, Error: {2}",
                                sDatasourceName, "IMS Token", PWWrapper.aaApi_GetLastErrorId()));
                        }
                    }
                    else
                    {
                        if (!PWWrapper.aaApi_Login(PWWrapper.DataSourceType.Unknown, sDatasourceName, sParts[0], sParts[1], string.Empty, true))
                        {
                            LoggedInOK = false;

                            throw new Exception(string.Format("Error logging in to '{0}' with {1}, Error: {2}",
                                sDatasourceName, "IMS Token", PWWrapper.aaApi_GetLastErrorId()));
                        }
                    }

                    Datasource = sDatasourceName;
                    LoggedInOK = true;
                }
            }
            else
            {
                LoggedInOK = false;
                throw new Exception("Invalid token");
            }
        }
        else if (!PWWrapper.aaApi_LoginWithSecurityToken(sDatasourceName, sValidatedToken, bDoAdminLogin, null, null))
        {
            LoggedInOK = false;

            throw new Exception(string.Format("Error logging in to '{0}' with {1}, Error: {2}",
                sDatasourceName, "IMS Token", PWWrapper.aaApi_GetLastErrorId()));
        }
        else
        {
            Datasource = sDatasourceName;
            LoggedInOK = true;
        }

        IsAdmin = PWWrapper.aaApi_HasAdminSetup();
    }

    /// <summary>
    /// Constructor which uses values from configuration file and decrypts password string (use EncryptPasswordForDSList.exe).
    /// Expects PWDatasSourceName, PWUser and PWPassword in the config file.
    /// Will use single sign on if password not set.
    /// Throws exception if login unsuccessful.
    /// Throws exception if decryption unsuccessful.
    /// Throws exception if datasource name not set.
    /// </summary>
    public PWSession()
    {
#if true
        string sDatasourceName = BPSUtilities.GetSetting("PWDataSourceName");
        string sUserName = BPSUtilities.GetSetting("PWUser");
        string sPassword = BPSUtilities.GetSetting("PWPassword");
        bool bLoginAsAdmin = BPSUtilities.GetBooleanSetting("LoginAsAdmin");
#else
        string sDatasourceName = System.Environment.GetEnvironmentVariable("PWDataSourceName");
        string sUserName = System.Environment.GetEnvironmentVariable("PWUser");
        string sPassword = System.Environment.GetEnvironmentVariable("PWPassword");
        bool bLoginAsAdmin = false;
        bool.TryParse(System.Environment.GetEnvironmentVariable("LoginAsAdmin"), out bLoginAsAdmin);
#endif
        if (string.IsNullOrEmpty(sDatasourceName))
        {
            LoggedInOK = false;
            throw new Exception("Datasource name not set");
        }

        string sDecryptedPassword = string.Empty;

        if (!string.IsNullOrEmpty(sPassword))
        {
            try
            {
                sDecryptedPassword = CryptoProvider.GetDecryptedPassword(sDatasourceName, sUserName, sPassword);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(string.Format("Error: {0}\n{1}", ex.Message, ex.StackTrace));
                throw new Exception("Error decrypting password. Are all app settings defined?");
            }
        }
        else
        {
            sUserName = string.Empty;
        }

        if (bLoginAsAdmin)
        {
            if (!DoAdminLogin(sDatasourceName, sUserName, sDecryptedPassword))
            {
                LoggedInOK = false;

                throw new Exception(string.Format("Error logging in to '{0}' as {1}, Error: {2}",
                    sDatasourceName, sUserName, PWWrapper.aaApi_GetLastErrorId()));
            }
            else
            {
                Datasource = sDatasourceName;
                User = sUserName;
                LoggedInOK = true;
            }
        }
        else
        {
            if (!DoLogin(sDatasourceName, sUserName, sDecryptedPassword))
            {
                LoggedInOK = false;

                throw new Exception(string.Format("Error logging in to '{0}' as {1}, Error: {2}",
                    sDatasourceName, sUserName, PWWrapper.aaApi_GetLastErrorId()));
            }
            else
            {
                Datasource = sDatasourceName;
                User = sUserName;
                LoggedInOK = true;
            }
        }

        IsAdmin = PWWrapper.aaApi_HasAdminSetup();
    }

#region IDisposable Members

    public void Dispose()
    {
        PWWrapper.aaApi_LogoutByHandle(PWWrapper.aaApi_GetActiveDatasource());
        GC.Collect();
    }

#endregion
}

public class CryptoProvider
{
    private static string EncryptDecrypt(string sDatasource, string sUserName, string sPassword, bool bEncrypt)
    {
        if (!string.IsNullOrEmpty(sDatasource) && !string.IsNullOrEmpty(sUserName) && !string.IsNullOrEmpty(sPassword))
        {
            string sKey = string.Format("{0}_{1}",
                sDatasource.ToLower(), sUserName.ToLower());

            try
            {
                if (bEncrypt)
                    return CryptoProvider.EncryptData(sPassword, sKey);
                else
                    return CryptoProvider.DecryptData(sPassword, sKey);
            }
            catch (Exception ex)
            {
                Console.WriteLine(string.Format("Error: {0}", ex.Message));
                Console.WriteLine(string.Format("Details: {0}", ex.StackTrace));
            }
        }

        return string.Empty;
    }

    public static string GetEncryptedPassword(string sDataSource, string sUserName, string sPassword)
    {
        return EncryptDecrypt(sDataSource, sUserName, sPassword, true);
    }

    public static string GetDecryptedPassword(string sDataSource, string sUserName, string sPassword)
    {
        return EncryptDecrypt(sDataSource, sUserName, sPassword, false);
    }

    private static byte[] TruncateHash(string sKey, int iLength)
    {
        SHA1CryptoServiceProvider sha1 = new SHA1CryptoServiceProvider();
        byte[] bytesKey = System.Text.Encoding.Unicode.GetBytes(sKey);
        byte[] bytesHash = sha1.ComputeHash(bytesKey);

        byte[] bytesRetVal = new byte[iLength];

        for (int i = 0; i < bytesRetVal.Length; i++)
        {
            if (i < bytesHash.Length - 1)
            {
                bytesRetVal[i] = bytesHash[i];
            }
            else
            {
                bytesRetVal[i] = 0;
            }
        }

        return bytesRetVal;
    }

    public static string EncryptData(string sData, string sKey)
    {
        byte[] bytesToEncrypt = System.Text.Encoding.Unicode.GetBytes(sData);

        TripleDESCryptoServiceProvider tripleDESAlg = new TripleDESCryptoServiceProvider();
        tripleDESAlg.Key = TruncateHash(sKey, (int)(tripleDESAlg.KeySize / 8));
        tripleDESAlg.IV = TruncateHash("", (int)(tripleDESAlg.BlockSize / 8));

        System.IO.MemoryStream ms = new System.IO.MemoryStream();
        CryptoStream encStream = new CryptoStream(ms, tripleDESAlg.CreateEncryptor(),
            CryptoStreamMode.Write);
        encStream.Write(bytesToEncrypt, 0, bytesToEncrypt.Length);
        encStream.FlushFinalBlock();
        return Convert.ToBase64String(ms.ToArray());
    }

    public static string DecryptData(string sEncrpytedData, string sKey)
    {
        byte[] bytesToDecrypt = Convert.FromBase64String(sEncrpytedData);

        TripleDESCryptoServiceProvider tripleDESAlg = new TripleDESCryptoServiceProvider();
        tripleDESAlg.Key = TruncateHash(sKey, (int)(tripleDESAlg.KeySize / 8));
        tripleDESAlg.IV = TruncateHash("", (int)(tripleDESAlg.BlockSize / 8));

        System.IO.MemoryStream ms = new System.IO.MemoryStream();
        CryptoStream decStream = new CryptoStream(ms, tripleDESAlg.CreateDecryptor(),
            CryptoStreamMode.Write);
        decStream.Write(bytesToDecrypt, 0, bytesToDecrypt.Length);
        decStream.FlushFinalBlock();
        return System.Text.Encoding.Unicode.GetString(ms.ToArray());
    }
}

public class SMTPMailSender
{
    const int IDOK = 1;

    // to ensure this can send mail from localhost
    // make sure that Default SMTP Virtual Server Properties
    // are set such that Access/Relay restrictions 
    // are set to "All except the list below"

    public static bool SendHTMLMail(string sMailSubject,
        string sMailBody, string sToEmails, string sFromEMail, string sFileAttachment,
        string sSmtpServer, int iSmtpPort, bool bEnableSSL)
    {
        if (string.IsNullOrEmpty(sSmtpServer) || iSmtpPort < 1 ||
                string.IsNullOrEmpty(sFromEMail) || string.IsNullOrEmpty(sToEmails))
        {
            BPSUtilities.WriteLog("Server: {0}, Port: {1}, From: {2}, To: {3}",
                sSmtpServer, iSmtpPort, // sUserName, 
                                        // sPassword, 
                sFromEMail, sToEmails);
            BPSUtilities.WriteLog("Unable to get email sending parameters");
            return false;
        }

        //Build The MSG
        System.Net.Mail.MailMessage msg = new System.Net.Mail.MailMessage();

        try
        {
            string[] sToEmailArray = sToEmails.Split(";".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);

            foreach (string sToEmail in sToEmailArray)
            {
                msg.To.Add(sToEmail);
            }

            if (msg.To.Count == 0)
            {
                BPSUtilities.WriteLog("No recipients defined");
                return false;
            }

            msg.From = new System.Net.Mail.MailAddress(sFromEMail);
            msg.Subject = sMailSubject;
            msg.SubjectEncoding = System.Text.Encoding.UTF8;
            // msg.Body = sBody;
            msg.BodyEncoding = System.Text.Encoding.ASCII;
            msg.IsBodyHtml = true;
            msg.Body = sMailBody;

            msg.Priority = System.Net.Mail.MailPriority.Normal;

            if (!string.IsNullOrEmpty(sFileAttachment))
            {
                if (File.Exists(sFileAttachment))
                {
                    System.Net.Mail.Attachment fileAttachment =
                        new System.Net.Mail.Attachment(sFileAttachment,
                        System.Net.Mime.MediaTypeNames.Application.Octet);

                    msg.Attachments.Add(fileAttachment);
                }
            }

            //Add the Credentials
            System.Net.Mail.SmtpClient client = new System.Net.Mail.SmtpClient();
            // System.Net.NetworkCredential netCred = System.Net.CredentialCache.DefaultNetworkCredentials; // new System.Net.NetworkCredential((sUserName, sPassword);
            client.Port = iSmtpPort; //25; //or use 587            
            client.Host = sSmtpServer; //"localhost";
            client.UseDefaultCredentials = true;
            // client.Credentials = netCred;
            client.EnableSsl = bEnableSSL;

            client.DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.Network; //use Network for SmtpDeliveryMethod

            object userState = msg;
            try
            {
                client.Send(msg);
                // client.SendAsync(msg, userState);
                BPSUtilities.WriteLog("Sent Mail to '{0}' OK", sToEmails);
            }
            catch (System.Net.Mail.SmtpException ex)
            {
                BPSUtilities.WriteLog(
                    String.Format("{0}", ex.Message));
                BPSUtilities.WriteLog(
                    String.Format("{0}", ex.StackTrace));

                Exception ex1 = (Exception)ex;

                while (ex1.InnerException != null)
                {
                    BPSUtilities.WriteLog(
                            "--------------------------------");
                    BPSUtilities.WriteLog(
                            "The following InnerException reported: " + ex1.InnerException.ToString());
                    ex1 = ex1.InnerException;
                }

                return false;
            }
        }
        catch (Exception ex)
        {
            BPSUtilities.WriteLog(
                String.Format("{0}", ex.Message));
            BPSUtilities.WriteLog(
                String.Format("{0}", ex.StackTrace));

            while (ex.InnerException != null)
            {
                BPSUtilities.WriteLog(
                        "--------------------------------");
                BPSUtilities.WriteLog(
                        "The following InnerException reported: " + ex.InnerException.ToString());
                ex = ex.InnerException;
            }

            return false;
        }

        msg.Dispose();

        // BPSUtilities.WriteLog("Sent mail OK");
        return true;
    }//SendMail

    public static bool SendHTMLMail(string sMailSubject,
        string sMailBody, string sToEmails, string sFromEMail, string sFileAttachment,
        string sSmtpServer, int iSmtpPort, bool bEnableSSL, string sUserName, string sEncryptedPassword)
    {
        string sPassword = string.Empty;

        if (!string.IsNullOrEmpty(sUserName) && !string.IsNullOrEmpty(sEncryptedPassword))
        {
            try
            {
                sPassword = CryptoProvider.GetDecryptedPassword(sSmtpServer, sUserName, sEncryptedPassword);
            }
            catch (Exception ex)
            {
                BPSUtilities.WriteLog("Error: {0}\n{1}", ex.Message, ex.StackTrace);
            }
        }

        if (string.IsNullOrEmpty(sSmtpServer) || iSmtpPort < 1 ||
                string.IsNullOrEmpty(sFromEMail) || string.IsNullOrEmpty(sToEmails))
        {
            BPSUtilities.WriteLog("Server: {0}, Port: {1}, From: {2}, To: {3}",
                sSmtpServer, iSmtpPort, // sUserName, 
                                        // sPassword, 
                sFromEMail, sToEmails);
            BPSUtilities.WriteLog("Unable to get email sending parameters");
            return false;
        }

        //Build The MSG
        System.Net.Mail.MailMessage msg = new System.Net.Mail.MailMessage();

        try
        {
            // msg.To.Add(m_sPlantManagerEmail);

            string[] sToEmailArray = sToEmails.Split(";".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);

            foreach (string sToEmail in sToEmailArray)
            {
                msg.To.Add(sToEmail);
            }

            if (msg.To.Count == 0)
            {
                BPSUtilities.WriteLog("No recipients defined");
                return false;
            }

            msg.From = new System.Net.Mail.MailAddress(sFromEMail);
            msg.Subject = sMailSubject;
            msg.SubjectEncoding = System.Text.Encoding.UTF8;
            // msg.Body = sBody;
            msg.BodyEncoding = System.Text.Encoding.ASCII;
            msg.IsBodyHtml = true;
            msg.Body = sMailBody;

            msg.Priority = System.Net.Mail.MailPriority.Normal;

            if (!string.IsNullOrEmpty(sFileAttachment))
            {
                if (File.Exists(sFileAttachment))
                {
                    System.Net.Mail.Attachment fileAttachment =
                        new System.Net.Mail.Attachment(sFileAttachment,
                        System.Net.Mime.MediaTypeNames.Application.Octet);

                    msg.Attachments.Add(fileAttachment);
                }
            }

            //Add the Creddentials
            System.Net.Mail.SmtpClient client = new System.Net.Mail.SmtpClient();

            if (!string.IsNullOrEmpty(sUserName) && !string.IsNullOrEmpty(sPassword))
            {
                System.Net.NetworkCredential netCred = new System.Net.NetworkCredential(sUserName, sPassword);
                client.Credentials = netCred;
            }
            else
            {
                client.UseDefaultCredentials = true;
            }

            client.Port = iSmtpPort; //25; //or use 587            
            client.Host = sSmtpServer; //"localhost";
            client.EnableSsl = bEnableSSL;

            client.DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.Network; //use Network for SmtpDeliveryMethod

            object userState = msg;

            try
            {
                client.Send(msg);
                // client.SendAsync(msg, userState);
                BPSUtilities.WriteLog("Sent Mail to '{0}' OK", sToEmails);
            }
            catch (System.Net.Mail.SmtpException ex)
            {
                BPSUtilities.WriteLog(
                    String.Format("{0}", ex.Message));
                BPSUtilities.WriteLog(
                    String.Format("{0}", ex.StackTrace));

                Exception ex1 = (Exception)ex;

                while (ex1.InnerException != null)
                {
                    BPSUtilities.WriteLog(
                            "--------------------------------");
                    BPSUtilities.WriteLog(
                            "The following InnerException reported: " + ex1.InnerException.ToString());
                    ex1 = ex1.InnerException;
                }

                return false;
            }
        }
        catch (Exception ex)
        {
            BPSUtilities.WriteLog(
                String.Format("{0}", ex.Message));
            BPSUtilities.WriteLog(
                String.Format("{0}", ex.StackTrace));

            while (ex.InnerException != null)
            {
                BPSUtilities.WriteLog(
                        "--------------------------------");
                BPSUtilities.WriteLog(
                        "The following InnerException reported: " + ex.InnerException.ToString());
                ex = ex.InnerException;
            }

            return false;
        }

        msg.Dispose();

        // BPSUtilities.WriteLog("Sent mail OK");
        return true;
    }//SendMail
}
public class ThreadManager
{

    // Usage:

    // Thread function
    // public static void DoLoginAndCreate(object oDsUserPwd) { }

    // functioning adding to ThreadManager queue
    // void SomeFunction () {   
    // ParameterizedThreadStart paramThd = new ParameterizedThreadStart(DoLoginAndCreate);
    // Thread thd = new Thread(paramThd);
    // thd.IsBackground = true;
    // ArrayList args = new ArrayList();
    // args.Add(someArgument);
    // args.Add(anotherArgument);
    // args.Add(sUnEncryptedPassword);
    // args.Add(bIsAdmin);
    // args.Add(Properties.Settings.Default.FileSizeK);

    // ThreadManager.ManageThreads(thd, args); }

    private static List<Thread> g_listOfThreads = new List<Thread>();

    public static bool Shutdown = false;

    public static bool Paused = false;

    public static int ThreadCount { get { return g_listOfThreads.Count; } }

    public static int ActiveThreadCount()
    {
        try
        {
            int iDeadCount = 0;

            lock (g_listOfThreads)
            {
                while (iDeadCount < g_listOfThreads.Count)
                {
                    iDeadCount = 0;

                    foreach (Thread th in g_listOfThreads)
                    {
                        if (!th.IsAlive)
                        {
                            iDeadCount++;
                        }
                    }
                }

                return g_listOfThreads.Count - iDeadCount;
            }
        }
        catch (Exception ex)
        {
            BPSUtilities.WriteLog("Error: {0}\n{1}", ex.Message, ex.StackTrace);
        }

        return 0;
    }

    public static void ShutdownThreads()
    {
        // return;

        Shutdown = true;

        lock (g_listOfThreads)
        {
            BPSUtilities.WriteLog("Shutting down threads");

            try
            {
                int iDeadCount = 0;

                while (iDeadCount < g_listOfThreads.Count)
                {
                    iDeadCount = 0;

                    foreach (Thread th in g_listOfThreads)
                    {
                        if (!th.IsAlive)
                        {
                            iDeadCount++;
                        }
                        else
                        {
                            try
                            {
                                th.Abort();
                            }
                            catch (Exception ex)
                            {
                                System.Diagnostics.Debug.WriteLine(string.Format("Error: {0}\n{1}", ex.Message, ex.StackTrace));
                            }
                        }
                    }

                    Thread.Sleep(10000);
                }

                g_listOfThreads.Clear();

                BPSUtilities.WriteLog("Thread shutdown complete");
            }
            catch (Exception ex)
            {
                BPSUtilities.WriteLog("Error: {0}\n{1}", ex.Message, ex.StackTrace);
            }
        } // lock (g_listOfThreads)
    }

    public static int ThreadLimit = 20;

    public static bool ManageThreads(Thread threadToStart, object alArgs)
    {
        bool bRetVal = false;

        if (!Shutdown)
        {
            lock (g_listOfThreads)
            {
                int iThreadLimit = ThreadLimit;

                if (g_listOfThreads.Count >= iThreadLimit)
                {
                    BPSUtilities.WriteLog("Limit reached. {0} managed threads running.", g_listOfThreads.Count);

                    try
                    {
                        int iDeadCount = 0;

                        while (iDeadCount < g_listOfThreads.Count)
                        {
                            iDeadCount = 0;

                            foreach (Thread th in g_listOfThreads)
                            {
                                if (!th.IsAlive)
                                {
                                    iDeadCount++;
                                }
                            }

                            Thread.Sleep(10000);
                        }

                        g_listOfThreads.Clear();

                        BPSUtilities.WriteLog("Thread list cleared. {0} managed threads running.", g_listOfThreads.Count);

                        // Shutdown = ProjectWiseCoveoConnector.ReadStopRegistryKey();
                    }
                    catch (Exception ex)
                    {
                        BPSUtilities.WriteLog("Error: {0}\n{1}", ex.Message, ex.StackTrace);
                    }
                }

                // if (!Paused)
                if (true)
                {
                    try
                    {
                        // Thread projectThread = new Thread(new ParameterizedThreadStart(CreateProjectProcess));

                        threadToStart.Start(alArgs);

                        g_listOfThreads.Add(threadToStart);

                        // BPSUtilities.WriteLog("{0} managed threads running.", g_listOfThreads.Count);

                        bRetVal = true;
                    }
                    catch (Exception ex)
                    {
                        BPSUtilities.WriteLog("Error: {0}\n{1}", ex.Message, ex.StackTrace);
                    }
                }
            } // lock (g_listOfThreads)
        }
        else
        {
            ShutdownThreads();
        }

        return bRetVal;
    }
}

public class XMLSpreadsheetDatasetTools
{
    private static string getColumnType(DataColumn dc)
    {
        string columnType = "String";
        switch (dc.DataType.ToString())
        {
            case "System.UInt64":
            case "System.UInt32":
            case "System.Int64":
            case "System.Double":
            case "System.Int32":
                columnType = "Number";
                break;
            //case "System.DateTime":
            //    columnType = "DateTime";
            //    break;
            default:
                columnType = "String";
                break;
        }
        return columnType;
    }

    private static string CleanUpString(string sInString)
    {
        string sOutString = (string)sInString.Clone();
        string sWorkString = "";

        if (!sOutString.Contains("&amp;"))
            sWorkString = sOutString.Replace("&", "&amp;");

        sOutString = sWorkString.Replace(">", "&gt;");
        sWorkString = sOutString.Replace("<", "&lt;");
        sOutString = sWorkString.Replace("\"", "&quot;");
        sOutString = sOutString.Replace("'", "&apos;");

        return sOutString;
    }


    private const string XML_HEADER = "<?xml version=\"1.0\"?><?mso-application progid=\"Excel.Sheet\"?><Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\" xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\" xmlns:html=\"http://www.w3.org/TR/REC-html40\">";
    private const string XML_HEADER1 = "<DocumentProperties xmlns=\"urn:schemas-microsoft-com:office:office\"><LastAuthor></LastAuthor><Created></Created><Version>11.6568</Version></DocumentProperties><ExcelWorkbook xmlns=\"urn:schemas-microsoft-com:office:excel\"><WindowHeight>12525</WindowHeight><WindowWidth>18075</WindowWidth><WindowTopX>0</WindowTopX>";
    private const string XML_HEADER2 = "<WindowTopY>15</WindowTopY><ProtectStructure>False</ProtectStructure><ProtectWindows>False</ProtectWindows></ExcelWorkbook><Styles><Style ss:ID=\"Default\" ss:Name=\"Normal\"><Alignment ss:Vertical=\"Bottom\"/><Borders/><Font/><Interior/><NumberFormat/><Protection/></Style><Style ss:ID=\"s22\"><Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Bottom\"/><Font x:Family=\"Swiss\" ss:Bold=\"1\"/></Style><Style ss:ID=\"s23\"><Alignment ss:Horizontal=\"Left\" ss:Vertical=\"Bottom\"/><Font x:Family=\"Swiss\" ss:Bold=\"1\"/></Style></Styles>";
    private const string XML_WORKSHEETHEADER = "<Worksheet ss:Name=\"{0}\"><Table x:FullColumns=\"1\" x:FullRows=\"1\">\n";
    private const string XML_WORKSHEETFOOTER = "</Table><WorksheetOptions xmlns=\"urn:schemas-microsoft-com:office:excel\"><Selected/><ProtectObjects>False</ProtectObjects><ProtectScenarios>False</ProtectScenarios></WorksheetOptions></Worksheet>";
    private const string XML_WORKBOOKFOOTER = "</Workbook>";
    private const string XML_ROW_FORMAT = "<Row><Cell ss:StyleID=\"s22\"><Data ss:Type=\"{2}\">{0}</Data></Cell><Cell ss:StyleID=\"s22\"><Data ss:Type=\"String\">{1}</Data></Cell></Row>";

    public static bool WriteDatasetToXMLSpreadsheet(DataSet ds,
        string sXMLFile)
    {
        using (System.IO.StreamWriter sw = new System.IO.StreamWriter(sXMLFile, false, Encoding.Unicode))
        {
            sw.Write(XML_HEADER);
            sw.Write(XML_HEADER1);
            sw.Write(XML_HEADER2);

            foreach (DataTable dt in ds.Tables)
            {
                sw.Write(string.Format(XML_WORKSHEETHEADER, dt.TableName));

                sw.WriteLine("<Row>");

                System.Collections.ArrayList alTypes = new System.Collections.ArrayList();

                foreach (DataColumn dc in dt.Columns)
                {
                    sw.WriteLine("<Cell ss:StyleID=\"s22\"><Data ss:Type=\"String\">{0}</Data></Cell>",
                        CleanUpString(dc.ColumnName));
                    alTypes.Add(getColumnType(dc));
                }

                sw.WriteLine("</Row>");

                foreach (DataRow dr in dt.Rows)
                {
                    sw.WriteLine("<Row>");

                    for (int i = 0; i < dr.ItemArray.Length; i++)
                    {
                        // 2008-03-02T00:00:00.000
                        sw.WriteLine("<Cell><Data ss:Type=\"{0}\">{1}</Data></Cell>",
                            alTypes[i], CleanUpString(dr[i].ToString()));
                    }
                    sw.WriteLine("</Row>");
                }

                sw.Write(XML_WORKSHEETFOOTER);
            }

            sw.Write(XML_WORKBOOKFOOTER);
        } // using (StreamWriter sw = new StreamWriter(dlg.FileName))

        return true;
    }

    private const string XML_HEADER_URL = "<?xml version=\"1.0\"?><?mso-application progid=\"Excel.Sheet\"?><Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\" xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\" xmlns:html=\"http://www.w3.org/TR/REC-html40\">";
    private const string XML_HEADER1_URL = "<DocumentProperties xmlns=\"urn:schemas-microsoft-com:office:office\"><LastAuthor></LastAuthor><Created></Created><Version>11.6568</Version></DocumentProperties><ExcelWorkbook xmlns=\"urn:schemas-microsoft-com:office:excel\"><WindowHeight>12525</WindowHeight><WindowWidth>18075</WindowWidth><WindowTopX>0</WindowTopX>";
    private const string XML_HEADER2_URL = "<WindowTopY>15</WindowTopY><ProtectStructure>False</ProtectStructure><ProtectWindows>False</ProtectWindows></ExcelWorkbook><Styles><Style ss:ID=\"Default\" ss:Name=\"Normal\"><Alignment ss:Vertical=\"Bottom\"/><Borders/><Font/><Interior/><NumberFormat/><Protection/></Style><Style ss:ID=\"s22\"><Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Bottom\"/><Font x:Family=\"Swiss\" ss:Bold=\"1\"/></Style><Style ss:ID=\"s23\"><Alignment ss:Horizontal=\"Left\" ss:Vertical=\"Bottom\"/><Font x:Family=\"Swiss\" ss:Bold=\"1\"/></Style>";
    private const string XML_HEADER3_URL = "<Style ss:ID=\"s63\" ss:Name=\"Hyperlink\"><Font ss:FontName=\"Arial\" ss:Color=\"#0000FF\" ss:Underline=\"Single\"/></Style><Style ss:ID=\"s62\"><Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Bottom\"/><Font ss:FontName=\"Arial\" x:Family=\"Swiss\" ss:Bold=\"1\"/></Style><Style ss:ID=\"s64\" ss:Parent=\"s63\"><Alignment ss:Vertical=\"Bottom\"/><Protection/></Style></Styles>";
    private const string XML_WORKSHEETHEADER_URL = "<Worksheet ss:Name=\"{0}\"><Table x:FullColumns=\"1\" x:FullRows=\"1\">\n";
    private const string XML_WORKSHEETFOOTER_URL = "</Table><WorksheetOptions xmlns=\"urn:schemas-microsoft-com:office:excel\"><Selected/><ProtectObjects>False</ProtectObjects><ProtectScenarios>False</ProtectScenarios></WorksheetOptions></Worksheet>";
    private const string XML_WORKBOOKFOOTER_URL = "</Workbook>";
    private const string XML_ROW_FORMAT_URL = "<Row><Cell ss:StyleID=\"s22\"><Data ss:Type=\"{2}\">{0}</Data></Cell><Cell ss:StyleID=\"s22\"><Data ss:Type=\"String\">{1}</Data></Cell></Row>";

    public static bool WriteDatasetToXMLSpreadsheetWithURLs2(DataSet ds,
        string sXMLFile)
    {
        try
        {
            using (System.IO.StreamWriter sw = new System.IO.StreamWriter(sXMLFile, false, Encoding.Unicode))
            {
                sw.Write(XML_HEADER_URL);
                sw.Write(XML_HEADER1_URL);
                sw.Write(XML_HEADER2_URL);
                sw.Write(XML_HEADER3_URL);

                foreach (DataTable dt in ds.Tables)
                {

                    try
                    {
                        sw.Write(string.Format(XML_WORKSHEETHEADER_URL, dt.TableName));

                        sw.WriteLine("<Row>");

                        System.Collections.ArrayList alTypes = new System.Collections.ArrayList();

                        foreach (DataColumn dc in dt.Columns)
                        {
                            sw.WriteLine("<Cell ss:StyleID=\"s22\"><Data ss:Type=\"String\">{0}</Data></Cell>",
                                dc.ColumnName);
                            alTypes.Add(getColumnType(dc));
                        }

                        sw.WriteLine("</Row>");

                        foreach (DataRow dr in dt.Rows)
                        {
                            sw.WriteLine("<Row>");

                            for (int i = 0; i < dr.ItemArray.Length; i++)
                            {
                                try
                                {
                                    // <Cell ss:HRef="pw:\\BRUMD27169REM:PWv8i\Documents\MicroStation J\detail.dgn"><Data
                                    // ss:Type="String">detail.dgn</Data></Cell>
                                    if (dr[i].ToString().StartsWith("pw://") || dr[i].ToString().StartsWith("pw:\\\\") ||
                                       dr[i].ToString().StartsWith("http://") || dr[i].ToString().StartsWith("http:\\\\") ||
                                       dr[i].ToString().StartsWith("file://") || dr[i].ToString().StartsWith("file:\\\\") ||
                                       dr[i].ToString().StartsWith("ftp://") || dr[i].ToString().StartsWith("ftp:\\\\"))
                                    {
                                        if (dr[i].ToString().Contains("|"))
                                        {
                                            string[] sParts = dr[i].ToString().Split("|".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);

                                            if (sParts.Length == 2)
                                            {
                                                sw.WriteLine("<Cell ss:StyleID=\"s64\" ss:HRef=\"{1}\"><Data ss:Type=\"{0}\">{2}</Data></Cell>",
                                                    alTypes[i], sParts[0], sParts[1]);
                                            }
                                            else
                                            {
                                                sw.WriteLine("<Cell ss:StyleID=\"s64\" ss:HRef=\"{1}\"><Data ss:Type=\"{0}\">{1}</Data></Cell>",
                                                    alTypes[i], dr[i].ToString());
                                            }
                                        }
                                        else
                                        {
                                            sw.WriteLine("<Cell ss:StyleID=\"s64\" ss:HRef=\"{1}\"><Data ss:Type=\"{0}\">{1}</Data></Cell>",
                                                alTypes[i], dr[i].ToString());
                                        }
                                    }
                                    else
                                    {
                                        // 2008-03-02T00:00:00.000
                                        sw.WriteLine("<Cell><Data ss:Type=\"{0}\">{1}</Data></Cell>",
                                            alTypes[i], dr[i].ToString());
                                    }
                                }
                                catch (Exception ex)
                                {
                                    BPSUtilities.WriteLog(ex.Message);
                                    BPSUtilities.WriteLog(ex.StackTrace);
                                }
                            }
                            sw.WriteLine("</Row>");
                        }

                        sw.Write(XML_WORKSHEETFOOTER_URL);

                    }
                    catch (Exception ex)
                    {
                        BPSUtilities.WriteLog(ex.Message);
                        BPSUtilities.WriteLog(ex.StackTrace);
                    }
                }

                sw.Write(XML_WORKBOOKFOOTER_URL);
            } // using (StreamWriter sw = new StreamWriter(dlg.FileName))
        }
        catch (Exception ex)
        {
            BPSUtilities.WriteLog(ex.Message);
            BPSUtilities.WriteLog(ex.StackTrace);
        }

        return true;
    }

    public static bool WriteDatasetToXMLSpreadsheetWithURLs(DataSet ds,
        string sXMLFile)
    {
        using (System.IO.StreamWriter sw = new System.IO.StreamWriter(sXMLFile, false, Encoding.Unicode))
        {
            sw.Write(XML_HEADER_URL);
            sw.Write(XML_HEADER1_URL);
            sw.Write(XML_HEADER2_URL);
            sw.Write(XML_HEADER3_URL);

            foreach (DataTable dt in ds.Tables)
            {
                sw.Write(string.Format(XML_WORKSHEETHEADER_URL, dt.TableName));

                sw.WriteLine("<Row>");

                System.Collections.ArrayList alTypes = new System.Collections.ArrayList();

                foreach (DataColumn dc in dt.Columns)
                {
                    sw.WriteLine("<Cell ss:StyleID=\"s22\"><Data ss:Type=\"String\">{0}</Data></Cell>",
                        dc.ColumnName);
                    alTypes.Add(getColumnType(dc));
                }

                sw.WriteLine("</Row>");

                foreach (DataRow dr in dt.Rows)
                {
                    sw.WriteLine("<Row>");

                    for (int i = 0; i < dr.ItemArray.Length; i++)
                    {
                        try
                        {
                            // <Cell ss:HRef="pw:\\BRUMD27169REM:PWv8i\Documents\MicroStation J\detail.dgn"><Data
                            // ss:Type="String">detail.dgn</Data></Cell>

                            if (dr[i].ToString().StartsWith("pw://") || dr[i].ToString().StartsWith("pw:\\\\"))
                            {
                                sw.WriteLine("<Cell ss:StyleID=\"s64\" ss:HRef=\"{1}\"><Data ss:Type=\"{0}\">{1}</Data></Cell>",
                                    alTypes[i], dr[i].ToString());
                            }
                            else if (dr[i].ToString().StartsWith("http://") || dr[i].ToString().StartsWith("http:\\\\"))
                            {
                                sw.WriteLine("<Cell ss:StyleID=\"s64\" ss:HRef=\"{1}\"><Data ss:Type=\"{0}\">{1}</Data></Cell>",
                                    alTypes[i], dr[i].ToString());
                            }
                            else if (dr[i].ToString().StartsWith("file://") || dr[i].ToString().StartsWith("file:\\\\"))
                            {
                                sw.WriteLine("<Cell ss:StyleID=\"s64\" ss:HRef=\"{1}\"><Data ss:Type=\"{0}\">{1}</Data></Cell>",
                                    alTypes[i], dr[i].ToString());
                            }
                            else if (dr[i].ToString().StartsWith("ftp://") || dr[i].ToString().StartsWith("ftp:\\\\"))
                            {
                                sw.WriteLine("<Cell ss:StyleID=\"s64\" ss:HRef=\"{1}\"><Data ss:Type=\"{0}\">{1}</Data></Cell>",
                                    alTypes[i], dr[i].ToString());
                            }
                            else
                            {
                                // 2008-03-02T00:00:00.000
                                sw.WriteLine("<Cell><Data ss:Type=\"{0}\">{1}</Data></Cell>",
                                    alTypes[i], dr[i].ToString());
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine(ex.Message);
                        }
                    }
                    sw.WriteLine("</Row>");
                }

                sw.Write(XML_WORKSHEETFOOTER_URL);
            }

            sw.Write(XML_WORKBOOKFOOTER_URL);
        } // using (StreamWriter sw = new StreamWriter(dlg.FileName))

        return true;
    }

    public static System.Data.DataSet ReadDataSetFromXMLSpreadsheet(string sXmlPath)
    {
        System.Data.DataSet ds = ReadXmlFileColumnNames(sXmlPath);

        ReadDataSet(sXmlPath, ref ds);

        return ds;
    }

    private static void ReadDataSet(string sXmlPath, ref System.Data.DataSet ds)
    {
        int iRowCount = 0;

        if (System.IO.File.Exists(sXmlPath))
        {
            try
            {
                XmlDocument xmlDoc = new XmlDocument();

                xmlDoc.Load(sXmlPath);

                int iWorksheetIndex = -1;

                // worksheets
                foreach (XmlNode xmlNode in xmlDoc.DocumentElement.ChildNodes)
                {
                    if (xmlNode.Name == "Worksheet")
                    {
                        iWorksheetIndex++;

                        // tables
                        foreach (XmlNode xmlNode2 in xmlNode.ChildNodes)
                        {
                            if (xmlNode2.Name == "Table")
                            {
                                iRowCount = 0;

                                // rows
                                foreach (XmlNode xmlNode3 in xmlNode2.ChildNodes)
                                {
                                    if (xmlNode3.Name == "Row")
                                    {
                                        // skip column names
                                        if (iRowCount == 0)
                                        {
                                            iRowCount++;
                                            continue;
                                        }

                                        // if cells in row
                                        if (xmlNode3.HasChildNodes)
                                        {
                                            int iColIndex = 0;

                                            if (iWorksheetIndex < ds.Tables.Count)
                                            {
                                                System.Data.DataRow dr = ds.Tables[iWorksheetIndex].NewRow();

                                                foreach (XmlNode cellNode in xmlNode3.ChildNodes)
                                                {
                                                    if (cellNode.Name == "Cell")
                                                    {
                                                        string sValue = cellNode.InnerText.Replace("\r", " ");

                                                        if (null != cellNode.Attributes["ss:Index"])
                                                        {
                                                            // this means some previous columns are blank
                                                            int iIndex = int.Parse(cellNode.Attributes["ss:Index"].Value);
                                                            iColIndex = iIndex - 1;
                                                        }

                                                        if (iColIndex < ds.Tables[iWorksheetIndex].Columns.Count)
                                                        {
                                                            dr[iColIndex++] = sValue.Replace("\n", " ");
                                                        }
                                                    }
                                                }

                                                ds.Tables[iWorksheetIndex].Rows.Add(dr);
                                            } // cells
                                        }
                                    }
                                } // rows
                            }
                        } // tables
                    }
                } // worksheets
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(ex.Message);
                System.Diagnostics.Debug.WriteLine(ex.StackTrace);
            }
        }
    }

    private static System.Data.DataSet ReadXmlFileColumnNames(string sXmlPath)
    {
        System.Data.DataSet ds = new System.Data.DataSet();

        if (System.IO.File.Exists(sXmlPath))
        {
            try
            {
                XmlDocument xmlDoc = new XmlDocument();

                xmlDoc.Load(sXmlPath);

                foreach (XmlNode xmlNode in xmlDoc.DocumentElement.ChildNodes)
                {
                    if (xmlNode.Name == "Worksheet")
                    {
                        if (null != xmlNode.Attributes["ss:Name"])
                        {
                            System.Data.DataTable dt = new System.Data.DataTable(xmlNode.Attributes["ss:Name"].Value);
                            ds.Tables.Add(dt);

                            foreach (XmlNode xmlNode2 in xmlNode.ChildNodes)
                            {
                                if (xmlNode2.Name == "Table")
                                {
                                    foreach (XmlNode xmlNode3 in xmlNode2.ChildNodes)
                                    {
                                        if (xmlNode3.Name == "Row")
                                        {
                                            // build array of column names
                                            if (xmlNode3.HasChildNodes)
                                            {
                                                foreach (XmlNode cellNode in xmlNode3.ChildNodes)
                                                {
                                                    if (cellNode.Name == "Cell")
                                                    {
                                                        string sValue = cellNode.InnerText.Replace("\r", "_");
                                                        string sValue2 = sValue.Replace("\n", "_");
                                                        string sColumnName = sValue2.Replace(" ", "_");

                                                        System.Diagnostics.Debug.WriteLine(sColumnName);

                                                        if (!string.IsNullOrEmpty(sColumnName))
                                                        {
                                                            System.Data.DataColumn myDataColumn = new System.Data.DataColumn();

                                                            myDataColumn.DataType = System.Type.GetType("System.String");
                                                            myDataColumn.ColumnName = sColumnName;
                                                            dt.Columns.Add(myDataColumn);
                                                        }
                                                    }
                                                }
                                            }

                                            break; // skip after first row
                                        }
                                    }
                                }

                                // break; // go to next worksheet
                            }
                        }
                    }
                }
            }
            catch
            {
            }
        }

        return ds;
    }
}

public class SavedSearches
{
    //HAADMSBUFFER aaApi_SQueryDataBufferSelectSubItems2(LONG lParQueryId, LONG lUserId, LONG lProjectId )
    [DllImport("dmscli.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr aaApi_SQueryDataBufferSelectSubItems2(int lParQueryId, int lUserId, int lProjectId);

    public static int GetSearchId(string sSearchName, bool bIsPersonal)
    {
        IntPtr iPtrBuffer = IntPtr.Zero;

        int iSearchId = 0;

        string sFolderPath = System.IO.Path.GetDirectoryName(sSearchName);

        string sBareSearchName = System.IO.Path.GetFileName(sSearchName);

        if (!string.IsNullOrEmpty(sFolderPath))
        {
            int iProjectID = -1;

            iProjectID = PWWrapper.aaApi_GetProjectIdByNamePath(sFolderPath);

            if (iProjectID > 0)
            {
                if (bIsPersonal)
                {
                    iPtrBuffer = aaApi_SQueryDataBufferSelectSubItems2(0, PWWrapper.aaApi_GetCurrentUserId(), iProjectID);
                }
                else
                {
                    iPtrBuffer = aaApi_SQueryDataBufferSelectSubItems2(0, 0, iProjectID);
                }

                int iCount = PWWrapper.aaApi_DmsDataBufferGetCount(iPtrBuffer);

                for (int i = 0; i < iCount; i++)
                {
                    string s = PWWrapper.aaApi_DmsDataBufferGetStringProperty(iPtrBuffer, (int)PWWrapper.SavedQueryProperty.SQRY_PROP_NAME, i);
                    int id = PWWrapper.aaApi_DmsDataBufferGetNumericProperty(iPtrBuffer, (int)PWWrapper.SavedQueryProperty.SQRY_PROP_QUERYID, i);

                    if (s.ToLower() == sBareSearchName.ToLower())
                    {
                        iSearchId = id;
                        break;
                    }
                }

                PWWrapper.aaApi_DmsDataBufferFree(iPtrBuffer);
            }
        }
        else
        {
            if (bIsPersonal)
            {
                iPtrBuffer = aaApi_SQueryDataBufferSelectSubItems2(0, PWWrapper.aaApi_GetCurrentUserId(), 0);
            }
            else
            {
                iPtrBuffer = PWWrapper.aaApi_SQueryDataBufferSelectAll();
            }

            int iCount = PWWrapper.aaApi_DmsDataBufferGetCount(iPtrBuffer);

            for (int i = 0; i < iCount; i++)
            {
                string s = PWWrapper.aaApi_DmsDataBufferGetStringProperty(iPtrBuffer, (int)PWWrapper.SavedQueryProperty.SQRY_PROP_NAME, i);
                int id = PWWrapper.aaApi_DmsDataBufferGetNumericProperty(iPtrBuffer, (int)PWWrapper.SavedQueryProperty.SQRY_PROP_QUERYID, i);

                if (s.ToLower() == sBareSearchName.ToLower())
                {
                    iSearchId = id;
                    break;
                }
            }

            PWWrapper.aaApi_DmsDataBufferFree(iPtrBuffer);
        }

        return iSearchId;
    }
}

public static class Extensions
{
    public static string SafeGet<TKey>(this SortedList<TKey, string> sortedList, TKey key)
    {
        if (sortedList.ContainsKey(key))
        {
            return sortedList[key];
        }

        return string.Empty;
    }


    public static bool AddWithCheck<TKey, TValue>(this SortedList<TKey, TValue> sortedList, TKey key, TValue value)
    {
        if (!sortedList.ContainsKey(key))
        {
            sortedList.Add(key, value);
            return true;
        }

        return false;
    }

    public static bool AddWithCheckNoNullsInKeysOrValues<TKey, TValue>(this SortedList<TKey, TValue> sortedList, TKey key, TValue value)
    {
        if (key != null && value != null)
        {
            if (!sortedList.ContainsKey(key))
            {
                sortedList.Add(key, value);
                return true;
            }
        }

        return false;
    }

    public static bool AddWithCheckNonZero(this SortedList<int, int> sortedList, int key, int value)
    {
        if (key != 0)
        {
            if (!sortedList.ContainsKey(key))
            {
                sortedList.Add(key, value);
                return true;
            }
        }

        return false;
    }


    public static bool AddFormat<TKey>(this SortedList<TKey, string> sortedList,
        TKey key,
        string formatString,
        params object[] argList)
    {
        if (!sortedList.ContainsKey(key))
        {
            sortedList.Add(key, string.Format(formatString, argList));
            return true;
        }

        return false;
    }

    public static string DumpContents<TKey, TValue>(this SortedList<TKey, TValue> sortedList,
        string keyFormatString)
    {
        StringBuilder sbContents = new StringBuilder();

        foreach (KeyValuePair<TKey, TValue> kvp in sortedList)
        {
            if (sbContents.Length == 0)
                sbContents.Append($"{string.Format(keyFormatString, kvp.Key.ToString())}{((kvp.Value == null) ? "" : kvp.Value.ToString())}");
            else
                sbContents.Append($"\n{string.Format(keyFormatString, kvp.Key.ToString())}{((kvp.Value == null) ? "" : kvp.Value.ToString())}");
        }

        return sbContents.ToString();
    }

    public static string DumpContents<TKey, TValue>(this Dictionary<TKey, TValue> dictionary,
        string keyFormatString)
    {
        StringBuilder sbContents = new StringBuilder();

        foreach (KeyValuePair<TKey, TValue> kvp in dictionary)
        {
            if (sbContents.Length == 0)
                sbContents.Append($"{string.Format(keyFormatString, kvp.Key.ToString())}{((kvp.Value == null) ? "" : kvp.Value.ToString())}");
            else
                sbContents.Append($"\n{string.Format(keyFormatString, kvp.Key.ToString())}{((kvp.Value == null) ? "" : kvp.Value.ToString())}");
        }

        return sbContents.ToString();
    }

}

public class Node<T> : IEqualityComparer, IEnumerable<T>, IEnumerable<Node<T>>
{
    public Node<T> Parent { get; private set; }
    public T Value { get; set; }
    private readonly List<Node<T>> _children = new List<Node<T>>();

    public Node(T value)
    {
        Value = value;
    }

    public Node<T> this[int index]
    {
        get
        {
            return _children[index];
        }
    }

    public Node<T> Add(T value, int index = -1)
    {
        var childNode = new Node<T>(value);
        Add(childNode, index);
        return childNode;
    }

    public void Add(Node<T> childNode, int index = -1)
    {
        if (index < -1)
        {
            throw new ArgumentException("The index can not be lower then -1");
        }
        if (index > Children.Count() - 1)
        {
            throw new ArgumentException("The index ({0}) can not be higher then index of the last iten. Use the AddChild() method without an index to add at the end".FormatInvariant(index));
        }
        if (!childNode.IsRoot)
        {
            throw new ArgumentException("The child node with value [{0}] can not be added because it is not a root node.".FormatInvariant(childNode.Value));
        }

        if (Root == childNode)
        {
            throw new ArgumentException("The child node with value [{0}] is the rootnode of the parent.".FormatInvariant(childNode.Value));
        }

        if (childNode.SelfAndDescendants.Any(n => this == n))
        {
            throw new ArgumentException("The childnode with value [{0}] can not be added to itself or its descendants.".FormatInvariant(childNode.Value));
        }
        childNode.Parent = this;
        if (index == -1)
        {
            _children.Add(childNode);
        }
        else
        {
            _children.Insert(index, childNode);
        }
    }

    public Node<T> AddFirstChild(T value)
    {
        var childNode = new Node<T>(value);
        AddFirstChild(childNode);
        return childNode;
    }

    public void AddFirstChild(Node<T> childNode)
    {
        Add(childNode, 0);
    }

    public Node<T> AddFirstSibling(T value)
    {
        var childNode = new Node<T>(value);
        AddFirstSibling(childNode);
        return childNode;
    }

    public void AddFirstSibling(Node<T> childNode)
    {
        Parent.AddFirstChild(childNode);
    }
    public Node<T> AddLastSibling(T value)
    {
        var childNode = new Node<T>(value);
        AddLastSibling(childNode);
        return childNode;
    }

    public void AddLastSibling(Node<T> childNode)
    {
        Parent.Add(childNode);
    }

    public Node<T> AddParent(T value)
    {
        var newNode = new Node<T>(value);
        AddParent(newNode);
        return newNode;
    }

    public void AddParent(Node<T> parentNode)
    {
        if (!IsRoot)
        {
            throw new ArgumentException("This node [{0}] already has a parent".FormatInvariant(Value), "parentNode");
        }
        parentNode.Add(this);
    }

    public IEnumerable<Node<T>> Ancestors
    {
        get
        {
            if (IsRoot)
            {
                return Enumerable.Empty<Node<T>>();
            }
            return Parent.ToIEnumarable().Concat(Parent.Ancestors);
        }
    }

    public IEnumerable<Node<T>> Descendants
    {
        get
        {
            return SelfAndDescendants.Skip(1);
        }
    }

    public IEnumerable<Node<T>> Children
    {
        get
        {
            return _children;
        }
    }

    public IEnumerable<Node<T>> Siblings
    {
        get
        {
            return SelfAndSiblings.Where(Other);

        }
    }

    private bool Other(Node<T> node)
    {
        return !ReferenceEquals(node, this);
    }

    public IEnumerable<Node<T>> SelfAndChildren
    {
        get
        {
            return this.ToIEnumarable().Concat(Children);
        }
    }

    public IEnumerable<Node<T>> SelfAndAncestors
    {
        get
        {
            return this.ToIEnumarable().Concat(Ancestors);
        }
    }

    public IEnumerable<Node<T>> SelfAndDescendants
    {
        get
        {
            return this.ToIEnumarable().Concat(Children.SelectMany(c => c.SelfAndDescendants));
        }
    }

    public IEnumerable<Node<T>> SelfAndSiblings
    {
        get
        {
            if (IsRoot)
            {
                return this.ToIEnumarable();
            }
            return Parent.Children;

        }
    }

    public IEnumerable<Node<T>> All
    {
        get
        {
            return Root.SelfAndDescendants;
        }
    }


    public IEnumerable<Node<T>> SameLevel
    {
        get
        {
            return SelfAndSameLevel.Where(Other);

        }
    }

    public int Level
    {
        get
        {
            return Ancestors.Count();
        }
    }

    public IEnumerable<Node<T>> SelfAndSameLevel
    {
        get
        {
            return GetNodesAtLevel(Level);
        }
    }

    public IEnumerable<Node<T>> GetNodesAtLevel(int level)
    {
        return Root.GetNodesAtLevelInternal(level);
    }

    private IEnumerable<Node<T>> GetNodesAtLevelInternal(int level)
    {
        if (level == Level)
        {
            return this.ToIEnumarable();
        }
        return Children.SelectMany(c => c.GetNodesAtLevelInternal(level));
    }

    public Node<T> Root
    {
        get
        {
            return SelfAndAncestors.Last();
        }
    }

    public void Disconnect()
    {
        if (IsRoot)
        {
            throw new InvalidOperationException("The root node [{0}] can not get disconnected from a parent.".FormatInvariant(Value));
        }
        Parent._children.Remove(this);
        Parent = null;
    }

    public bool IsRoot
    {
        get { return Parent == null; }
    }

    IEnumerator<T> IEnumerable<T>.GetEnumerator()
    {
        return _children.Values().GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return _children.GetEnumerator();
    }

    public IEnumerator<Node<T>> GetEnumerator()
    {
        return _children.GetEnumerator();
    }

    public override string ToString()
    {
        return Value.ToString();
    }

    public static IEnumerable<Node<T>> CreateTree<TId>(IEnumerable<T> values, Func<T, TId> idSelector, Func<T, TId?> parentIdSelector)
        where TId : struct
    {
        var valuesCache = values.ToList();
        if (!valuesCache.Any())
            return Enumerable.Empty<Node<T>>();
        T itemWithIdAndParentIdIsTheSame = valuesCache.FirstOrDefault(v => IsSameId(idSelector(v), parentIdSelector(v)));
        if (itemWithIdAndParentIdIsTheSame != null) // Hier verwacht je ook een null terug te kunnen komen
        {
            throw new ArgumentException("At least one value has the same Id and parentId [{0}]".FormatInvariant(itemWithIdAndParentIdIsTheSame));
        }

        var nodes = valuesCache.Select(v => new Node<T>(v));
        return CreateTree(nodes, idSelector, parentIdSelector);

    }

    public static IEnumerable<Node<T>> CreateTree<TId>(IEnumerable<Node<T>> rootNodes, Func<T, TId> idSelector, Func<T, TId?> parentIdSelector)
        where TId : struct

    {
        var rootNodesCache = rootNodes.ToList();
        var duplicates = rootNodesCache.Duplicates(n => n).ToList();
        if (duplicates.Any())
        {
            throw new ArgumentException("One or more values contains {0} duplicate keys. The first duplicate is: [{1}]".FormatInvariant(duplicates.Count, duplicates[0]));
        }

        foreach (var rootNode in rootNodesCache)
        {
            var parentId = parentIdSelector(rootNode.Value);
            var parent = rootNodesCache.FirstOrDefault(n => IsSameId(idSelector(n.Value), parentId));

            if (parent != null)
            {
                parent.Add(rootNode);
            }
            else if (parentId != null)
            {

                throw new ArgumentException("A value has the parent ID [{0}] but no other nodes has this ID".FormatInvariant(parentId.Value));
            }
        }
        var result = rootNodesCache.Where(n => n.IsRoot);
        return result;
    }


    private static bool IsSameId<TId>(TId id, TId? parentId)
        where TId : struct
    {
        return parentId != null && id.Equals(parentId.Value);
    }

#region Equals en ==

    public static bool operator ==(Node<T> value1, Node<T> value2)
    {
        if ((object)(value1) == null && (object)value2 == null)
        {
            return true;
        }
        return ReferenceEquals(value1, value2);
    }

    public static bool operator !=(Node<T> value1, Node<T> value2)
    {
        return !(value1 == value2);
    }

    public override bool Equals(Object anderePeriode)
    {
        var valueThisType = anderePeriode as Node<T>;
        return this == valueThisType;
    }

    public bool Equals(Node<T> value)
    {
        return this == value;
    }

    public bool Equals(Node<T> value1, Node<T> value2)
    {
        return value1 == value2;
    }

    bool IEqualityComparer.Equals(object value1, object value2)
    {
        var valueThisType1 = value1 as Node<T>;
        var valueThisType2 = value2 as Node<T>;

        return Equals(valueThisType1, valueThisType2);
    }

    public int GetHashCode(object obj)
    {
        return GetHashCode(obj as Node<T>);
    }

    public override int GetHashCode()
    {
        return GetHashCode(this);
    }

    public int GetHashCode(Node<T> value)
    {
        return base.GetHashCode();
    }

#endregion
}

public static class NodeExtensions
{
    public static IEnumerable<T> Values<T>(this IEnumerable<Node<T>> nodes)
    {
        return nodes.Select(n => n.Value);
    }
}

public static class OtherExtensions
{
    public static IEnumerable<TSource> Duplicates<TSource, TKey>(this IEnumerable<TSource> source, Func<TSource, TKey> selector)
    {
        var grouped = source.GroupBy(selector);
        var moreThen1 = grouped.Where(i => i.IsMultiple());

        return moreThen1.SelectMany(i => i);
    }

    public static bool IsMultiple<T>(this IEnumerable<T> source)
    {
        var enumerator = source.GetEnumerator();
        return enumerator.MoveNext() && enumerator.MoveNext();
    }

    public static IEnumerable<T> ToIEnumarable<T>(this T item)
    {
        yield return item;
    }

    public static string FormatInvariant(this string text, params object[] parameters)
    {
        // This is not the "real" implementation, but that would go out of Scope
        return string.Format(CultureInfo.InvariantCulture, text, parameters);
    }
}

public class FolderClass
{
    public int Id { get; set; }
    public int? ParentId { get; set; }
    public string Name { get; set; }
    public string Description { get; set; }
    public string CalculatedPath { get; set; } = string.Empty;
    public bool PathDetermined { get; set; } = false;

    public static string GetPath(Node<FolderClass> folderNode, bool bUseDescription)
    {
        if (folderNode == null)
            return string.Empty;

        StringBuilder sbPath = new StringBuilder();

        // builds from bottom up
        foreach (var node in folderNode.Ancestors)
        {
            // if (!string.IsNullOrEmpty(node.Value.CalculatedPath))
            if (node.Value.PathDetermined)
            {
                sbPath.Insert(0, $"\\{node.Value.CalculatedPath}");
                break;
            }
            else
            {
                if (bUseDescription)
                    sbPath.Insert(0, $"\\{node.Value.Description}");
                else
                    sbPath.Insert(0, $"\\{node.Value.Name}");
            }
        }

        if (bUseDescription)
            sbPath.Append($"\\{folderNode.Value.Description}");
        else
            sbPath.Append($"\\{folderNode.Value.Name}");

        string sPathToReturn = sbPath.ToString().StartsWith(@"\\") ? sbPath.ToString().Substring(2) : sbPath.ToString();

        // folderNode.Value.CalculatedPath = sPathToReturn.StartsWith(@"\") ? sPathToReturn.Substring(1) : sPathToReturn;
        // folderNode.Value.PathDetermined = true;

        return sPathToReturn.StartsWith(@"\") ? sPathToReturn.Substring(1) : sPathToReturn;
        // return folderNode.Value.CalculatedPath;
    }
}
