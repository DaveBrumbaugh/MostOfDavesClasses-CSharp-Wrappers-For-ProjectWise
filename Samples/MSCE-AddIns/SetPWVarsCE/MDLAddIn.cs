using System;
using System.Collections.Generic;

namespace SetPWVarsCE
{
    /// <summary>
    /// 
    /// </summary>
    [Bentley.MstnPlatformNET.AddInAttribute(MdlTaskID = "SetPWVarsCE")]
    public sealed class SetPWVarsCE : Bentley.MstnPlatformNET.AddIn
    {
        private static SetPWVarsCE s_instance = null;

        /// <summary>
        /// Active DgnFile and DgnModel.
        /// </summary>
        private Bentley.DgnPlatformNET.DgnFile m_ActiveDgnFile;
        private Bentley.DgnPlatformNET.DgnModel m_ActiveModel;

        public SetPWVarsCE(System.IntPtr mdlDesc)
            : base(mdlDesc)
        {
            s_instance = this;
        }

        internal static SetPWVarsCE Instance
        {
            get
            {
                return s_instance;
            }
        }

        protected override int Run(string[] commandLine)
        {
            s_instance = this;

            m_ActiveModel = Bentley.MstnPlatformNET.Session.Instance.GetActiveDgnModel();
            m_ActiveDgnFile = Bentley.MstnPlatformNET.Session.Instance.GetActiveDgnFile();

            this.ReloadEvent += SetPWVarsCE_ReloadEvent;
            this.UnloadingEvent += SetPWVarsCE_UnloadingEvent1;
            this.NewDesignFileEvent += SetPWVarsCE_NewDesignFileEvent;

            return 0;
        }

        public static void UpdatePWEnvVars(int lProjectNo, int lDocumentNo)
        {
            if (1 == PWWrapper.aaApi_SelectDocument(lProjectNo, lDocumentNo))
            {
                Bentley.DgnPlatformNET.ConfigurationVariableLevel level = Bentley.DgnPlatformNET.ConfigurationVariableLevel.Application;

                Bentley.DgnPlatformNET.ConfigurationManager.DefineVariable("PWVAR_VAULTID", $"{lProjectNo}", level);
                Bentley.DgnPlatformNET.ConfigurationManager.DefineVariable("PWVAR_DOCID", $"{lDocumentNo}", level);
                Bentley.DgnPlatformNET.ConfigurationManager.DefineVariable("PWVAR_ORIGINALNO", 
                    $"{PWWrapper.aaApi_GetDocumentNumericProperty(PWWrapper.DocumentProperty.OriginalNumber, 0)}", level);
                Bentley.DgnPlatformNET.ConfigurationManager.DefineVariable("PWVAR_DOCNAME",
                    $"{PWWrapper.aaApi_GetDocumentStringProperty(PWWrapper.DocumentProperty.Name, 0)}", level);
                Bentley.DgnPlatformNET.ConfigurationManager.DefineVariable("PWVAR_FILENAME",
                    $"{PWWrapper.aaApi_GetDocumentStringProperty(PWWrapper.DocumentProperty.FileName, 0)}", level);
                Bentley.DgnPlatformNET.ConfigurationManager.DefineVariable("PWVAR_DOCDESC",
                    $"{PWWrapper.aaApi_GetDocumentStringProperty(PWWrapper.DocumentProperty.Desc, 0)}", level);
                Bentley.DgnPlatformNET.ConfigurationManager.DefineVariable("PWVAR_DOCVERSION",
                    $"{PWWrapper.aaApi_GetDocumentStringProperty(PWWrapper.DocumentProperty.Version, 0)}", level);
                Bentley.DgnPlatformNET.ConfigurationManager.DefineVariable("PWVAR_DOCCREATETIME",
                    $"{PWWrapper.aaApi_GetDocumentStringProperty(PWWrapper.DocumentProperty.CreateTime, 0)}", level);
                Bentley.DgnPlatformNET.ConfigurationManager.DefineVariable("PWVAR_DOCUPDATETIME",
                    $"{PWWrapper.aaApi_GetDocumentStringProperty(PWWrapper.DocumentProperty.UpdateTime, 0)}", level);
                Bentley.DgnPlatformNET.ConfigurationManager.DefineVariable("PWVAR_DOCFILEUPDATETIME",
                    $"{PWWrapper.aaApi_GetDocumentStringProperty(PWWrapper.DocumentProperty.FileUpdateTime, 0)}", level);
                Bentley.DgnPlatformNET.ConfigurationManager.DefineVariable("PWVAR_DOCLASTRTLOCKTIME",
                    $"{PWWrapper.aaApi_GetDocumentStringProperty(PWWrapper.DocumentProperty.LastRtLockTime, 0)}", level);
                Bentley.DgnPlatformNET.ConfigurationManager.DefineVariable("PWVAR_DOCWORKFLOW",
                    $"{PWWrapper.GetWorkflowName(PWWrapper.aaApi_GetDocumentNumericProperty(PWWrapper.DocumentProperty.WorkFlowID, 0))}", level);
                Bentley.DgnPlatformNET.ConfigurationManager.DefineVariable("PWVAR_DOCWORKFLOWSTATE",
                    $"{PWWrapper.GetStateName(PWWrapper.aaApi_GetDocumentNumericProperty(PWWrapper.DocumentProperty.StateID, 0))}", level);
                Bentley.DgnPlatformNET.ConfigurationManager.DefineVariable("PWVAR_VAULTPATHNAME",
                    $"{PWWrapper.GetProjectNamePath2(lProjectNo)}", level);
                Bentley.DgnPlatformNET.ConfigurationManager.DefineVariable("PWVAR_FULLFILEPATHNAME",
                    $"{PWWrapper.GetDocumentNamePath(lProjectNo, lDocumentNo)}", level);
                Bentley.DgnPlatformNET.ConfigurationManager.DefineVariable("PWVAR_DOCGUID",
                    $"{PWWrapper.GetGuidStringFromIds(lProjectNo, lDocumentNo)}", level);
                Bentley.DgnPlatformNET.ConfigurationManager.DefineVariable("PWVAR_DOCLINK",
                    $"{PWWrapper.GetDocumentURL(lProjectNo, lDocumentNo)}", level);

                if (1 == PWWrapper.aaApi_SelectProject(lProjectNo))
                {
                    Bentley.DgnPlatformNET.ConfigurationManager.DefineVariable("PWVAR_VAULTNAME",
                        $"{PWWrapper.aaApi_GetProjectStringProperty(PWWrapper.ProjectProperty.Name, 0)}", level);
                    Bentley.DgnPlatformNET.ConfigurationManager.DefineVariable("PWVAR_VAULTDESC",
                        $"{PWWrapper.aaApi_GetProjectStringProperty(PWWrapper.ProjectProperty.Desc, 0)}", level);
                }

                System.Collections.Generic.SortedList<string, string> slProps = PWWrapper.GetProjectPropertyValuesInList(lProjectNo);

                foreach (System.Collections.Generic.KeyValuePair<string, string> kvp in slProps)
                {
                    Bentley.DgnPlatformNET.ConfigurationManager.DefineVariable($"PWVAR_PROJPROP_{kvp.Key.ToUpper().Replace(" ", "_")}",
                        kvp.Value, level);
                }

                System.Collections.Generic.SortedList<string, string> slAttrs = PWWrapper.GetAllAttributeColumnValuesInList(lProjectNo, lDocumentNo);

                foreach (System.Collections.Generic.KeyValuePair<string, string> kvp in slAttrs)
                {
                    Bentley.DgnPlatformNET.ConfigurationManager.DefineVariable($"PWVAR_ATTR_{kvp.Key.ToUpper().Replace(" ", "_")}",
                        kvp.Value, level);
                }
            } // document selected
            else
            {
                BPSUtilities.WriteLog("Could not select document.");
            }
        }

        private void ListReferences(bool bIntegrated)
        {
            BPSUtilities.WriteLog($"Active Model Name: {Bentley.MstnPlatformNET.Session.Instance.GetActiveDgnModel().GetModelInfo().Name}");

            Bentley.DgnPlatformNET.DgnAttachmentCollection col = Bentley.MstnPlatformNET.Session.Instance.GetActiveDgnModel().GetDgnAttachments();

            int iAttachmentNo = 1;

            try
            {
                foreach (Bentley.DgnPlatformNET.DgnAttachment dgnAttachment in col.GetEnumerator().ConvertToList())
                {
                    BPSUtilities.WriteLog($"Attachment: {iAttachmentNo++}");

                    try
                    {
                        BPSUtilities.WriteLog($"Is Missing: {(dgnAttachment.IsMissingFile ? "True" : "False")}");
                        BPSUtilities.WriteLog($"Logical Name: {dgnAttachment.LogicalName}");
                        BPSUtilities.WriteLog($"Attachment Filename: {dgnAttachment.AttachFileName}");
                        BPSUtilities.WriteLog($"Attachment Model Name: {dgnAttachment.AttachModelName}");
                        BPSUtilities.WriteLog($"Attachment Attach Full File Spec: {dgnAttachment.GetAttachFullFileSpec(false)}");

                        Bentley.DgnPlatformNET.DgnDocumentMoniker dgnMoniker = dgnAttachment.GetAttachMoniker();

                        if (dgnMoniker != null)
                        {
                            BPSUtilities.WriteLog($"Parent Search Path: {dgnMoniker.ParentSearchPath}");
                            BPSUtilities.WriteLog($"Portable Name: {dgnMoniker.PortableName}");
                            BPSUtilities.WriteLog($"Provider Id: {dgnMoniker.ProviderId}");
                            BPSUtilities.WriteLog($"Saved File Name: {dgnMoniker.SavedFileName}");
                            BPSUtilities.WriteLog($"Short Display Name: {dgnMoniker.ShortDisplayName}");
                        }
                        else
                        {
                            BPSUtilities.WriteLog("No Document Moniker found.");
                        }
                    }
                    catch (Exception ex)
                    {
                        BPSUtilities.WriteLog($"ListReferences Error: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                BPSUtilities.WriteLog($"ListReferences Error: {ex.Message}");
            }
        }

        private void SetPWVarsCE_NewDesignFileEvent(Bentley.MstnPlatformNET.AddIn sender, NewDesignFileEventArgs eventArgs)
        {
            if (eventArgs.WhenCode == NewDesignFileEventArgs.When.AfterDesignFileOpen)
            {
                string sFileName = Bentley.MstnPlatformNET.Session.Instance.GetActiveDgnFile().GetFileName();

                BPSUtilities.WriteLog($"Filename is '{sFileName}'");

                int iProjectNo = 0, iDocumentNo = 0;

                PWWrapper.aaApi_Initialize(0);

                if (mcmMain_GetDocumentIdByFilePath(sFileName, 1,
                    ref iProjectNo, ref iDocumentNo))
                {
                    BPSUtilities.WriteLog($"IDs: {iProjectNo}, {iDocumentNo}");

                    if (iProjectNo > 0 && iDocumentNo > 0)
                        UpdatePWEnvVars(iProjectNo, iDocumentNo);
                    else
                        BPSUtilities.WriteLog("No integrated session of ProjectWise.");
                }
                else
                {
                    BPSUtilities.WriteLog("No integrated session of ProjectWise.");
                }

                ListReferences(true);
            }
        }

        private void SetPWVarsCE_UnloadingEvent1(Bentley.MstnPlatformNET.AddIn sender, UnloadingEventArgs eventArgs)
        {
            base.OnUnloading(eventArgs);
        }

        private void SetPWVarsCE_ReloadEvent(Bentley.MstnPlatformNET.AddIn sender, ReloadEventArgs eventArgs)
        {
        }

        private bool mcmMain_GetDocumentIdByFilePath(string sFileName, int iValidateWithChkl,
            ref int iProjectNo, ref int iDocumentNo)
        {
            bool bRetVal = false;

            Guid[] docGuids = new Guid[1];
            int iNumGuids = 0;

            try
            {
                IntPtr pGuid = IntPtr.Zero;

                int iRetVal = PWWrapper.aaApi_GetGuidsFromFileName(ref pGuid, ref iNumGuids, sFileName, iValidateWithChkl);

                if (iNumGuids == 1)
                {
                    Guid docGuid = (Guid)System.Runtime.InteropServices.Marshal.PtrToStructure(pGuid, typeof(Guid));

                    if (1 == PWWrapper.aaApi_GUIDSelectDocument(ref docGuid))
                    {
                        bRetVal = true;

                        iProjectNo =
                            PWWrapper.aaApi_GetDocumentNumericProperty(PWWrapper.DocumentProperty.ProjectID, 0);
                        iDocumentNo =
                            PWWrapper.aaApi_GetDocumentNumericProperty(PWWrapper.DocumentProperty.ID, 0);
                    }
                }
            }
            catch (Exception ex)
            {
                BPSUtilities.WriteLog($"Error: {ex.Message}\n{ex.StackTrace}");
            }

            return bRetVal;
        }

    }

    public static class Extensions
    {
        public static List<T> ConvertToList<T>(this IEnumerator<T> e)
        {
            var list = new List<T>();
            while (e.MoveNext())
            {
                list.Add(e.Current);
            }
            return list;
        }
    }

}
