using System.Windows.Forms;
using System;

namespace QRCodeAddInForMSCE
{
    /// <summary>
    /// Keyins Class
    /// </summary>
    public sealed class Keyins
    {
        private static bool mcmMain_GetDocumentIdByFilePath(string sFileName, int iValidateWithChkl,
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

        public static void Place(string unparsed)
        {
            if (string.IsNullOrEmpty(unparsed))
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
                    {
                        int GMAIL_PROJECTWISE_WEB_VIEW_SETTING = -5250;

                        string sWebViewURL = PWWrapper.GetPWStringSetting(GMAIL_PROJECTWISE_WEB_VIEW_SETTING);

                        string sProjectGUIDString = PWWrapper.GetProjectGuidStringFromId(iProjectNo);
                        string sDocumentGUIDString = PWWrapper.GetGuidStringFromIds(iProjectNo, iDocumentNo);

                        if (!string.IsNullOrEmpty(sWebViewURL))
                        {
                            unparsed = $"{sWebViewURL}?project={sProjectGUIDString}&item={sDocumentGUIDString}";
                        }
                        else
                        {
                            BPSUtilities.WriteLog("No web view link address set.");
                        }
                    }
                    else
                    {
                        BPSUtilities.WriteLog("No integrated session of ProjectWise.");
                    }
                }
                else
                {
                    BPSUtilities.WriteLog("No integrated session of ProjectWise.");
                }

                if (string.IsNullOrEmpty(unparsed))
                {
                    unparsed = "www.bentley.com";
                }

                BPSUtilities.WriteLog($"Make code for this: {unparsed}");

                PlaceQRCode.InstallNewInstance(unparsed);
            }
            else
            {
                BPSUtilities.WriteLog($"Make code for this: {unparsed}");

                PlaceQRCode.InstallNewInstance(unparsed);
            }
        }
    }
}