using System.Collections.Generic;
using System.Text;

namespace QRCodeAddInForMSCE
{
    /// <summary>
    /// 
    /// </summary>
    [Bentley.MstnPlatformNET.AddInAttribute(MdlTaskID = "QRCodeAddInForMSCE")]
    public sealed class QRCodeAddInForMSCE : Bentley.MstnPlatformNET.AddIn
    {
        private static QRCodeAddInForMSCE s_instance = null;

        /// <summary>
        /// Active DgnFile and DgnModel.
        /// </summary>
        private Bentley.DgnPlatformNET.DgnFile m_ActiveDgnFile;
        private Bentley.DgnPlatformNET.DgnModel m_ActiveModel;

        public static SortedList<string, Bentley.DgnPlatformNET.Elements.Element> ListOfPolyhedra =
            new SortedList<string, Bentley.DgnPlatformNET.Elements.Element>(System.StringComparer.CurrentCultureIgnoreCase);

        public static Bentley.DgnPlatformNET.Elements.Element StaticElement = null;

        public QRCodeAddInForMSCE(System.IntPtr mdlDesc)
            : base(mdlDesc)
        {
            s_instance = this;
        }

        internal static QRCodeAddInForMSCE Instance
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

            return 0;
        }
    }
}
