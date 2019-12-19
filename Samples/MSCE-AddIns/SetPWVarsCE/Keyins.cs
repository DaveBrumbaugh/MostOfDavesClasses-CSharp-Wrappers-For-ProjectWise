using System.Windows.Forms;
using System;
using System.Text;
using System.Data;
using System.Runtime.InteropServices;
using System.Collections;
using System.IO;
using System.Collections.Generic;

namespace SetPWVarsCE
{
    /// <summary>
    /// Keyins Class
    /// </summary>
    public sealed class Keyins
    {
        public static void SaveReport(string unparsed)
        {
            string sReportFile = Path.Combine(Path.GetTempPath(), $"{BPSUtilities.GetARandomString(8, "abcdefghijklmnopqrstuvwxyz")}.rpt");

            try
            {
                Bentley.DgnPlatformNET.ConfigurationManager.WriteActiveConfigurationSummary(sReportFile);
            }
            catch (Exception ex)
            {
                BPSUtilities.WriteLog("Error: {0}\n{1}", ex.Message, ex.StackTrace);
            }

            if (File.Exists(sReportFile))
            {
                StringBuilder sbDSN = new StringBuilder(1024);

                bool bCheckPW = false;

                if (PWWrapper.aaApi_GetActiveDatasourceName(sbDSN, sbDSN.Capacity))
                {
                    bCheckPW = true;
                }

                DataSet ds = new DataSet();
                ds.Tables.Add(new DataTable("ConfigurationVariables"));

                ds.Tables[0].Columns.Add(new DataColumn("Name", typeof(string)));
                ds.Tables[0].Columns.Add(new DataColumn("Level", typeof(string)));
                ds.Tables[0].Columns.Add(new DataColumn("Value", typeof(string)));
                ds.Tables[0].Columns.Add(new DataColumn("ExpandedValue", typeof(string)));
                ds.Tables[0].Columns.Add(new DataColumn("InvalidPWPaths", typeof(string)));

                SortedList<string, int> slVariableLevels = new SortedList<string, int>()
                {
                    { "Application", (int)Bentley.DgnPlatformNET.ConfigurationVariableLevel.Application},
                    { "Predefined", (int)Bentley.DgnPlatformNET.ConfigurationVariableLevel.Predefined},
                    { "Organization", (int)Bentley.DgnPlatformNET.ConfigurationVariableLevel.Organization},
                    { "WorkSpace", (int)Bentley.DgnPlatformNET.ConfigurationVariableLevel.WorkSpace},
                    { "System Environment", (int)Bentley.DgnPlatformNET.ConfigurationVariableLevel.SystemEnvironment },
                    { "System", (int)Bentley.DgnPlatformNET.ConfigurationVariableLevel.System },
                    { "User", (int)Bentley.DgnPlatformNET.ConfigurationVariableLevel.User },
                    { "WorkSet", (int)Bentley.DgnPlatformNET.ConfigurationVariableLevel.WorkSet },
                    { "Role", (int)Bentley.DgnPlatformNET.ConfigurationVariableLevel.Role }
                };

                using (StreamReader sr = new StreamReader(sReportFile))
                {
                    while (!sr.EndOfStream)
                    {
                        string sLine = sr.ReadLine();

                        if (sLine.StartsWith("%level"))
                        {
                            DataRow dr = ds.Tables[0].NewRow();

                            string[] sSplits = sLine.Split(new string[1] { "  " }, StringSplitOptions.RemoveEmptyEntries);

                            if (sSplits.Length > 1)
                            {
                                dr["Level"] = sSplits[1];

                                string sLine2 = sr.ReadLine();

                                string[] sSplits2 = sLine2.Split(new string[1] { " = " }, StringSplitOptions.RemoveEmptyEntries);

                                if (sSplits2.Length > 1)
                                {
                                    dr["Name"] = sSplits2[0].Trim();
                                    dr["Value"] = sSplits2[1].Trim();

                                    string sExpandedValue = string.Empty;

                                    if (slVariableLevels.ContainsKey(sSplits[1]))
                                    {
                                        sExpandedValue = Bentley.DgnPlatformNET.ConfigurationManager.GetVariable(sSplits2[0].Trim(),
                                            (Bentley.DgnPlatformNET.ConfigurationVariableLevel)slVariableLevels[sSplits[1]]);
                                    }
                                    else
                                    {
                                        sExpandedValue = Bentley.DgnPlatformNET.ConfigurationManager.GetVariable(sSplits2[0].Trim());
                                    }

                                    dr["ExpandedValue"] = sExpandedValue;

                                    if (sExpandedValue.Contains("pw:") && bCheckPW)
                                    {
                                        string[] sSplits3 = sExpandedValue.Split(";".ToCharArray());

                                        SortedList<string, string> slUniqueValues = new SortedList<string, string>();

                                        foreach (string sSplit in sSplits3)
                                            slUniqueValues.AddWithCheck(sSplit, sSplit);

                                        StringBuilder sbInvalidPaths = new StringBuilder();

                                        foreach (string de in slUniqueValues.Keys)
                                        {
                                            if (de.ToLower().StartsWith("pw:"))
                                            {
                                                if (GetFolderNo(de.ToString()) < 1)
                                                {
                                                    if (sbInvalidPaths.Length > 0)
                                                        sbInvalidPaths.Append(";");

                                                    sbInvalidPaths.Append(de);
                                                }
                                            }
                                        }

                                        dr["InvalidPWPaths"] = sbInvalidPaths.ToString();
                                    }

                                    ds.Tables[0].Rows.Add(dr);
                                }
                            }
                        }
                    } // for each pair of lines in the file
                }

                if (!string.IsNullOrEmpty(unparsed))
                {
                    if (!unparsed.ToLower().EndsWith(".xlsx"))
                        unparsed += ".xlsx";
                }
                else
                {
                    SaveFileDialog dlg = new SaveFileDialog();
                    dlg.Title = "Select Configuration File Report Output Location";
                    dlg.Filter = "XLSX Files|*.xlsx|All Files|*.*";
                    dlg.DefaultExt = ".xlsx";
                    dlg.AddExtension = true;

                    if (dlg.ShowDialog() == DialogResult.OK)
                    {
                        unparsed = dlg.FileName;
                    }
                }

                if (ds.Tables[0].Rows.Count > 0 && !string.IsNullOrEmpty(unparsed))
                {
                    try
                    {
                        XLSXDataSetTools.DataSetToXLSXFast(ds, unparsed);

                        if (File.Exists(unparsed))
                            MessageBox.Show($"Wrote '{unparsed}'", "SetPWVarsCE");

                    }
                    catch (Exception ex)
                    {
                        BPSUtilities.WriteLog($"{ex.Message}\n{ex.StackTrace}");
                        MessageBox.Show($"Error writing '{unparsed}'", "SetPWVarsCE");
                    }
                }
            }
        }

        public static int GetFolderNo(string sFolderPath)
        {
            if (sFolderPath.ToLower().StartsWith("pw:"))
            {
                int iIndex = sFolderPath.ToLower().IndexOf(@"\documents\");

                if (iIndex == -1)
                    iIndex = sFolderPath.ToLower().IndexOf(@"/documents/");

                if (iIndex > -1)
                {
                    string sPWPath = sFolderPath.Substring(iIndex + @"/documents".Length);

                    int iFolderNo = PWWrapper.ProjectNoFromPath(sPWPath);

                    if (iFolderNo < 1)
                    {
                        iFolderNo = PWWrapper.ProjectNoFromPath(System.IO.Path.GetDirectoryName(sPWPath));

                        if (iFolderNo > 0)
                        {
                            string sFileName = Path.GetFileName(sPWPath);

                            if (PWWrapper.aaApi_SelectDocumentsByNameProp(iFolderNo, sPWPath, null, null, null) < 1)
                                iFolderNo = 0;
                        }
                    }

                    return iFolderNo;
                }
            }

            return 0;
        }
    }
}