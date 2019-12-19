using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using System.Data;
using ExcelDataReader;

namespace SetPWVarsCE
{
    public class PWUriParser : GenericUriParser
    {
        //You may want to have your constructor do more, but it usually
        //isn't necesssary. See the MSDN documentation for
        //GenericUriParserOptions for a full explanation of what it does.
        //Basically it lets you define escaping rules and the presence of 
        //certain URI fields.
        public PWUriParser(GenericUriParserOptions options)
            : base(options)
        { }

        private string defaultPort = string.Empty;

        protected override void InitializeAndValidate(Uri uri,
            out UriFormatException parsingError)
        {
            //This function is called whenever a new Uri is created
            //whose scheme matches the one registered to this parser
            //(more on that later). If the Uri doesn't meet
            //certain specifications, set parsingError to an appropriate
            //UriFormatException.

            parsingError = null;
        }

        protected override void OnRegister(string schemeName,
            int defaultPort)
        {
            //This event is fired whenever your register a UriParser
            //(more on that later). The only use I can think of for this
            //is storing the default port when a UriParser matching the
            //correct scheme is registered.
            if (schemeName == "pw") this.defaultPort = defaultPort.ToString();
        }

        protected override bool IsWellFormedOriginalString(Uri uri)
        {
            //This method is similar to InitializeAndValidate. The
            //difference is that a valid URI is not necessarily
            //well-formed. You can use this to enforce certain
            //formatting rules if you wish.

            return true;
        }
    }

    public class XLSXDataSetTools
    {
        // man, this is slow.  Would love to find another way...
        private static System.Data.DataSet ReadData(string sWBName)
        {
            System.Data.DataSet ds = new System.Data.DataSet();

            if (System.IO.File.Exists(sWBName))
            {
                try
                {
                    try
                    {
                        UriParser.Register(new PWUriParser(GenericUriParserOptions.AllowEmptyAuthority |
                            GenericUriParserOptions.DontCompressPath | GenericUriParserOptions.DontConvertPathBackslashes |
                            GenericUriParserOptions.DontUnescapePathDotsAndSlashes | GenericUriParserOptions.GenericAuthority |
                            GenericUriParserOptions.NoFragment), "pw", 5800);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine(string.Format("Error: {0}\n{1}", ex.Message, ex.StackTrace));
                    }

                    var wb = new XLWorkbook(sWBName);

                    foreach (var ws in wb.Worksheets)
                    {
                        if (ds.Tables.Contains(ws.Name))
                            continue;

                        // BPSUtilities.WriteLog("Reading table '{0}'...", ws.Name);

                        DataTable dt = new DataTable(ws.Name);

                        int iIndex = 1;

                        while (!string.IsNullOrEmpty(ws.Row(1).Cell(iIndex).Value.ToString()))
                        {
                            string sColumnName = ws.Row(1).Cell(iIndex).Value.ToString();

                            try
                            {
                                dt.Columns.Add(new DataColumn(sColumnName, ws.Row(2).Cell(iIndex).Value.GetType()));
                            }
                            catch
                            {
                                try
                                {
                                    dt.Columns.Add(new DataColumn(sColumnName, Type.GetType("System.String")));
                                }
                                catch (Exception ex)
                                {
                                    BPSUtilities.WriteLog("Error: {0}\n{1}", ex.Message, ex.StackTrace);

                                    // logic will really be screwed up if column cant' get added
                                    return null;
                                }
                            }

                            iIndex++;
                        } // for each column

                        if (dt.Columns.Count == 0)
                            continue;

                        int iNumCols = dt.Columns.Count;

                        int iRow = 2;

                        while (!ws.Row(iRow).IsEmpty())
                        {
                            DataRow dr = dt.NewRow();

                            for (int iCol = 1; iCol <= iNumCols; iCol++)
                            {
                                try
                                {
                                    IXLCell cell = ws.Row(iRow).Cell(iCol);

                                    if (cell.Hyperlink != null)
                                    {
                                        if (cell.Hyperlink.ExternalAddress != null)
                                        {
                                            dr[iCol - 1] = cell.Hyperlink.ExternalAddress.OriginalString;
                                        }
                                        else
                                        {
                                            dr[iCol - 1] = cell.Value;
                                        }
                                    }
                                    else
                                    {
                                        dr[iCol - 1] = cell.Value;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    string.Format("Error: {0}\n{1}", ex.Message, ex.StackTrace);
                                }
                            }

                            try
                            {
                                dt.Rows.Add(dr);
                            }
                            catch (Exception ex)
                            {
                                BPSUtilities.WriteLog("Error: {0}\n{1}", ex.Message, ex.StackTrace);
                            }

                            iRow++;
                            // BPSUtilities.WriteLog("{0}: {1}", dt.TableName, iRow);
                        } // for each row

                        try
                        {
                            // BPSUtilities.WriteLog("Read {0} rows.", dt.Rows.Count);
                            ds.Tables.Add(dt);
                        }
                        catch (Exception ex)
                        {
                            BPSUtilities.WriteLog("Error: {0}\n{1}", ex.Message, ex.StackTrace);
                        }
                    }
                }
                catch (Exception ex)
                {
                    BPSUtilities.WriteLog("Error: {0}\n{1}", ex.Message, ex.StackTrace);
                }
            }
            else
            {
                BPSUtilities.WriteLog("Workbook '{0}' not found", sWBName);
            }

            return ds;
        }

        private static System.Data.DataSet ReadData(string sWBName, bool bSkipAuditTrail)
        {
            System.Data.DataSet ds = new System.Data.DataSet();

            if (System.IO.File.Exists(sWBName))
            {
                try
                {
                    try
                    {
                        UriParser.Register(new PWUriParser(GenericUriParserOptions.AllowEmptyAuthority |
                            GenericUriParserOptions.DontCompressPath | GenericUriParserOptions.DontConvertPathBackslashes |
                            GenericUriParserOptions.DontUnescapePathDotsAndSlashes | GenericUriParserOptions.GenericAuthority |
                            GenericUriParserOptions.NoFragment), "pw", 5800);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine(string.Format("Error: {0}\n{1}", ex.Message, ex.StackTrace));
                    }

                    var wb = new XLWorkbook(sWBName);

                    foreach (var ws in wb.Worksheets)
                    {
                        if (ds.Tables.Contains(ws.Name))
                            continue;

                        if (bSkipAuditTrail && ws.Name.ToLower() == "audittrail")
                            continue;

                        DataTable dt = new DataTable(ws.Name);

                        BPSUtilities.WriteLog("Reading table '{0}'...", ws.Name);

                        int iIndex = 1;

                        while (!string.IsNullOrEmpty(ws.Row(1).Cell(iIndex).Value.ToString()))
                        {
                            string sColumnName = ws.Row(1).Cell(iIndex).Value.ToString();

                            try
                            {
                                dt.Columns.Add(new DataColumn(sColumnName, ws.Row(2).Cell(iIndex).Value.GetType()));
                            }
                            catch
                            {
                                try
                                {
                                    dt.Columns.Add(new DataColumn(sColumnName, Type.GetType("System.String")));
                                }
                                catch (Exception ex)
                                {
                                    BPSUtilities.WriteLog("Error: {0}\n{1}", ex.Message, ex.StackTrace);

                                    // logic will really be screwed up if column cant' get added
                                    return null;
                                }
                            }

                            iIndex++;
                        } // for each column

                        if (dt.Columns.Count == 0)
                            continue;

                        int iNumCols = dt.Columns.Count;

                        int iRow = 2;

                        while (!ws.Row(iRow).IsEmpty())
                        {
                            DataRow dr = dt.NewRow();

                            for (int iCol = 1; iCol <= iNumCols; iCol++)
                            {
                                try
                                {
                                    IXLCell cell = ws.Row(iRow).Cell(iCol);

                                    if (cell.Hyperlink != null)
                                    {
                                        if (cell.Hyperlink.ExternalAddress != null)
                                        {
                                            dr[iCol - 1] = cell.Hyperlink.ExternalAddress.OriginalString;
                                        }
                                        else
                                        {
                                            dr[iCol - 1] = cell.Value;
                                        }
                                    }
                                    else
                                    {
                                        dr[iCol - 1] = cell.Value;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    string.Format("Error: {0}\n{1}", ex.Message, ex.StackTrace);
                                }
                            }

                            try
                            {
                                dt.Rows.Add(dr);
                            }
                            catch (Exception ex)
                            {
                                BPSUtilities.WriteLog("Error: {0}\n{1}", ex.Message, ex.StackTrace);
                            }

                            iRow++;
                            // BPSUtilities.WriteLog("{0}: {1}", dt.TableName, iRow);
                        } // for each row

                        try
                        {
                            BPSUtilities.WriteLog("Read {0} rows.", dt.Rows.Count);
                            ds.Tables.Add(dt);
                        }
                        catch (Exception ex)
                        {
                            BPSUtilities.WriteLog("Error: {0}\n{1}", ex.Message, ex.StackTrace);
                        }
                    }
                }
                catch (Exception ex)
                {
                    BPSUtilities.WriteLog("Error: {0}\n{1}", ex.Message, ex.StackTrace);
                }
            }
            else
            {
                BPSUtilities.WriteLog("Workbook '{0}' not found", sWBName);
            }

            return ds;
        }

        public static System.Data.DataSet DataSetFromXSLX(string sWBName)
        {
            System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("en-US");
            System.Data.DataSet ds = ReadData(sWBName);

            return ds;
        }

        public static System.Data.DataSet DataSetFromXSLX(string sWBName, bool bSkipAuditTrail)
        {
            System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("en-US");
            System.Data.DataSet ds = ReadData(sWBName, bSkipAuditTrail);

            return ds;
        }


        public static bool DataSetToXLSX(DataSet ds, string sWBName)
        {
            bool bRetVal = false;
            try
            {
                System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("en-US");

                var wb = new XLWorkbook();

                try
                {
                    UriParser.Register(new PWUriParser(GenericUriParserOptions.AllowEmptyAuthority |
                        GenericUriParserOptions.DontCompressPath | GenericUriParserOptions.DontConvertPathBackslashes |
                        GenericUriParserOptions.DontUnescapePathDotsAndSlashes | GenericUriParserOptions.GenericAuthority |
                        GenericUriParserOptions.NoFragment), "pw", 5800);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(string.Format("Error: {0}\n{1}", ex.Message, ex.StackTrace));
                }

                foreach (DataTable dt in ds.Tables)
                {
                    if (dt.TableName.Length > 30)
                        dt.TableName = dt.TableName.Substring(0, 30);

                    var ws = wb.Worksheets.Add(dt.TableName);

                    // ws.Cell("A1").Value = dt.TableName;

                    int iColumnIndex = 1;
                    int iRowIndex = 1;

                    SortedList<string, int> slColumnsToColumnIndices = new SortedList<string, int>(StringComparer.CurrentCultureIgnoreCase);

                    foreach (DataColumn dc in dt.Columns)
                    {
                        if (!slColumnsToColumnIndices.ContainsKey(dc.ColumnName))
                            slColumnsToColumnIndices.Add(dc.ColumnName, iColumnIndex);
                        ws.Cell(iRowIndex, iColumnIndex++).Value = dc.ColumnName;
                    }

                    iColumnIndex = 1;
                    iRowIndex = 2;

                    foreach (DataRow dr in dt.Rows)
                    {
                        foreach (DataColumn dc in dt.Columns)
                        {
                            string sColumnValue = dr[dc.ColumnName].ToString();

                            if (sColumnValue.StartsWith("pw://") || sColumnValue.StartsWith("pw:\\\\") ||
                               sColumnValue.StartsWith("http://") || sColumnValue.StartsWith("http:\\\\") ||
                               sColumnValue.StartsWith("https://") || sColumnValue.StartsWith("https:\\\\") ||
                               sColumnValue.StartsWith("file://") || sColumnValue.StartsWith("file:\\\\") ||
                               sColumnValue.StartsWith("ftp://") || sColumnValue.StartsWith("ftp:\\\\"))
                            {
                                if (sColumnValue.Contains("|"))
                                {
                                    string[] sParts = sColumnValue.Split("|".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);

                                    if (sParts.Length > 0)
                                    {
                                        Uri hyperlink = new Uri(Uri.EscapeUriString(sParts[0]));
                                        // Uri hyperlink = new Uri("http://www.bentley.com");

                                        if (sParts.Length == 2)
                                        {
                                            ws.Cell(iRowIndex, iColumnIndex).Value = sParts[1];
                                            ws.Cell(iRowIndex, iColumnIndex++).Hyperlink = new XLHyperlink(hyperlink);
                                        }
                                        else
                                        {
                                            ws.Cell(iRowIndex, iColumnIndex).Value = sParts[0];
                                            ws.Cell(iRowIndex, iColumnIndex++).Hyperlink = new XLHyperlink(hyperlink);
                                        }
                                    }
                                }
                                else
                                {
                                    Uri hyperlink = new Uri(Uri.EscapeUriString(sColumnValue));
                                    // Uri hyperlink = new Uri("http://www.bentley.com");

                                    ws.Cell(iRowIndex, iColumnIndex).Value = sColumnValue;
                                    ws.Cell(iRowIndex, iColumnIndex++).Hyperlink = new XLHyperlink(hyperlink);
                                }
                            }
                            else
                            {
                                ws.Cell(iRowIndex, iColumnIndex++).Value = dr[dc.ColumnName];
                            }
                        }

                        iColumnIndex = 1;
                        iRowIndex++;
                    }

                    var rngTable = ws.Range(1, 1, iRowIndex - 1, dt.Columns.Count);

                    // headers
                    var rngHeaders = rngTable.Range(1, 1, 1, dt.Columns.Count); // The address is relative to rngTable (NOT the worksheet)
                    rngHeaders.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    rngHeaders.Style.Font.Bold = true;
                    rngHeaders.Style.Fill.BackgroundColor = XLColor.Aqua;

                    rngTable.Style.Border.BottomBorder = XLBorderStyleValues.Thin;

                    // title
                    //rngTable.Cell(1, 1).Style.Font.Bold = true;
                    //rngTable.Cell(1, 1).Style.Fill.BackgroundColor = XLColor.CornflowerBlue;
                    //rngTable.Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                    //rngTable.Row(1).Merge();

                    //Add a thick outside border
                    rngTable.Style.Border.OutsideBorder = XLBorderStyleValues.Thick;

                    // You can also specify the border for each side with:
                    // rngTable.FirstColumn().Style.Border.LeftBorder = XLBorderStyleValues.Thick;
                    // rngTable.LastColumn().Style.Border.RightBorder = XLBorderStyleValues.Thick;
                    // rngTable.FirstRow().Style.Border.TopBorder = XLBorderStyleValues.Thick;
                    // rngTable.LastRow().Style.Border.BottomBorder = XLBorderStyleValues.Thick;

                    ws.Columns(1, dt.Columns.Count).AdjustToContents();


                } // for each table

                // string sWBName = Guid.NewGuid().ToString() + ".xlsx";

                wb.SaveAs(sWBName);

                bRetVal = true;

                // BPSUtilities.WriteLog("Wrote " + sWBName);
            }
            catch (Exception ex)
            {
                BPSUtilities.WriteLog("Error: {0}\n{1}", ex.Message, ex.StackTrace);
            }

            return bRetVal;
        }
        public static bool DataSetToXLSXFast(DataSet ds, string sWBName)
        {
            bool bRetVal = false;
            try
            {
                System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("en-US");

                var wb = new XLWorkbook();

                try
                {
                    // might not really need this since no hyperlinks anyway...
                    UriParser.Register(new PWUriParser(GenericUriParserOptions.AllowEmptyAuthority |
                        GenericUriParserOptions.DontCompressPath | GenericUriParserOptions.DontConvertPathBackslashes |
                        GenericUriParserOptions.DontUnescapePathDotsAndSlashes | GenericUriParserOptions.GenericAuthority |
                        GenericUriParserOptions.NoFragment), "pw", 5800);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(string.Format("Error: {0}\n{1}", ex.Message, ex.StackTrace));
                }

                foreach (DataTable dt in ds.Tables)
                {
                    if (dt.TableName.Length > 30)
                        dt.TableName = dt.TableName.Substring(0, 30);
                }

                wb.Worksheets.Add(ds);

                wb.SaveAs(sWBName);

                bRetVal = true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error: {ex.Message}\n{ex.StackTrace}");
            }

            return bRetVal;
        }

        private static System.Data.DataSet ReadDataReallyFast(string sWBName)
        {
            System.Data.DataSet ds = new System.Data.DataSet();

            if (System.IO.File.Exists(sWBName))
            {
                try
                {
                    using (var stream = System.IO.File.OpenRead(sWBName))
                    {
                        using (var reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            ds = reader.AsDataSet(new ExcelDataSetConfiguration()
                            {
                                ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                                {
                                    UseHeaderRow = true
                                }
                            });
                        }
                    }
                }
                catch (Exception ex)
                {
                    BPSUtilities.WriteLog("Error: {0}\n{1}", ex.Message, ex.StackTrace);
                }
            }
            else
            {
                BPSUtilities.WriteLog("Workbook '{0}' not found", sWBName);
            }

            return ds;
        }

        public static System.Data.DataSet DataSetFromXLSXFast(string sWBName)
        {
            System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("en-US");
            System.Data.DataSet ds = ReadDataReallyFast(sWBName);

            return ds;
        }


    }
}
