using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Xml;
using Bentley.GeometryNET;
using Bentley.DgnPlatformNET;
using Bentley.DgnPlatformNET.Elements;
using Bentley.MstnPlatformNET;
using Bentley.Interop.MicroStationDGN;
using System.IO;

namespace PolyhedraCE
{
    public class CreatePolyhedronX64 : DgnPrimitiveTool // Bentley.Interop.MicroStationDGN.IPrimitiveCommandEvents
    {
        double uorPerMaster()
        {
            Bentley.DgnPlatformNET.DgnModel oModel = Session.Instance.GetActiveDgnModel();
            Bentley.DgnPlatformNET.ModelInfo myModelInfo;
            myModelInfo = oModel.GetModelInfo();
            return myModelInfo.UorPerMaster;
        }
        ///////////////////////////////////////////////////////////////////////////////////
        double masterPerUOR()
        {
            Bentley.DgnPlatformNET.DgnModel oModel = Session.Instance.GetActiveDgnModel();
            Bentley.DgnPlatformNET.ModelInfo myModelInfo;
            myModelInfo = oModel.GetModelInfo();
            return 1 / myModelInfo.UorPerMaster;
        }

        SortedList<string, uint> slColorList = new SortedList<string, uint>(StringComparer.CurrentCultureIgnoreCase);

        // should create on 0,0,0 center
        private Bentley.DgnPlatformNET.Elements.Element CreatePolyhedronCell(string sPolyhedronName,
            List<DPoint3d> points, List<List<int>> faces, List<DPoint3d> colors)
        {
            try
            {
                List<Bentley.DgnPlatformNET.Elements.Element> listElements = new List<Bentley.DgnPlatformNET.Elements.Element>();

                if (points.Count > 0 && faces.Count > 0)
                {
                    int iFaceIndex = 0;

                    foreach (List<int> listVertices in faces)
                    {
                        List<DPoint3d> listPoints = new List<DPoint3d>();

                        foreach (int iVertexIndex in listVertices)
                        {
                            DPoint3d pt = points[iVertexIndex];
                            pt.ScaleInPlace(uorPerMaster());
                            listPoints.Add(pt);
                        }

                        DPoint3d color = colors.ToArray()[iFaceIndex++];

                        Bentley.DgnPlatformNET.Elements.ShapeElement shapeElement =
                            new Bentley.DgnPlatformNET.Elements.ShapeElement(Session.Instance.GetActiveDgnModel(), null, listPoints.ToArray());

                        DVector3d normal = new DVector3d();
                        DPoint3d somePoint = new DPoint3d();
                        DPoint3d defNormal = listPoints.ToArray()[0];

                        shapeElement.IsPlanar(out normal, out somePoint, ref defNormal);

                        uint iColorIndex = (uint)listPoints.Count - 1;

                        ElementPropertiesSetter pSetter = new ElementPropertiesSetter();

                        pSetter.SetColor(iColorIndex);
                        pSetter.SetFillColor(iColorIndex);
                        shapeElement.AddSolidFill(iColorIndex, false);

                        pSetter.Apply(shapeElement);

                        listElements.Add(shapeElement);
                    }
                }

                if (listElements.Count > 0)
                {
                    DMatrix3d rotation = new DMatrix3d(1, 0, 0, 0, 1, 0, 0, 0, 1);   // Identity
                    Bentley.DgnPlatformNET.Elements.CellHeaderElement cellHeaderElement = new CellHeaderElement(Session.Instance.GetActiveDgnModel(), sPolyhedronName,
                        new DPoint3d(0, 0, 0), rotation, listElements);

                    return cellHeaderElement;
                }
            }
            catch (Exception ex)
            {
                BPSUtilities.WriteLog($"CreatePolyhedronCell {ex.Message}\n{ex.StackTrace}");
            }


            return null;
        }

        public Bentley.DgnPlatformNET.Elements.Element SetPolyhedron(string sPolyhedronName, string sXMLFileName)
        {
            BPSUtilities.WriteLog($"Creating '{sPolyhedronName}' from '{sXMLFileName}'");

            List<DPoint3d> points = new List<DPoint3d>();
            List<List<int>> faceVertices = new List<List<int>>();
            List<DPoint3d> rgbColors = new List<DPoint3d>();

            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(sXMLFileName);

                XmlNode polyhedronNode = null;

                XmlNode root = xmlDoc.DocumentElement;

                if (sPolyhedronName.Length > 0)
                {
                    polyhedronNode =
                        root.SelectSingleNode(String.Format("polyhedron[@name = \"{0}\"]", sPolyhedronName));
                }

                if (polyhedronNode != null)
                {
                    BPSUtilities.WriteLog($"Vertices: {polyhedronNode.Attributes["vertices"].Value}");

                    int iNumVertices = Int32.Parse(polyhedronNode.Attributes["vertices"].Value);

                    BPSUtilities.WriteLog($"Faces: {polyhedronNode.Attributes["faces"].Value}");

                    int iNumFaces = Int32.Parse(polyhedronNode.Attributes["faces"].Value);

                    XmlNode vertices = polyhedronNode.SelectSingleNode("vertices");

                    if (vertices != null)
                    {
                        foreach (XmlNode vertex in vertices.ChildNodes)
                        {
                            // BPSUtilities.WriteLog(vertex.InnerText);
                            string[] split = vertex.InnerText.Trim().Split(new Char[] { ' ', ',', ':' }, StringSplitOptions.RemoveEmptyEntries);

                            string[] values = new string[3];

                            int iValIndex = 0;

                            foreach (string s in split)
                            {
                                if (s.Length > 0 && iValIndex < 3)
                                    values[iValIndex++] = s;
                            }

                            try
                            {
                                points.Add(new DPoint3d(Double.Parse(values[0]), Double.Parse(values[1]), Double.Parse(values[2])));

                            }
                            catch (Exception ex)
                            {
                                BPSUtilities.WriteLog($"X: {values[0]},Y: {values[1]},Z: {values[2]}");

                                BPSUtilities.WriteLog($"Error parsing point, {ex.Message}");
                            }
                        }
                    }
                    else
                    {
                        BPSUtilities.WriteLog("Vertices node not found.");
                    }

                    XmlNode faces = polyhedronNode.SelectSingleNode("faces");

                    int iFaceCount = 0;

                    if (faces != null)
                    {
                        foreach (XmlNode face in faces.ChildNodes)
                        {
                            // BPSUtilities.WriteLog($"vertexcount: {face.Attributes["vertexcount"].Value}");

                            int iVtxCount = Int32.Parse(face.Attributes["vertexcount"].Value);

                            string[] split = face.InnerText.Trim().Split(new Char[] { ' ', ',', ':' }, StringSplitOptions.RemoveEmptyEntries);

                            List<int> listVertices = new List<int>();

                            foreach (string s in split)
                            {
                                try
                                {
                                    listVertices.Add(int.Parse(s));
                                }
                                catch (Exception ex)
                                {
                                    BPSUtilities.WriteLog($"Error parsing '{s}' {ex.Message}");
                                    BPSUtilities.WriteLog($"Face list '{face.InnerText}'");
                                }
                            }

                            faceVertices.Add(listVertices);

                            string[] colorSplit =
                                face.Attributes["RGB"].Value.Trim().Split(new Char[] { ' ', ',', ':' }, StringSplitOptions.RemoveEmptyEntries);

                            if (colorSplit.Length == 3)
                            {
                                try
                                {
                                    rgbColors.Add(new DPoint3d(Double.Parse(colorSplit[0]),
                                        Double.Parse(colorSplit[1]),
                                        Double.Parse(colorSplit[2])));
                                }
                                catch (Exception ex)
                                {
                                    BPSUtilities.WriteLog($"Error parsing RGB, {ex.Message}");
                                    BPSUtilities.WriteLog($"R: {colorSplit[0]}, G: {colorSplit[1]}, B: {colorSplit[2]}");
                                }
                            }

                            iFaceCount++;
                        } // foreach face
                    }
                } // found a node
                else
                {
                    BPSUtilities.WriteLog("Node not found.");
                    return null;
                }
            }
            catch (Exception ex)
            {
                BPSUtilities.WriteLog($"{ex.Message}\n{ex.StackTrace}");
                return null;
            }

            if (points.Count > 0 && faceVertices.Count > 0 && rgbColors.Count > 0)
                return CreatePolyhedronCell(sPolyhedronName, points, faceVertices, rgbColors);

            return null;
        }

        public void Keyin(string Keyin)
        {
        }

        public void Cleanup()
        {
        }

        protected override void OnRestartTool()
        {
            try
            {
                BPSUtilities.WriteLog("OnRestartTool");
                // throw new NotImplementedException();
                InstallNewInstance(PolyhedronName);
                m_haveFirstPoint = false;
                base.BeginDynamics();
            }
            catch (Exception ex)
            {
                BPSUtilities.WriteLog($"OnRestartTool {ex.Message}\n{ex.StackTrace}");
            }
        }

        public void StopDynamics()
        {
            try
            {
                if (this.DynamicsStarted)
                    base.EndDynamics();
            }
            catch (Exception ex)
            {
                BPSUtilities.WriteLog($"EndDynamics {ex.Message}\n{ex.StackTrace}");
            }
        }


        public void StartDynamics()
        {
            try
            {
                if (!this.DynamicsStarted)
                    base.BeginDynamics();
            }
            catch (Exception ex)
            {
                BPSUtilities.WriteLog($"StartDynamics {ex.Message}\n{ex.StackTrace}");
            }
        }

        protected override void OnPostInstall()
        {
            try
            {
                BPSUtilities.WriteLog("OnPostInstall");

                NotificationManager.OutputPrompt("Enter data point to place polyhedron. Reset to exit.");
                // AccuSnap.SnapEnabled = true;
                // AccuSnap.LocateEnabled = true;
                base.OnPostInstall();

                m_haveFirstPoint = false;
                base.BeginDynamics();

                BPSUtilities.WriteLog($"OnPostInstall Dynamics {(this.DynamicsStarted ? "is" : "is not")} started.");

            }
            catch (Exception ex)
            {
                BPSUtilities.WriteLog($"OnPostInstall {ex.Message}\n{ex.StackTrace}");
            }
        }

        protected override bool OnResetButton(DgnButtonEvent ev)
        {
            ExitTool();
            return true;
        }

        protected override bool OnDataButton(DgnButtonEvent ev)
        {
            try
            {
                if (m_haveFirstPoint)
                {
                    // we're going to write the element

                    base.EndDynamics();

                    // now to work on rotation and scale
                    DPoint3d translation = m_firstPoint;

                    DMatrix3d viewMatrix = ev.Viewport.GetRotation();

                    DPoint3d[] blkPts = new DPoint3d[4];

                    // rotate points to orthogonal
                    blkPts[0] = new DPoint3d(viewMatrix.Multiply(new DVector3d(m_firstPoint)));

                    blkPts[2] = new DPoint3d(viewMatrix.Multiply(new DVector3d(ev.Point)));

                    blkPts[2].Z = blkPts[0].Z;

                    DPoint3d ext = DPoint3d.Subtract(blkPts[2], blkPts[0]);

                    DVector3d angleVec = new DVector3d(ext);

                    double dScale = Math.Max(0.1, Math.Abs(ext.Magnitude / uorPerMaster()));

                    //RedrawElems redrawElems = new RedrawElems();
                    //redrawElems.SetDynamicsViewsFromActiveViewSet(Bentley.MstnPlatformNET.Session.GetActiveViewport());
                    //redrawElems.DrawMode = DgnDrawMode.Normal;
                    //redrawElems.DrawPurpose = DrawPurpose.ForceRedraw;

                    try
                    {
                        DMatrix3d invertedViewportRotation = new DMatrix3d(1, 0, 0, 0, 1, 0, 0, 0, 1);   // Identity

                        if (ev.Viewport.GetRotation().TryInvert(out invertedViewportRotation))
                        {
                            DMatrix3d cellRotation = new DMatrix3d(1, 0, 0, 0, 1, 0, 0, 0, 1);   // Identity

                            List<Bentley.DgnPlatformNET.Elements.Element> listChildElements = new List<Bentley.DgnPlatformNET.Elements.Element>();

                            foreach (Bentley.DgnPlatformNET.Elements.Element child in PolyhedraCE.StaticElement.GetChildren())
                                listChildElements.Add(child);

                            Bentley.DgnPlatformNET.Elements.CellHeaderElement copiedElement = new CellHeaderElement(Session.Instance.GetActiveDgnModel(), PolyhedronName,
                                new DPoint3d(0, 0, 0), cellRotation, listChildElements);

                            if (copiedElement != null)
                            {
                                if (copiedElement.IsValid)
                                {
                                    invertedViewportRotation.ScaleInPlace(dScale);

                                    DMatrix3d scaleRotateInViewRotateAroundPoint = DMatrix3d.Multiply(invertedViewportRotation, DMatrix3d.Rotation(2, angleVec.AngleXY));

                                    DTransform3d translateAndRotatetoUpInViewAndScale =
                                        DTransform3d.FromMatrixAndTranslation(scaleRotateInViewRotateAroundPoint, translation);

                                    TransformInfo transformInfo2 = new TransformInfo(translateAndRotatetoUpInViewAndScale);

                                    copiedElement.ApplyTransform(transformInfo2);

                                    copiedElement.AddToModel();

                                    BPSUtilities.WriteLog("Added element to model");

                                    RedrawElems redrawElems = new RedrawElems();
                                    redrawElems.SetDynamicsViewsFromActiveViewSet(ev.Viewport);
                                    redrawElems.DrawMode = DgnDrawMode.Normal;
                                    redrawElems.DrawPurpose = DrawPurpose.ForceRedraw;

                                    redrawElems.DoRedraw(copiedElement);
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        BPSUtilities.WriteLog($"OnDataBurron {ex.Message}\n{ex.StackTrace}");
                    }

                    m_haveFirstPoint = false;

                    m_firstPoint = new DPoint3d(0, 0, 0);

                    base.BeginDynamics();

                    return true;
                }

                m_haveFirstPoint = true;

                m_firstPoint = ev.Point;

                base.BeginDynamics();
            }
            catch (Exception ex)
            {
                BPSUtilities.WriteLog($"OnDataButton {ex.Message}\n{ex.StackTrace}");
            }

            return true;
        }

        public DTransform3d m_translateAndRotate = DTransform3d.Identity;

        public string PolyhedronName { get; set; }

        public CreatePolyhedronX64(int toolId, int prompt, string sPolyhedronName) : base(toolId, prompt)
        {
            try
            {
                // m_points = new List<DPoint3d>();

                PolyhedronName = sPolyhedronName;

                // string sXMLFileName = string.Empty;

                BPSUtilities.WriteLog($"CreatePolyhedronX64 Initialized '{PolyhedronName}'");

                string sOutputXML2 = Path.Combine(Path.GetTempPath(), $"{BPSUtilities.GetARandomString(8, BPSUtilities.LOWER_CASE)}.xml");

                if (PolyhedraCE.ListOfPolyhedra.Count == 0)
                {
                    string sOutputXML = Path.Combine(Path.GetTempPath(), $"{BPSUtilities.GetARandomString(8, BPSUtilities.LOWER_CASE)}.xml");

                    try
                    {
                        System.Reflection.Assembly thisExe = System.Reflection.Assembly.GetExecutingAssembly();

                        foreach (string sResName in thisExe.GetManifestResourceNames())
                        {
                            if (sResName.ToLower().EndsWith(".polyhedra.xml"))
                            {
                                using (var resourceStream = thisExe.GetManifestResourceStream(sResName))
                                {
                                    resourceStream.CopyTo(new System.IO.FileStream(sOutputXML, FileMode.Create));
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        BPSUtilities.WriteLog($"{ex.Message}\n{ex.StackTrace}");
                    }

                    File.Copy(sOutputXML, sOutputXML2);

                    try
                    {
                        XmlDocument xmlDoc = new XmlDocument();
                        xmlDoc.Load(sOutputXML2);

                        XmlNode root = xmlDoc.DocumentElement;

                        // need form to list available polyhedra

                        XmlNodeList nodeList = root.SelectNodes(".//polyhedron");

                        List<string> listPolyhedrons = new List<string>();

                        foreach (XmlNode node in nodeList)
                        {
                            string sPolyName = node.Attributes["name"].Value;
                            listPolyhedrons.Add(sPolyName);
                        }

                        foreach (string sPolyName in listPolyhedrons)
                        {
                            try
                            {
                                BPSUtilities.WriteLog($"Processing '{sPolyName}'...");

                                Bentley.DgnPlatformNET.Elements.Element elm = SetPolyhedron(sPolyName, sOutputXML2);

                                if (elm != null)
                                {
                                    PolyhedraCE.ListOfPolyhedra.AddWithCheck(sPolyName, elm);
                                    // elm.AddToModel();
                                    BPSUtilities.WriteLog($"Added '{sPolyName}'");
                                }
                                else
                                {
                                    BPSUtilities.WriteLog($"Error creating '{sPolyName}'");
                                }
                            }
                            catch (Exception ex)
                            {
                                BPSUtilities.WriteLog($"{ex.Message}\n{ex.StackTrace}");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        BPSUtilities.WriteLog($"{ex.Message}\n{ex.StackTrace}");
                    }
                }
                else
                {
                }

                BPSUtilities.WriteLog($"List contains {PolyhedraCE.ListOfPolyhedra.Count} entries.");

                BPSUtilities.WriteLog($"CreatePolyhedronX64: Dynamics {(this.DynamicsStarted ? "is" : "is not")} started.");
            }
            catch (Exception ex)
            {
                BPSUtilities.WriteLog($"CreatePolyhedronX64 {ex.Message}\n{ex.StackTrace}");
            }
        }

        public static void InstallNewInstance(string sPolyhedronName)
        {
            try
            {
                BPSUtilities.WriteLog("InstallNewInstance");
                CreatePolyhedronX64 createPolyhedronTool = new CreatePolyhedronX64(0, 0, sPolyhedronName);
                createPolyhedronTool.InstallTool();
            }
            catch (Exception ex)
            {
                BPSUtilities.WriteLog($"InstallNewInstance {ex.Message}\n{ex.StackTrace}");
            }
        }

        // private static PolyhedaListForm m_polyhedraListForm = null;

        protected override void ExitTool()
        {
            try
            {
                BPSUtilities.WriteLog("ExitTool");
                PolyhedraCE.StaticElement.Dispose();
                PolyhedraCE.StaticElement = null;

                foreach (Bentley.DgnPlatformNET.Elements.Element elm in PolyhedraCE.ListOfPolyhedra.Values)
                {
                    elm.Dispose();
                }

                PolyhedraCE.ListOfPolyhedra.Clear();

                base.ExitTool();
            }
            catch (Exception ex)
            {
                BPSUtilities.WriteLog($"ExitTool {ex.Message}\n{ex.StackTrace}");
            }
        }

        private DPoint3d m_firstPoint = new DPoint3d(0, 0, 0);
        private bool m_haveFirstPoint = false;

        protected override void OnDynamicFrame(Bentley.DgnPlatformNET.DgnButtonEvent ev)
        {
            try
            {
                if (PolyhedraCE.StaticElement == null)
                {
                    if (!string.IsNullOrEmpty(PolyhedronName))
                    {
                        if (PolyhedraCE.ListOfPolyhedra.ContainsKey(PolyhedronName))
                        {
                            BPSUtilities.WriteLog($"Found cell for '{PolyhedronName}'");
                            PolyhedraCE.StaticElement = PolyhedraCE.ListOfPolyhedra[PolyhedronName];
                        }
                        else
                        {
                            m_haveFirstPoint = false;
                            BPSUtilities.WriteLog($"Cell for '{PolyhedronName}' not found.");
                            this.EndDynamics();
                        }
                    }
                    else
                    {
                        BPSUtilities.WriteLog("Polyhedron Name not set.");
                        m_haveFirstPoint = false;
                        this.EndDynamics();
                    }
                }

                try
                {
                    Bentley.DgnPlatformNET.Elements.CellHeaderElement cellElem = (Bentley.DgnPlatformNET.Elements.CellHeaderElement)(PolyhedraCE.StaticElement);
                    if (cellElem.CellName != PolyhedronName)
                    {
                        BPSUtilities.WriteLog($"Current cell is {cellElem.CellName}");

                        if (!string.IsNullOrEmpty(PolyhedronName))
                        {
                            if (PolyhedraCE.ListOfPolyhedra.ContainsKey(PolyhedronName))
                            {
                                BPSUtilities.WriteLog($"Found cell for '{PolyhedronName}'");

                                if (PolyhedraCE.StaticElement != null)
                                    PolyhedraCE.StaticElement.Dispose();

                                PolyhedraCE.StaticElement = PolyhedraCE.ListOfPolyhedra[PolyhedronName];
                            }
                            else
                            {
                                m_haveFirstPoint = false;
                                BPSUtilities.WriteLog($"Cell for '{PolyhedronName}' not found.");
                                this.EndDynamics();
                            }
                        }
                        else
                        {
                            BPSUtilities.WriteLog("Polyhedron Name not set.");
                            m_haveFirstPoint = false;
                            this.EndDynamics();
                        }
                    }
                }
                catch (Exception ex)
                {
                    BPSUtilities.WriteLog($"Error casting cell {ex.Message}");
                }

                if (PolyhedraCE.StaticElement != null)
                {
                    RedrawElems redrawElems = new RedrawElems();
                    redrawElems.SetDynamicsViewsFromActiveViewSet(Bentley.MstnPlatformNET.Session.GetActiveViewport());
                    redrawElems.DrawMode = DgnDrawMode.TempDraw;
                    redrawElems.DrawPurpose = DrawPurpose.Dynamics;

                    try
                    {
                        DMatrix3d invertedViewportRotation = new DMatrix3d(1, 0, 0, 0, 1, 0, 0, 0, 1);   // Identity

                        if (ev.Viewport.GetRotation().TryInvert(out invertedViewportRotation))
                        {
                            DMatrix3d cellRotation = new DMatrix3d(1, 0, 0, 0, 1, 0, 0, 0, 1);   // Identity

                            List<Bentley.DgnPlatformNET.Elements.Element> listChildElements = new List<Bentley.DgnPlatformNET.Elements.Element>();

                            foreach (Bentley.DgnPlatformNET.Elements.Element child in PolyhedraCE.StaticElement.GetChildren())
                                listChildElements.Add(child);

                            Bentley.DgnPlatformNET.Elements.CellHeaderElement copiedElement = new CellHeaderElement(Session.Instance.GetActiveDgnModel(), PolyhedronName,
                                new DPoint3d(0, 0, 0), cellRotation, listChildElements);

                            if (copiedElement != null)
                            {
                                if (copiedElement.IsValid)
                                {
                                    if (!m_haveFirstPoint)
                                    {
                                        DPoint3d translation = ev.Point;

                                        DMatrix3d viewMatrix = ev.Viewport.GetRotation();

                                        DPoint3d[] blkPts = new DPoint3d[4];

                                        DPoint3d low = DPoint3d.Zero, high = DPoint3d.Zero;

                                        ev.Viewport.GetViewCorners(out low, out high);

                                        // rotate points to orthogonal
                                        blkPts[0] = new DPoint3d(viewMatrix.Multiply(new DVector3d(low)));

                                        blkPts[2] = new DPoint3d(viewMatrix.Multiply(new DVector3d(high)));

                                        blkPts[2].Z = blkPts[0].Z;

                                        DPoint3d ext = DPoint3d.Subtract(blkPts[2], blkPts[0]);

                                        double dScale = Math.Max(0.15 * Math.Abs(ext.Magnitude / uorPerMaster()), 1.0);

                                        DPoint3d dPtScale = ev.Viewport.GetScale();

                                        // dScale = dPtScale.Magnitude;

                                        // BPSUtilities.WriteLog($"Scale: {dScale}, View: {ev.ViewNumber}");

                                        invertedViewportRotation.ScaleInPlace(dScale);

                                        // works
                                        DTransform3d translateAndRotateToUpInViewWithViewBasedScale =
                                            DTransform3d.FromMatrixAndTranslation(invertedViewportRotation, translation);

                                        TransformInfo transformInfo = new TransformInfo(translateAndRotateToUpInViewWithViewBasedScale);

                                        // works
                                        copiedElement.ApplyTransform(transformInfo);
                                    }
                                    else
                                    {
                                        // now to work on rotation and scale
                                        DPoint3d translation = m_firstPoint;

                                        DMatrix3d viewMatrix = ev.Viewport.GetRotation();

                                        DPoint3d[] blkPts = new DPoint3d[4];

                                        // rotate points to orthogonal
                                        blkPts[0] = new DPoint3d(viewMatrix.Multiply(new DVector3d(m_firstPoint)));

                                        blkPts[2] = new DPoint3d(viewMatrix.Multiply(new DVector3d(ev.Point)));

                                        blkPts[2].Z = blkPts[0].Z;

                                        DPoint3d ext = DPoint3d.Subtract(blkPts[2], blkPts[0]);

                                        DVector3d angleVec = new DVector3d(ext);

                                        double dScale = Math.Max(0.1, Math.Abs(ext.Magnitude / uorPerMaster()));

                                        invertedViewportRotation.ScaleInPlace(dScale);

                                        DMatrix3d scaleRotateInViewRotateAroundPoint = DMatrix3d.Multiply(invertedViewportRotation, DMatrix3d.Rotation(2, angleVec.AngleXY));

                                        DTransform3d translateAndRotatetoUpInViewAndScale =
                                            DTransform3d.FromMatrixAndTranslation(scaleRotateInViewRotateAroundPoint, translation);

                                        TransformInfo transformInfo2 = new TransformInfo(translateAndRotatetoUpInViewAndScale);

                                        copiedElement.ApplyTransform(transformInfo2);
                                    }

                                    redrawElems.DoRedraw(copiedElement);
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        BPSUtilities.WriteLog($"{ex.Message}\n{ex.StackTrace}");
                    }
                    //    if (!m_haveFirstPoint)
                    //{
                    //    BPSUtilities.WriteLog("Don't have first point.");
                    //    // move it around on cursor
                    //    redrawElems.DoRedraw(GetTransformedElement(m_element, ev));
                    //}
                    //else
                    //{
                    //    BPSUtilities.WriteLog("Do have first point.");
                    //    redrawElems.DoRedraw(GetTransformedElement(m_element, ev));
                    //}
                }
                else
                {
                    BPSUtilities.WriteLog("Element is null");
                    this.EndDynamics();
                }

            }
            catch (Exception ex)
            {
                BPSUtilities.WriteLog($"OnDynamicFrame: {ex.Message}\n{ex.StackTrace}");
            }
        }
    }
}
