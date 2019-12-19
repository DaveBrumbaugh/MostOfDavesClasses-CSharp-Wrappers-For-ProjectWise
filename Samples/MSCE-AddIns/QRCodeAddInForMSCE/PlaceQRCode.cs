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

namespace QRCodeAddInForMSCE
{
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

    public class PlaceQRCode : DgnPrimitiveTool // Bentley.Interop.MicroStationDGN.IPrimitiveCommandEvents
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
                BPSUtilities.WriteLog("PlaceQRCode OnRestartTool");
                // throw new NotImplementedException();
                InstallNewInstance("");
                m_haveFirstPoint = false;
                // base.BeginDynamics();
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

        private static Bentley.DgnPlatformNET.Elements.Element CodeElement { get; set; }

        protected override void OnPostInstall()
        {
            try
            {
                BPSUtilities.WriteLog("PlaceQRCode OnPostInstall");

                NotificationManager.OutputPrompt("Enter data point for corner of QR Code. Reset to exit.");
                // AccuSnap.SnapEnabled = true;
                // AccuSnap.LocateEnabled = true;
                base.OnPostInstall();

                m_haveFirstPoint = false;
                // base.BeginDynamics();
                m_firstPoint = DPoint3d.Zero;

                if (!string.IsNullOrEmpty(LongURL))
                {
                    QRCodeEncoderLibrary.QREncoder enc = QRCodeImageFunctions.GetQREncoderObject(LongURL);

                    BPSUtilities.WriteLog($"QRCodeDimension: {enc.QRCodeDimension}");

                    if (enc.QRCodeDimension > 0)
                    {
                        DPoint3d[] blkPts = new DPoint3d[4];

                        blkPts[0] = blkPts[1] = blkPts[2] = blkPts[3] = DPoint3d.Zero;

                        blkPts[0].X = enc.QuietZone;
                        blkPts[0].Y = enc.QuietZone;

                        blkPts[1].X = enc.QuietZone + 1.0;
                        blkPts[1].Y = enc.QuietZone;
                        blkPts[2].X = enc.QuietZone + 1.0;
                        blkPts[2].Y = enc.QuietZone + 1.0;
                        blkPts[3].X = enc.QuietZone;
                        blkPts[3].Y = enc.QuietZone + 1.0;

                        DPoint3d[] blkPts2 = new DPoint3d[4];

                        DVector3d yVec = new DVector3d(0, 1, 0);
                        DVector3d xVec = new DVector3d(1, 0, 0);

                        DMatrix3d cellRotation = new DMatrix3d(1, 0, 0, 0, 1, 0, 0, 0, 1);   // Identity

                        List<Bentley.DgnPlatformNET.Elements.Element> listChildElements = new List<Bentley.DgnPlatformNET.Elements.Element>();

                        // bool[,] pixels = enc.ConvertQRCodeMatrixToPixels();

                        for (int i = 0; i < enc.QRCodeDimension; i++)
                        {
                            // reset row
                            for (int index = 0; index < 4; index++)
                                blkPts2[index] = blkPts[index];

                            // move up
                            for (int index = 0; index < 4; index++)
                                blkPts2[index] = DPoint3d.Add(blkPts2[index], yVec, (double)i);

                            for (int j = 0; j < enc.QRCodeDimension; j++)
                            {
                                if (enc.QRCodeMatrix[i, j])
                                {
                                    Bentley.DgnPlatformNET.Elements.ShapeElement shapeElement =
                                        new Bentley.DgnPlatformNET.Elements.ShapeElement(Session.Instance.GetActiveDgnModel(), null, blkPts2);

                                    ElementPropertiesSetter pSetter = new ElementPropertiesSetter();

                                    pSetter.SetColor((uint)Bentley.MstnPlatformNET.Settings.GetActiveFillColor().Index);
                                    pSetter.SetFillColor((uint)Bentley.MstnPlatformNET.Settings.GetActiveFillColor().Index);
                                    shapeElement.AddSolidFill((uint)Bentley.MstnPlatformNET.Settings.GetActiveFillColor().Index, true);

                                    pSetter.Apply(shapeElement);

                                    if (shapeElement.IsValid)
                                        listChildElements.Add(shapeElement);
                                }

                                // move across
                                for (int index = 0; index < 4; index++)
                                    blkPts2[index] = DPoint3d.Add(blkPts2[index], xVec);
                            }
                        }

                        DPoint3d[] blkPts3 = new DPoint3d[5];

                        blkPts3[0] = blkPts3[1] = blkPts3[2] = blkPts3[3] = blkPts3[4] = DPoint3d.Zero;

                        blkPts3[1].X = 2 * enc.QuietZone + enc.QRCodeDimension;
                        blkPts3[2].X = 2 * enc.QuietZone + enc.QRCodeDimension;
                        blkPts3[2].Y = 2 * enc.QuietZone + enc.QRCodeDimension;
                        blkPts3[3].Y = 2 * enc.QuietZone + enc.QRCodeDimension;

                        Bentley.DgnPlatformNET.Elements.LineStringElement border = new LineStringElement(Session.Instance.GetActiveDgnModel(), null, blkPts3);

                        ElementPropertiesSetter pSetter2 = new ElementPropertiesSetter();

                        pSetter2.SetColor((uint)Bentley.MstnPlatformNET.Settings.GetActiveFillColor().Index);

                        pSetter2.Apply(border);

                        listChildElements.Add(border);

                        Bentley.DgnPlatformNET.Elements.CellHeaderElement codeElement = new CellHeaderElement(Session.Instance.GetActiveDgnModel(), "QRCode",
                            new DPoint3d(0, 0, 0), cellRotation, listChildElements);

                        // scale down to one...

                        DMatrix3d rotMatrix = DMatrix3d.Identity;

                        rotMatrix.ScaleInPlace((double) 1.0 / (double) enc.QRCodeDimension);

                        DTransform3d translateAndScale =
                            DTransform3d.FromMatrixAndTranslation(rotMatrix, DPoint3d.Zero);

                        TransformInfo transformInfo = new TransformInfo(translateAndScale);

                        codeElement.ApplyTransform(transformInfo);

                        CodeElement = codeElement;

                        // might want to tag with the URL...

                        if (CodeElement.IsValid)
                            base.BeginDynamics();
                    }
                }

                BPSUtilities.WriteLog($"PlaceQRCode OnPostInstall Dynamics {(this.DynamicsStarted ? "is" : "is not")} started.");
            }
            catch (Exception ex)
            {
                BPSUtilities.WriteLog($"PlaceQRCode OnPostInstall {ex.Message}\n{ex.StackTrace}");
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

                    if (CodeElement != null)
                    {
                        if (CodeElement.IsValid)
                        {
                            List<Bentley.DgnPlatformNET.Elements.Element> listChildElements = CodeElement.GetChildren().GetEnumerator().ConvertToList();

                            Bentley.DgnPlatformNET.Elements.CellHeaderElement codeElement = new CellHeaderElement(Session.Instance.GetActiveDgnModel(), "QRCode",
                                DPoint3d.Zero, DMatrix3d.Identity, listChildElements);

                            if (codeElement.IsValid)
                            {
                                DMatrix3d invertedViewportRotation = new DMatrix3d(1, 0, 0, 0, 1, 0, 0, 0, 1);   // Identity

                                if (ev.Viewport.GetRotation().TryInvert(out invertedViewportRotation))
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

                                    // double dScale = Math.Max(0.1, Math.Abs(ext.Magnitude / uorPerMaster()));

                                    invertedViewportRotation.ScaleInPlace(ext.Magnitude);

                                    DVector3d xVec = new DVector3d(1, 0, 0);

                                    DVector3d yVec = new DVector3d(0, 1, 0);

                                    Angle angle = Angle.Zero;

                                    //// UR and LR
                                    if (angleVec.AngleTo(xVec).Radians >= 0.0 && angleVec.AngleTo(xVec).Radians <= (Angle.PI.Radians / 2.0))
                                    {
                                        if (angleVec.AngleTo(yVec).Radians >= 0.0 && angleVec.AngleTo(yVec).Radians <= (Angle.PI.Radians / 2.0))
                                        {
                                            angle = Angle.Zero;
                                        }
                                        else
                                        {
                                            angle.Radians = Angle.NEGATIVE_PI.Radians / 2.0;
                                        }
                                    }
                                    // UL and LL
                                    else
                                    {
                                        if (angleVec.AngleTo(yVec).Radians >= 0.0 && angleVec.AngleTo(yVec).Radians <= (Angle.PI.Radians / 2.0))
                                        {
                                            angle.Radians = Angle.PI.Radians / 2.0;
                                        }
                                        else
                                        {
                                            angle.Radians = Angle.PI.Radians;
                                        }
                                    }

                                    DMatrix3d scaleRotateInViewRotateAroundPoint = DMatrix3d.Multiply(invertedViewportRotation,
                                    // DMatrix3d.Rotation(2, angleVec.AngleXY));
                                        DMatrix3d.Rotation(2, angle));

                                    DTransform3d translateAndRotatetoUpInViewAndScale =
                                        DTransform3d.FromMatrixAndTranslation(scaleRotateInViewRotateAroundPoint, translation);

                                    TransformInfo transformInfo2 = new TransformInfo(translateAndRotatetoUpInViewAndScale);

                                    codeElement.ApplyTransform(transformInfo2);

                                    //DPoint3d[] blkPts = new DPoint3d[4];

                                    //DMatrix3d viewMatrix = ev.Viewport.GetRotation();

                                    //// rotate points to orthogonal
                                    //blkPts[0] = new DPoint3d(viewMatrix.Multiply(new DVector3d(m_firstPoint)));

                                    //blkPts[2] = new DPoint3d(viewMatrix.Multiply(new DVector3d(ev.Point)));

                                    //blkPts[2].Z = blkPts[0].Z;

                                    //blkPts[1].X = blkPts[0].X;
                                    //blkPts[1].Y = blkPts[2].Y;
                                    //blkPts[1].Z = blkPts[0].Z;

                                    //blkPts[3].X = blkPts[2].X;
                                    //blkPts[3].Y = blkPts[0].Y;
                                    //blkPts[3].Z = blkPts[0].Z;

                                    //DVector3d dVec = new DVector3d(blkPts[0], blkPts[2]);

                                    //double dWidth = blkPts[0].Distance(blkPts[1]);
                                    //double dHeight = blkPts[0].Distance(blkPts[3]);

                                    //double dDelta = Math.Max(Math.Abs(dWidth), Math.Abs(dHeight));

                                    //blkPts[3] = blkPts[2] = blkPts[1] = blkPts[0];

                                    //DVector3d xVec = new DVector3d(1, 0, 0);

                                    //DVector3d yVec = new DVector3d(0, 1, 0);

                                    //// UR and LR
                                    //if (dVec.AngleTo(xVec).Radians >= 0.0 && dVec.AngleTo(xVec).Radians <= (Angle.PI.Radians / 2.0))
                                    //{
                                    //    if (dVec.AngleTo(yVec).Radians >= 0.0 && dVec.AngleTo(yVec).Radians <= (Angle.PI.Radians / 2.0))
                                    //    {
                                    //        blkPts[2].X = blkPts[0].X + dDelta;
                                    //        blkPts[2].Y = blkPts[0].Y + dDelta;
                                    //    }
                                    //    else
                                    //    {
                                    //        blkPts[2].X = blkPts[0].X + dDelta;
                                    //        blkPts[2].Y = blkPts[0].Y - dDelta;
                                    //    }
                                    //}
                                    //// UL and LL
                                    //else
                                    //{
                                    //    if (dVec.AngleTo(yVec).Radians >= 0.0 && dVec.AngleTo(yVec).Radians <= (Angle.PI.Radians / 2.0))
                                    //    {
                                    //        blkPts[2].X = blkPts[0].X - dDelta;
                                    //        blkPts[2].Y = blkPts[0].Y + dDelta;
                                    //    }
                                    //    else
                                    //    {
                                    //        blkPts[2].X = blkPts[0].X - dDelta;
                                    //        blkPts[2].Y = blkPts[0].Y - dDelta;
                                    //    }
                                    //}

                                    //blkPts[1].X = blkPts[0].X;
                                    //blkPts[1].Y = blkPts[2].Y;
                                    //blkPts[1].Z = blkPts[0].Z;

                                    //blkPts[3].X = blkPts[2].X;
                                    //blkPts[3].Y = blkPts[0].Y;
                                    //blkPts[3].Z = blkPts[0].Z;

                                    //for (int i = 0; i < 4; i++)
                                    //{
                                    //    blkPts[i] = new DPoint3d(invertedViewportRotation.Multiply(new DVector3d(blkPts[i])));
                                    //}

                                    //Bentley.DgnPlatformNET.Elements.ShapeElement shapeElement =
                                    //    new Bentley.DgnPlatformNET.Elements.ShapeElement(Session.Instance.GetActiveDgnModel(), null, blkPts);

                                    RedrawElems redrawElems = new RedrawElems();
                                    redrawElems.SetDynamicsViewsFromActiveViewSet(Bentley.MstnPlatformNET.Session.GetActiveViewport());
                                    redrawElems.DrawMode = DgnDrawMode.Normal;
                                    redrawElems.DrawPurpose = DrawPurpose.ForceRedraw;

                                    codeElement.AddToModel();

                                    redrawElems.DoRedraw(codeElement);

                                    m_haveFirstPoint = false;

                                    m_firstPoint = new DPoint3d(0, 0, 0);

                                    base.BeginDynamics();

                                    return true;
                                }
                            }
                        }
                    }
                }
                else
                {
                    m_haveFirstPoint = true;

                    m_firstPoint = ev.Point;

                    // base.BeginDynamics();
                }
            }
            catch (Exception ex)
            {
                BPSUtilities.WriteLog($"OnDataButton {ex.Message}\n{ex.StackTrace}");
            }

            return true;
        }

        public string LongURL { get; set; }

        public PlaceQRCode(int toolId, int prompt, string sUnparsed) : base(toolId, prompt)
        {
            m_haveFirstPoint = false;
            m_firstPoint = DPoint3d.Zero;
            LongURL = sUnparsed;
        }

        public static void InstallNewInstance(string sUnparsed)
        {
            try
            {
                BPSUtilities.WriteLog("PlaceQRCode InstallNewInstance");
                PlaceQRCode qrCodeTool = new PlaceQRCode(0, 0, sUnparsed);
                qrCodeTool.InstallTool();
            }
            catch (Exception ex)
            {
                BPSUtilities.WriteLog($"PlaceQRCode InstallNewInstance {ex.Message}\n{ex.StackTrace}");
            }
        }

        // private static PolyhedaListForm m_polyhedraListForm = null;

        protected override void ExitTool()
        {
            try
            {
                BPSUtilities.WriteLog("PlaceQRCode ExitTool");

                base.ExitTool();
            }
            catch (Exception ex)
            {
                BPSUtilities.WriteLog($"PlaceQRCode ExitTool {ex.Message}\n{ex.StackTrace}");
            }
        }

        private DPoint3d m_firstPoint = new DPoint3d(0, 0, 0);
        private bool m_haveFirstPoint = false;

        protected override void OnDynamicFrame(Bentley.DgnPlatformNET.DgnButtonEvent ev)
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
                    if (CodeElement != null)
                    {
                        if (CodeElement.IsValid)
                        {
                            List<Bentley.DgnPlatformNET.Elements.Element> listChildElements = CodeElement.GetChildren().GetEnumerator().ConvertToList();

                            Bentley.DgnPlatformNET.Elements.CellHeaderElement codeElement = new CellHeaderElement(Session.Instance.GetActiveDgnModel(), "QRCode",
                                DPoint3d.Zero, DMatrix3d.Identity, listChildElements);

                            if (codeElement.IsValid)
                            {
                                if (m_haveFirstPoint)
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

                                    // double dScale = Math.Max(0.1, Math.Abs(ext.Magnitude / uorPerMaster()));

                                    invertedViewportRotation.ScaleInPlace(ext.Magnitude);

                                    DVector3d xVec = new DVector3d(1, 0, 0);

                                    DVector3d yVec = new DVector3d(0, 1, 0);

                                    Angle angle = Angle.Zero;

                                    //// UR and LR
                                    if (angleVec.AngleTo(xVec).Radians >= 0.0 && angleVec.AngleTo(xVec).Radians <= (Angle.PI.Radians / 2.0))
                                    {
                                        if (angleVec.AngleTo(yVec).Radians >= 0.0 && angleVec.AngleTo(yVec).Radians <= (Angle.PI.Radians / 2.0))
                                        {
                                            angle = Angle.Zero;
                                        }
                                        else
                                        {
                                            angle.Radians = Angle.NEGATIVE_PI.Radians / 2.0;
                                        }
                                    }
                                    // UL and LL
                                    else
                                    {
                                        if (angleVec.AngleTo(yVec).Radians >= 0.0 && angleVec.AngleTo(yVec).Radians <= (Angle.PI.Radians / 2.0))
                                        {
                                            angle.Radians = Angle.PI.Radians / 2.0;
                                        }
                                        else
                                        {
                                            angle.Radians = Angle.PI.Radians;
                                        }
                                    }

                                    DMatrix3d scaleRotateInViewRotateAroundPoint = DMatrix3d.Multiply(invertedViewportRotation,
                                    // DMatrix3d.Rotation(2, angleVec.AngleXY));
                                        DMatrix3d.Rotation(2, angle));

                                    // DMatrix3d scaleRotateInViewRotateAroundPoint = DMatrix3d.Multiply(invertedViewportRotation, 
                                       // DMatrix3d.Rotation(2, angleVec.AngleXY));

                                    DTransform3d translateAndRotatetoUpInViewAndScale =
                                        DTransform3d.FromMatrixAndTranslation(scaleRotateInViewRotateAroundPoint, translation);

                                    TransformInfo transformInfo2 = new TransformInfo(translateAndRotatetoUpInViewAndScale);

                                    codeElement.ApplyTransform(transformInfo2);

                                    // below works great for drawing a block
                                    //DPoint3d[] blkPts = new DPoint3d[4];

                                    //DMatrix3d viewMatrix = ev.Viewport.GetRotation();

                                    //// rotate points to orthogonal
                                    //blkPts[0] = new DPoint3d(viewMatrix.Multiply(new DVector3d(m_firstPoint)));

                                    //blkPts[2] = new DPoint3d(viewMatrix.Multiply(new DVector3d(ev.Point)));

                                    //blkPts[2].Z = blkPts[0].Z;

                                    //blkPts[1].X = blkPts[0].X;
                                    //blkPts[1].Y = blkPts[2].Y;
                                    //blkPts[1].Z = blkPts[0].Z;

                                    //blkPts[3].X = blkPts[2].X;
                                    //blkPts[3].Y = blkPts[0].Y;
                                    //blkPts[3].Z = blkPts[0].Z;

                                    //DVector3d dVec = new DVector3d(blkPts[0], blkPts[2]);

                                    //double dWidth = blkPts[0].Distance(blkPts[1]);
                                    //double dHeight = blkPts[0].Distance(blkPts[3]);

                                    //double dDelta = Math.Abs(dWidth);

                                    //if (Math.Abs(dHeight) > Math.Abs(dWidth))
                                    //    dDelta = Math.Abs(dHeight);

                                    //blkPts[3] = blkPts[2] = blkPts[1] = blkPts[0];

                                    //DVector3d xVec = new DVector3d(1, 0, 0);

                                    //DVector3d yVec = new DVector3d(0, 1, 0);

                                    //double xFactor = 1.0, yFactor = 1.0;

                                    //// UR and LR
                                    //if (dVec.AngleTo(xVec).Radians >= 0.0 && dVec.AngleTo(xVec).Radians <= (Angle.PI.Radians / 2.0))
                                    //{
                                    //    if (dVec.AngleTo(yVec).Radians >= 0.0 && dVec.AngleTo(yVec).Radians <= (Angle.PI.Radians / 2.0))
                                    //    {
                                    //        blkPts[2].X = blkPts[0].X + dDelta;
                                    //        blkPts[2].Y = blkPts[0].Y + dDelta;
                                    //    }
                                    //    else
                                    //    {
                                    //        blkPts[2].X = blkPts[0].X + dDelta;
                                    //        blkPts[2].Y = blkPts[0].Y - dDelta;
                                    //        yFactor = -1.0;
                                    //    }
                                    //}
                                    //// UL and LL
                                    //else
                                    //{
                                    //    if (dVec.AngleTo(yVec).Radians >= 0.0 && dVec.AngleTo(yVec).Radians <= (Angle.PI.Radians / 2.0))
                                    //    {
                                    //        blkPts[2].X = blkPts[0].X - dDelta;
                                    //        blkPts[2].Y = blkPts[0].Y + dDelta;
                                    //        xFactor = -1.0;
                                    //    }
                                    //    else
                                    //    {
                                    //        blkPts[2].X = blkPts[0].X - dDelta;
                                    //        blkPts[2].Y = blkPts[0].Y - dDelta;
                                    //        xFactor = -1.0;
                                    //        yFactor = -1.0;
                                    //    }
                                    //}

                                    //blkPts[1].X = blkPts[0].X;
                                    //blkPts[1].Y = blkPts[2].Y;
                                    //blkPts[1].Z = blkPts[0].Z;

                                    //blkPts[3].X = blkPts[2].X;
                                    //blkPts[3].Y = blkPts[0].Y;
                                    //blkPts[3].Z = blkPts[0].Z;

                                    //DPoint3d[] blkPts2 = new DPoint3d[4];

                                    //blkPts2[0] = blkPts2[1] = blkPts2[2] = blkPts2[3] = blkPts[0];

                                    //for (int i = 0; i < 4; i++)
                                    //{
                                    //    blkPts[i] = new DPoint3d(invertedViewportRotation.Multiply(new DVector3d(blkPts[i])));
                                    //}

                                    //Bentley.DgnPlatformNET.Elements.ShapeElement shapeElement =
                                    //    new Bentley.DgnPlatformNET.Elements.ShapeElement(Session.Instance.GetActiveDgnModel(), null, blkPts);

                                    //blkPts2[0].X += (xFactor) * dDelta / 4.0;
                                    //blkPts2[0].Y += (yFactor) * dDelta / 4.0;

                                    //blkPts2[2].X += (xFactor) * 3.0 * (dDelta / 4.0);
                                    //blkPts2[2].Y += (yFactor) * 3.0 * (dDelta / 4.0);

                                    //blkPts2[1].X = blkPts2[0].X;
                                    //blkPts2[1].Y = blkPts2[2].Y;
                                    //blkPts2[1].Z = blkPts2[0].Z;

                                    //blkPts2[3].X = blkPts2[2].X;
                                    //blkPts2[3].Y = blkPts2[0].Y;
                                    //blkPts2[3].Z = blkPts2[0].Z;

                                    //for (int i = 0; i < 4; i++)
                                    //{
                                    //    blkPts2[i] = new DPoint3d(invertedViewportRotation.Multiply(new DVector3d(blkPts2[i])));
                                    //}

                                    //Bentley.DgnPlatformNET.Elements.ShapeElement shapeElement2 =
                                    //    new Bentley.DgnPlatformNET.Elements.ShapeElement(Session.Instance.GetActiveDgnModel(), null, blkPts2);

                                    //redrawElems.DoRedraw(shapeElement);
                                    //redrawElems.DoRedraw(shapeElement2);
                                }
                                else
                                {
                                    // picking something arbitrary to use as scale factor
                                    invertedViewportRotation.ScaleInPlace(Bentley.MstnPlatformNET.Settings.TextHeight);

                                    DTransform3d scaleAndRotateToUpInView = DTransform3d.FromMatrixAndTranslation(invertedViewportRotation, ev.Point);

                                    TransformInfo transformInfo = new TransformInfo(scaleAndRotateToUpInView);

                                    codeElement.ApplyTransform(transformInfo);
                                }

                                redrawElems.DoRedraw(codeElement);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                BPSUtilities.WriteLog($"{ex.Message}\n{ex.StackTrace}");
            }
        }
    }
}
