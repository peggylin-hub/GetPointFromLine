using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Xbim.Ifc4.Kernel;
using Xbim.Ifc4.Interfaces;
using Xbim.Ifc;
using System.Reflection;
using Xbim.Ifc4.RepresentationResource;
using Xbim.Ifc4.GeometricModelResource;
using Xbim.Ifc4.ProductExtension;
using Xbim.Ifc4.TopologyResource;
using Xbim.Ifc4.GeometryResource;
using Xbim.Ifc4.MeasureResource;
using Xbim.Ifc2x3.Interfaces;
using Xbim.Ifc4.SharedBldgElements;
using Xbim.Ifc4.GeometricConstraintResource;
using Xbim.Common;

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Geometry;

using Autodesk.Civil.DatabaseServices;
using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices.Styles;
using System.Windows.Forms;

namespace TunnalCal.C3D
{
    class Alignments
    {
        public static Alignment CreateAlignmentFromIFC(string ifcFile)
        {
            Alignment aln = null;
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            CivilDocument Cdoc = CivilApplication.ActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            ObjectId aln_id = ObjectId.Null;
            ObjectId layerId = ObjectId.Null;
            ObjectId aln_styleId = ObjectId.Null;
            ObjectId profStyleId = ObjectId.Null;
            ObjectId alnLabelSetId = ObjectId.Null;
            ObjectId profLabelSetId = ObjectId.Null;

            using (Transaction CADtrans = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)CADtrans.GetObject(db.BlockTableId, OpenMode.ForRead, false);
                BlockTableRecord btr = (BlockTableRecord)CADtrans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite, false);

                LayerTable acLyrTbl = CADtrans.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;
                foreach (ObjectId id in acLyrTbl)
                {
                    layerId = id;
                    break;
                }

                AlignmentStyleCollection style_col = Cdoc.Styles.AlignmentStyles;
                if (style_col.Count > 0)
                    aln_styleId = style_col[0];

                AlignmentLabelSetStyleCollection alignmentLabelSets = Cdoc.Styles.LabelSetStyles.AlignmentLabelSetStyles;
                if (alignmentLabelSets.Count > 0)
                    alnLabelSetId = alignmentLabelSets[0];

                aln_id = Alignment.Create(Cdoc, "Alingment_" + DateTime.Now.ToShortTimeString(), ObjectId.Null, layerId, aln_styleId, alnLabelSetId, AlignmentType.Centerline);
                aln = CADtrans.GetObject(aln_id, OpenMode.ForRead) as Alignment;

                int noOfSpiral = 0;

                string str = string.Empty;
                using (var model = IfcStore.Open(ifcFile))
                {
                    var ifcAlignments = model.Instances.OfType<IfcAlignment>();

                    try
                    {
                        foreach (IfcAlignment algn in ifcAlignments)
                        {

                            IfcCurve cv = algn.Axis;

                            IfcAlignmentCurve algnCv = cv as IfcAlignmentCurve;
                            if (algnCv != null)
                            {
                                IfcAlignment2DHorizontal algn_hort = algnCv.Horizontal;

                                foreach (IfcAlignment2DHorizontalSegment seg in algn_hort.Segments)
                                {
                                    IfcCurveSegment2D seg_2d = seg.CurveGeometry;


                                    if (seg_2d.GetType() == typeof(IfcLineSegment2D))
                                    {
                                        //get ifc data
                                        IfcLineSegment2D line = seg_2d as IfcLineSegment2D;
                                        IfcCartesianPoint pt = line.StartPoint;
                                        IfcPlaneAngleMeasure dir = line.StartDirection;//this is bearing angle
                                        IfcPositiveLengthMeasure len = line.SegmentLength;

                                        double angle = bearing2AutocadAngle(dir);
                                        ///create civil 3d line entity
                                        //create point 1
                                        Point3d pt1 = new Point3d(pt.X, pt.Y, 0);
                                        //create point 2
                                        double x = Math.Sin(angle) * len;
                                        double y = Math.Cos(angle) * len;
                                        Vector3d vector = new Vector3d(x, y, 0);
                                        Matrix3d mat = Matrix3d.Displacement(vector);
                                        Point3d pt2 = pt1.TransformBy(mat);

                                        ed.WriteMessage("Crate line\n");
                                        aln.Entities.AddFixedLine(aln.Entities.Count(), pt1, pt2);
                                    }

                                    if (seg_2d.GetType() == typeof(IfcTransitionCurveSegment2D))
                                    {
                                        noOfSpiral++;

                                        IfcTransitionCurveSegment2D tran = seg_2d as IfcTransitionCurveSegment2D;

                                        double rad = 0;
                                        double startRadius = 0, endRadius = 0;
                                        if (tran.StartRadius != null)
                                            startRadius = tran.StartRadius.Value;

                                        if (tran.EndRadius != null)
                                            endRadius = tran.EndRadius.Value;

                                        if (startRadius != 0 && endRadius == 0)
                                            rad = startRadius;
                                        else if (endRadius != 0 && startRadius == 0)
                                            rad = endRadius;

                                        IfcCartesianPoint pt = tran.StartPoint;

                                        IfcPositiveLengthMeasure len = tran.SegmentLength;

                                        bool IsEndRadiusCCW = false, IsStartRadiusCCW = false, isClockwise = false;
                                        if (tran.IsEndRadiusCCW != null)
                                        {
                                            IsEndRadiusCCW = tran.IsEndRadiusCCW;
                                        }

                                        if (tran.IsStartRadiusCCW != null)
                                        {
                                            IsStartRadiusCCW = tran.IsStartRadiusCCW;
                                        }

                                        if (IsStartRadiusCCW == true)
                                            isClockwise = IsStartRadiusCCW;

                                        double dir = tran.StartDirection;

                                        IfcTransitionCurveType tranType = tran.TransitionCurveType;

                                        //determine if it in curve or out curve by no?
                                        SpiralCurveType spiralCurveType = SpiralCurveType.InCurve;
                                        if (noOfSpiral % 2 == 0)
                                            spiralCurveType = SpiralCurveType.OutCurve;

                                        if (tran.TransitionCurveType == IfcTransitionCurveType.CLOTHOIDCURVE)
                                        {
                                            ed.WriteMessage("Crate Transition Curve\n");
                                            if (rad == 0)
                                                aln.Entities.AddFixedSpiral(aln.Entities.Count(), startRadius, endRadius, len, Autodesk.Civil.SpiralType.Clothoid);
                                            else if (startRadius != 0)
                                                aln.Entities.AddFloatSpiral(aln.Entities.Count(), rad, len, isClockwise, Autodesk.Civil.SpiralType.Clothoid);
                                                //aln.Entities.AddFixedSpiral(aln.Entities.Count(), rad, len, spiralCurveType, Autodesk.Civil.SpiralType.Clothoid);
                                        }
                                    }

                                    if (seg_2d.GetType() == typeof(IfcCircularArcSegment2D))
                                    {
                                        IfcCircularArcSegment2D cir = seg_2d as IfcCircularArcSegment2D;
                                        IfcPositiveLengthMeasure len = cir.SegmentLength;
                                        IfcCartesianPoint pt = cir.StartPoint;
                                        IfcPlaneAngleMeasure dir = cir.StartDirection;
                                        IfcPositiveLengthMeasure rad = cir.Radius;

                                        double angle = bearing2AutocadAngle(dir);
                                        bool isCCW = (bool) cir.IsCCW;

                                        Point3d startPt = new Point3d(pt.X, pt.Y, 0);
                                        double x = Math.Sin(angle) * len;
                                        double y = Math.Cos(angle) * len;
                                        Vector3d vector = new Vector3d(x, y, 0);

                                        //ed.WriteMessage("Crate circle Curve\n");
                                      //  aln.Entities.AddFixedCurve(startPt, vector, rad, isCCW);
                                    }

                                    if (seg_2d.GetType() == typeof(IfcCurveSegment2D))
                                    {
                                        IfcCurveSegment2D crv = seg_2d as IfcCurveSegment2D;

                                        IfcPositiveLengthMeasure len = crv.SegmentLength;

                                        IfcCartesianPoint pt = crv.StartPoint;

                                        IfcPlaneAngleMeasure dir = crv.StartDirection;

                                    }


                                }


                                #region >>>>=========== Create profile entities
                                // get the standard style and label set
                                // these calls will fail on templates without a style named "Standard"
                                ProfileStyleCollection prof_tyle_col = Cdoc.Styles.ProfileStyles;
                                if (prof_tyle_col.Count > 0)
                                    profStyleId = prof_tyle_col[0];

                                ProfileLabelSetStyleCollection profilLabelSets = Cdoc.Styles.LabelSetStyles.ProfileLabelSetStyles;
                                if (profilLabelSets.Count > 0)
                                    profLabelSetId = profilLabelSets[0];

                                ObjectId oProfileId = Profile.CreateByLayout("Profile", aln_id, layerId, profStyleId, profLabelSetId);
                                Profile PROF = CADtrans.GetObject(oProfileId, OpenMode.ForWrite) as Profile;

                                //get profile information from ifc
                                IfcAlignment2DVertical algn_vert = algnCv.Vertical;

                                foreach (IfcAlignment2DVerticalSegment seg in algn_vert.Segments)
                                {

                                    if (seg.GetType() == typeof(IfcAlignment2DVerSegLine))
                                    {
                                        IfcAlignment2DVerSegLine line = seg as IfcAlignment2DVerSegLine;

                                        double len = line.HorizontalLength;

                                        double z = Convert.ToDouble(line.StartHeight.Value);

                                        double slope = Convert.ToDouble(line.StartGradient.Value);

                                    }

                                    if (seg.GetType() == typeof(IfcAlignment2DVerSegParabolicArc))
                                    {
                                        IfcAlignment2DVerSegParabolicArc tran = seg as IfcAlignment2DVerSegParabolicArc;

                                        double parabola = Convert.ToDouble(tran.ParabolaConstant);

                                        bool isConvex = (bool)tran.IsConvex.Value;

                                        double ch = Convert.ToDouble(tran.StartDistAlong);

                                        double len = Convert.ToDouble(tran.HorizontalLength);

                                        double z = Convert.ToDouble(tran.StartHeight.Value);

                                        double slope = Convert.ToDouble(tran.StartGradient.Value);

                                    }

                                    if (seg.GetType() == typeof(IfcAlignment2DVerSegCircularArc))
                                    {
                                        IfcAlignment2DVerSegCircularArc cir = seg as IfcAlignment2DVerSegCircularArc;

                                        double radius = Convert.ToDouble(cir.Radius);

                                        bool isConvex = (bool)cir.IsConvex.Value;

                                        double ch = Convert.ToDouble(cir.StartDistAlong);

                                        double len = Convert.ToDouble(cir.HorizontalLength);

                                        double z = Convert.ToDouble(cir.StartHeight.Value);

                                        double slope = Convert.ToDouble(cir.StartGradient.Value);

                                    }


                                    //Point2d startPoint = new Point2d(station_start, H_start);
                                    //Point2d endPoint = new Point2d(station_end, H_end);
                                    //ProfileTangent oTangent = PROF.Entities.AddFixedTangent(startPoint, endPoint);
                                    
                                    #endregion
                                }
                            }

                        }
                    }
                    catch { }


                }

                CADtrans.Commit();
            }

            return aln;
        }

        private static double bearing2AutocadAngle(double bearing)
        {
            return Math.PI * 2.5 - bearing;  
        }
    }
}
