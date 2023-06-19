#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.Creation;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.ConstrainedExecution;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Input;
using static System.Windows.Forms.LinkLabel;
using Application = Autodesk.Revit.ApplicationServices.Application;
using Document = Autodesk.Revit.DB.Document;
using Excel = Microsoft.Office.Interop.Excel;
#endregion

namespace velcon_Addin
{
    [Transaction(TransactionMode.Manual)]
    public class Command4112: IExternalCommand
    {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;
            Selection sel = uidoc.Selection;

            List<Element> Spaces = (List<Element>)new FilteredElementCollector(doc, doc.ActiveView.Id).OfCategory(BuiltInCategory.OST_MEPSpaces).WhereElementIsNotElementType().ToElements();
            List<Element> fixtures = (List<Element>)new FilteredElementCollector(doc, doc.ActiveView.Id).OfCategory(BuiltInCategory.OST_ElectricalFixtures).WhereElementIsNotElementType().ToElements();

            RevitLinkInstanceSelectionFilter selFilterRevitLinkInstance = new RevitLinkInstanceSelectionFilter();
            Reference selRevitLinkInstance = null;
            try
            {
                selRevitLinkInstance = sel.PickObject(ObjectType.Element, selFilterRevitLinkInstance, "�������� ��������� ����!");
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }

            IEnumerable<RevitLinkInstance> revitLinkInstance = new FilteredElementCollector(doc)
                .OfClass(typeof(RevitLinkInstance))
                .Where(li => li.Id == selRevitLinkInstance.ElementId)
                .Cast<RevitLinkInstance>();
            if (revitLinkInstance.Count() == 0)
            {
                TaskDialog.Show("Revit", "��������� ���� �� ������!");
                return Result.Cancelled;
            }
            Document linkDoc = revitLinkInstance.First().GetLinkDocument();
            Transform transform = revitLinkInstance.First().GetTotalTransform();

            //��������� ���� �� ���������� �����
            List<Wall> wallsInLinkList = new FilteredElementCollector(linkDoc)
                .OfCategory(BuiltInCategory.OST_Walls)
                .OfClass(typeof(Wall))
                .WhereElementIsNotElementType()
                .Cast<Wall>()
                .ToList();
            Dictionary<Element, List<Element>> dict = new Dictionary<Element, List<Element>>();
            Dictionary<Element, List<Line>> boundary =  new Dictionary<Element, List<Line>>();
            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("������� ��� �������");

                foreach (Element sp in Spaces)
                {
                    dict.Add(sp, new List<Element>());
                    foreach (Element fixture in fixtures)
                    {
                        if ((fixture as FamilyInstance).Space != null)
                            if ((fixture as FamilyInstance).Space.Id == sp.Id)
                            {
                                dict[sp].Add(fixture);
                            }
                    }
                    boundary.Add(sp, new List<Line>());
                    GeometryElement ge = (sp as Element).get_Geometry(new Options() { View = doc.ActiveView, IncludeNonVisibleObjects = true, ComputeReferences = true });
                    foreach (var v in ge)
                    {
                        if (v is Line && (Math.Abs((v as Line).Direction.X) == 1 || Math.Abs((v as Line).Direction.Y) == 1))
                            boundary[sp].Add(v as Line);
                    }

                    int linecCol = boundary[sp].Count();
                    List<Line> spLines = new List<Line>();
                   for (int i = 0; i < linecCol; i++)
                    {
                        for (int k = 0; k < linecCol; k++)
                        {
                            if (i!=k)
                            {
                                Line li = boundary[sp][i];
                                Line lk = boundary[sp][k];
                                if (li.Direction.IsAlmostEqualTo(lk.Direction) || li.Direction.IsAlmostEqualTo(lk.Direction.Negate()))
                                {
                                    XYZ sli = li.GetEndPoint(0);
                                    XYZ eli = li.GetEndPoint(1);
                                    XYZ slk = lk.GetEndPoint(0);
                                    XYZ elk = lk.GetEndPoint(1);
                                    if (sli.IsAlmostEqualTo(elk))
                                    {
                                        Line addL = Line.CreateBound(slk, eli);
                                        boundary[sp].Remove(li);
                                        boundary[sp].Remove(lk);
                                        boundary[sp].Add(addL);
                                        linecCol--;
                                        k--;
                                        if (i!=0)
                                        {
                                            i--;
                                        }
                                    }
                                    else if (eli.IsAlmostEqualTo(slk))
                                    {
                                        Line addL = Line.CreateBound(sli, elk);
                                        boundary[sp].Remove(li);
                                        boundary[sp].Remove(lk);
                                        boundary[sp].Add(addL);
                                        linecCol--;
                                        k--;
                                        if (i!=0)
                                        {
                                            i--;
                                        }
                                    }
                                }
                            }
                        }


                    }
                    Dictionary<Line, List<Element>> lines = new Dictionary<Line, List<Element>>();
                    List<Element> ground = new List<Element>();
                    foreach (Line l in boundary[sp])
                    {
                        lines.Add(l, new List<Element>());
                        XYZ min = null;
                        XYZ max = null;
                        if (Math.Abs(l.Direction.Y) == 1)
                        {
                            if (l.GetEndPoint(0).Y > l.GetEndPoint(1).Y)
                            {
                                min = l.GetEndPoint(1);
                                max = l.GetEndPoint(0);
                            }
                            else
                            {
                                min = l.GetEndPoint(0);
                                max = l.GetEndPoint(1);
                            }
                        }
                        else if (Math.Abs(l.Direction.X) == 1)
                        {
                            if (l.GetEndPoint(0).X > l.GetEndPoint(1).X)
                            {
                                min = l.GetEndPoint(1);
                                max = l.GetEndPoint(0);
                            }
                            else
                            {
                                min = l.GetEndPoint(0);
                                max = l.GetEndPoint(1);
                            }
                        }
                        List<Element> newdict = dict[sp];
                        for (int i = 0; i < newdict.Count(); i++)
                        {
                            XYZ elPoint = (newdict[i].Location as LocationPoint).Point;
                            if (Math.Abs((newdict[i] as FamilyInstance).FacingOrientation.X) == 1 && Math.Abs(l.Direction.Y) == 1)
                            {
                                
                                if ((elPoint.X <= l.GetEndPoint(0).X + 0.5 && elPoint.X >= l.GetEndPoint(0).X - 0.5 && elPoint.Y >= min.Y && elPoint.Y <= max.Y) ||
                                    (elPoint.X <= l.GetEndPoint(0).X + 0.5 && elPoint.X >= l.GetEndPoint(0).X - 0.5 && elPoint.Y >= min.Y && elPoint.Y <= max.Y))
                                {
                                    lines[l].Add(newdict[i]);
                                    newdict.RemoveAt(i);
                                    i--;
                                }
                            }
                            else if (Math.Abs((newdict[i] as FamilyInstance).FacingOrientation.Y) == 1 && Math.Abs(l.Direction.X) == 1)
                            {
                                if ((elPoint.Y <= l.GetEndPoint(0).Y + 0.5 && elPoint.Y >= l.GetEndPoint(0).Y - 0.5 && elPoint.X >= min.X && elPoint.X <= max.X) ||
                                    (elPoint.Y <= l.GetEndPoint(0).Y + 0.5 && elPoint.Y >= l.GetEndPoint(0).Y - 0.5 && elPoint.X >= min.X && elPoint.X <= max.X))
                                {
                                    lines[l].Add(newdict[i]);
                                    newdict.RemoveAt(i);
                                    i--;
                                }
                            }
                        }                    
                    }                          
                    List<Line> lkeys = lines.Keys.ToList();
                    foreach (Line l in lkeys)
                    {
                        if (lines[l].Count == 0)
                        {
                            lines.Remove(l);
                        }
                    }
                    foreach (Element e in dict[sp])
                    {
                        bool usl = true;
                        foreach (Line l in lines.Keys)
                        {
                            if (lines[l].Contains(e)) usl = false;
                        }
                        if (usl)
                            ground.Add(e);
                    }
                    List<Line> Keys = lines.Keys.ToList();
                    foreach (Line l in Keys)
                    {
                        List<Element> ElsToAdd = new List<Element>();
                        double xmin = double.MaxValue;
                        double ymin = double.MaxValue;
                        double xmax = double.MinValue;
                        double ymax = double.MinValue;
                        Element min = null;
                        Element max = null;
                        List<Element> elLines = lines[l];
                        foreach (Element el in elLines)
                        {
                            XYZ lp = (el.Location as LocationPoint).Point;
                            XYZ face = (el as FamilyInstance).FacingOrientation;
                            if (Math.Abs(face.Y) == 1)
                            {
                                if (lp.X > xmax)
                                {
                                    xmax = lp.X;
                                    max = el;
                                }
                                if (lp.X < xmin)
                                {
                                    xmin = lp.X;
                                    min = el;
                                }
                            }
                            else
                            {
                                if (lp.Y > ymax)
                                {
                                    ymax = lp.Y;
                                    max = el;
                                }
                                if (lp.Y < ymin)
                                {
                                    ymin = lp.Y;
                                    min = el;
                                }
                            }
                            Element bliz = null;
                            double dist = double.MaxValue;
                            Element bliz2 = null;
                            double dist2 = double.MaxValue;
                            foreach (Element e in lines[l])
                            {
                                if (el.Id != e.Id)
                                {
                                    XYZ lp2 = (e.Location as LocationPoint).Point;
                                    if (lp.DistanceTo(lp2) < dist)
                                    {
                                        bliz = e;
                                        dist = lp2.DistanceTo(lp);
                                    }
                                }
                            }
                            foreach (Element e in lines[l])
                            {
                                if (el.Id != e.Id)
                                {
                                    XYZ lp2 = (e.Location as LocationPoint).Point;
                                    if (Math.Abs((e as FamilyInstance).FacingOrientation.X) == 1)
                                    {
                                        if (((bliz.Location as LocationPoint).Point.Y < lp.Y && lp2.Y > lp.Y) ||
                                            ((bliz.Location as LocationPoint).Point.Y > lp.Y && lp2.Y < lp.Y))
                                        {
                                            if (lp.DistanceTo(lp2) < dist2)
                                            {
                                                bliz2 = e;
                                                dist2 = lp2.DistanceTo(lp);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if (((bliz.Location as LocationPoint).Point.X < lp.X && lp2.X > lp.X) ||
                                            ((bliz.Location as LocationPoint).Point.X > lp.X && lp2.X < lp.X))
                                        {
                                            if (lp.DistanceTo(lp2) < dist2)
                                            {
                                                bliz2 = e;
                                                dist2 = lp2.DistanceTo(lp);
                                            }
                                        }
                                    }
                                }
                            }
                            if (dist >= 3.280839895013123 && dist != double.MaxValue)
                            {
                                if (ElsToAdd.Contains(bliz))
                                {
                                    int ind = ElsToAdd.IndexOf(bliz);
                                    if ((ind == 0 && ElsToAdd[ind + 1] != el) && (ind != 0 && ElsToAdd[ind-1] != el))
                                    {
                                        ElsToAdd.Add(el);
                                        ElsToAdd.Add(bliz);
                                    }
                                }
                                else
                                {
                                    ElsToAdd.Add(el);
                                    ElsToAdd.Add(bliz);
                                }
                            }
                            if (dist2 >= 3.280839895013123 && dist2 != double.MaxValue)
                            {
                                if (ElsToAdd.Contains(bliz2))
                                {
                                    int ind = ElsToAdd.IndexOf(bliz2);
                                    if ((ind == 0 && ElsToAdd[ind + 1] != el) && (ind != 0 && ElsToAdd[ind - 1] != el))
                                    {
                                        ElsToAdd.Add(el);
                                        ElsToAdd.Add(bliz2);
                                    }
                                }
                                else
                                {
                                    ElsToAdd.Add(el);
                                    ElsToAdd.Add(bliz2);
                                }
                            }
                        }
                        lines[l] = ElsToAdd;
                        for (int i = 0; i < lines[l].Count(); i = i + 2)
                        {
                            //������� ����� ���������
                            XYZ facing = (lines[l][0] as FamilyInstance).FacingOrientation;
                            ReferenceArray ra = new ReferenceArray();
                            Reference clr1 = (lines[l][i] as FamilyInstance).GetReferences(FamilyInstanceReferenceType.CenterLeftRight).First();
                            Reference clr2 = (lines[l][i+1] as FamilyInstance).GetReferences(FamilyInstanceReferenceType.CenterLeftRight).First();
                            ra.Append(clr1);
                            ra.Append(clr2);
                            Line l1_ = null;
                            if (Math.Abs(facing.X) == 1)
                            {
                               l1_ = Line.CreateBound((lines[l][i].Location as LocationPoint).Point + facing * 1.5, new XYZ((lines[l][i].Location as LocationPoint).Point.X, (lines[l][i+1].Location as LocationPoint).Point.Y, (lines[l][i].Location as LocationPoint).Point.Z) + facing * 1.5);
                            }
                            else
                            {
                                l1_ = Line.CreateBound((lines[l][i].Location as LocationPoint).Point + facing * 1.5, new XYZ((lines[l][i+1].Location as LocationPoint).Point.X, (lines[l][i].Location as LocationPoint).Point.Y, (lines[l][i].Location as LocationPoint).Point.Z) + facing * 1.5);

                            }
                            Element dim = doc.Create.NewDimension(doc.ActiveView, l1_, ra);
                        }
                        //������ ��� min � max, �������� �������
                        double dist1 = double.MaxValue;
                        double dist22 = double.MaxValue;
                        Element clWall1 = null;
                        Element clWall2 = null;
                        XYZ linkP1 = null;
                        XYZ linkP2 = null;
                        Line cl1 = null;
                        Line cl2 = null;
                        foreach (Element wall in wallsInLinkList)
                        {
                            if ((linkDoc.GetElement(wall.LevelId) as Level).Elevation == (doc.GetElement(min.LevelId) as Level).Elevation)
                            {
                                if (((wall.Location as LocationCurve).Curve as Line).Direction.IsAlmostEqualTo((min as FamilyInstance).FacingOrientation,0.05) ||
                                    ((wall.Location as LocationCurve).Curve as Line).Direction.Negate().IsAlmostEqualTo((min as FamilyInstance).FacingOrientation, 0.05))

                                {
                                    Line linkP = Line.CreateBound(transform.OfPoint((wall.Location as LocationCurve).Curve.GetEndPoint(0)),
                                        transform.OfPoint((wall.Location as LocationCurve).Curve.GetEndPoint(1)));
                                    Curve linkc = (wall.Location as LocationCurve).Curve.CreateTransformed(transform);
                                    XYZ wallp = (transform.OfPoint((wall.Location as LocationCurve).Curve.GetEndPoint(0)) + transform.OfPoint((wall.Location as LocationCurve).Curve.GetEndPoint(1))) / 2;
                                    double v = double.MinValue;
                                    double n = double.MinValue;
                                    double vs = double.MinValue;
                                    if (Math.Abs((min as FamilyInstance).FacingOrientation.X) == 1)
                                    {
                                            v = linkP.GetEndPoint(0).X;
                                        n = linkP.GetEndPoint(1).X;
                                        if (v<n)
                                        {
                                            double nzap = n;
                                            n = v;
                                            v= nzap;
                                        }
                                        vs = linkP.GetEndPoint(1).Y;
                                    }
                                    else
                                    {
                                            v = linkP.GetEndPoint(0).Y;
                                        n = linkP.GetEndPoint(1).Y;
                                        if (v < n)
                                        {
                                            double nzap = n;
                                            n = v;
                                            v = nzap;
                                        }
                                        vs = linkP.GetEndPoint(1).X;
                                    }

                                    foreach (Line lbound in boundary[sp])
                                    {
                                        double vslbound = double.MinValue;
                                        if (Math.Abs(lbound.Direction.X) == 1)
                                        {
                                            vslbound = lbound.GetEndPoint(0).Y;
                                        }
                                        else
                                        {
                                            vslbound = lbound.GetEndPoint(0).X;
                                        }

                                        if ((vs <= vslbound + 1.7 && vs >= vslbound - 1.7 && (min.Location as LocationPoint).Point.Y <= v + 1.5 && (min.Location as LocationPoint).Point.Y >= n - 1.5) ||
                                            (vs <= vslbound + 1.7 && vs >= vslbound - 1.7 && (min.Location as LocationPoint).Point.X <= v + 1.5 && (min.Location as LocationPoint).Point.X >= n - 1.5))
                                        {
                                            double dist1_ = linkP.Distance((min.Location as LocationPoint).Point); //lp.DistanceTo(linkP);
                                            if (dist1_ <= dist1)
                                            {
                                                dist1 = dist1_;
                                                clWall1 = wall;
                                                linkP1 = wallp;
                                                cl1 = lbound;
                                            }
                                            double dist2_ = linkP.Distance((max.Location as LocationPoint).Point); //lp.DistanceTo(linkP);
                                            if (dist2_ <= dist22)
                                            {
                                                dist22 = dist2_;
                                                clWall2 = wall;
                                                linkP2 = wallp;
                                                cl2 = lbound;
                                            }
                                        }
                                    }

                                }
                                
                            }
                        }
                        ReferenceArray ra1 = new ReferenceArray();
                        ReferenceArray ra2 = new ReferenceArray();
                        ReferenceArray ra3 = new ReferenceArray();
                        ReferenceArray ra4 = new ReferenceArray();
                        Reference exteriorFaceRef1 = HostObjectUtils.GetSideFaces(clWall1 as Wall, ShellLayerType.Exterior).First<Reference>();
                        Reference exteriorFaceRef2 = HostObjectUtils.GetSideFaces(clWall1 as Wall, ShellLayerType.Interior).First<Reference>();

                        Reference linkToExteriorFaceRef1 = exteriorFaceRef1.CreateLinkReference(revitLinkInstance.First());
                        Reference linkToExteriorFaceRef2 = exteriorFaceRef2.CreateLinkReference(revitLinkInstance.First());
                        string sr = linkToExteriorFaceRef1.ConvertToStableRepresentation(doc);
                        var s = sr.Split(':');
                        sr = "";
                        foreach (string s1 in s)
                        {
                            if (s1.Contains("RVTLINK"))
                            {
                                sr += ":0:RVTLINK";
                            }
                            else
                            {
                                sr += ":" + s1;
                            }
                        }
                        sr = sr.Substring(1);
                        ra1.Append(Reference.ParseFromStableRepresentation(doc, sr));
                        sr = linkToExteriorFaceRef2.ConvertToStableRepresentation(doc);
                        s = sr.Split(':');
                        sr = "";
                        foreach (string s1 in s)
                        {
                            if (s1.Contains("RVTLINK"))
                            {
                                sr += ":0:RVTLINK";
                            }
                            else
                            {
                                sr += ":" + s1;
                            }
                        }
                        sr = sr.Substring(1);
                        ra2.Append(Reference.ParseFromStableRepresentation(doc, sr));

                        exteriorFaceRef1 = HostObjectUtils.GetSideFaces(clWall2 as Wall, ShellLayerType.Exterior).First<Reference>();
                        exteriorFaceRef2 = HostObjectUtils.GetSideFaces(clWall2 as Wall, ShellLayerType.Interior).First<Reference>();

                        linkToExteriorFaceRef1 = exteriorFaceRef1.CreateLinkReference(revitLinkInstance.First());
                        linkToExteriorFaceRef2 = exteriorFaceRef2.CreateLinkReference(revitLinkInstance.First());
                        sr = linkToExteriorFaceRef1.ConvertToStableRepresentation(doc);
                        s = sr.Split(':');
                        sr = "";
                        foreach (string s1 in s)
                        {
                            if (s1.Contains("RVTLINK"))
                            {
                                sr += ":0:RVTLINK";
                            }
                            else
                            {
                                sr += ":" + s1;
                            }
                        }
                        sr = sr.Substring(1);
                        ra3.Append(Reference.ParseFromStableRepresentation(doc, sr));
                        sr = linkToExteriorFaceRef2.ConvertToStableRepresentation(doc);
                        s = sr.Split(':');
                        sr = "";
                        foreach (string s1 in s)
                        {
                            if (s1.Contains("RVTLINK"))
                            {
                                sr += ":0:RVTLINK";
                            }
                            else
                            {
                                sr += ":" + s1;
                            }
                        }
                        sr = sr.Substring(1);
                        ra4.Append(Reference.ParseFromStableRepresentation(doc, sr));

                        ra1.Append((min as FamilyInstance).GetReferences(FamilyInstanceReferenceType.CenterLeftRight).First());
                        ra2.Append((min as FamilyInstance).GetReferences(FamilyInstanceReferenceType.CenterLeftRight).First());
                        ra3.Append((max as FamilyInstance).GetReferences(FamilyInstanceReferenceType.CenterLeftRight).First());
                        ra4.Append((max as FamilyInstance).GetReferences(FamilyInstanceReferenceType.CenterLeftRight).First());
                        
                        Line l1 = null;
                        Line l2 = null;
                        XYZ point1 = (min.Location as LocationPoint).Point;
                        XYZ point2 = (max.Location as LocationPoint).Point;
                        XYZ fomin = (min as FamilyInstance).FacingOrientation;

                        if (Math.Abs(fomin.Y) == 1)
                        {
                            l1 = Line.CreateBound(point1 + (min as FamilyInstance).FacingOrientation * 1.5, new XYZ(linkP1.X, point1.Y, point1.Z) + (min as FamilyInstance).FacingOrientation * 1.5);
                            l2 = Line.CreateBound(point2 + (max as FamilyInstance).FacingOrientation * 1.5, new XYZ(linkP2.X, point2.Y, point2.Z) + (max as FamilyInstance).FacingOrientation * 1.5);
                        }
                        else if (Math.Abs(fomin.X) == 1)
                        {
                            l1 = Line.CreateBound(point1 + (min as FamilyInstance).FacingOrientation * 1.5, new XYZ(point1.X, linkP1.Y, point1.Z) + (min as FamilyInstance).FacingOrientation * 1.5);
                            l2 = Line.CreateBound(point2 + (max as FamilyInstance).FacingOrientation * 1.5, new XYZ(point2.X, linkP2.Y, point2.Z) + (max as FamilyInstance).FacingOrientation * 1.5);
                        }



                        if ((min as FamilyInstance).FacingOrientation.X == 1)
                        {
                            if (l1 != null)
                            {
                                Element d1 = doc.Create.NewDimension(doc.ActiveView, l1, ra1);
                                Element d2 = doc.Create.NewDimension(doc.ActiveView, l1, ra2);
                                if (((cl1.GetEndPoint(0) + cl1.GetEndPoint(1)) / 2).X > point1.X)
                                {
                                    if ((d1 as Dimension).Value > (d2 as Dimension).Value)
                                    {
                                        doc.Delete(d1.Id);
                                    }
                                    else
                                    {
                                        doc.Delete(d2.Id);
                                    }
                                }
                                else
                                {
                                    if ((d1 as Dimension).Value < (d2 as Dimension).Value)
                                    {
                                        doc.Delete(d1.Id);
                                    }
                                    else
                                    {
                                        doc.Delete(d2.Id);
                                    }
                                }
                            }
                            if (l2 != null)
                            {
                                Element d3 = doc.Create.NewDimension(doc.ActiveView, l2, ra3);
                                Element d4 = doc.Create.NewDimension(doc.ActiveView, l2, ra4);
                                if (((cl2.GetEndPoint(0) + cl2.GetEndPoint(1)) / 2).X > point2.X)
                                {
                                    if ((d3 as Dimension).Value > (d4 as Dimension).Value)
                                    {
                                        doc.Delete(d3.Id);
                                    }
                                    else
                                    {
                                        doc.Delete(d4.Id);
                                    }
                                }
                                else
                                {
                                    if ((d3 as Dimension).Value < (d4 as Dimension).Value)
                                    {
                                        doc.Delete(d3.Id);
                                    }
                                    else
                                    {
                                        doc.Delete(d4.Id);
                                    }
                                }
                            }
                        }
                        else if ((min as FamilyInstance).FacingOrientation.X == -1)
                        {
                            if (l1 != null)
                            {
                                Element d1 = doc.Create.NewDimension(doc.ActiveView, l1, ra1);
                                Element d2 = doc.Create.NewDimension(doc.ActiveView, l1, ra2);
                                if (((cl1.GetEndPoint(0) + cl1.GetEndPoint(1)) / 2).X > point1.X)
                                {
                                    if ((d1 as Dimension).Value < (d2 as Dimension).Value)
                                    {
                                        doc.Delete(d1.Id);
                                    }
                                    else
                                    {
                                        doc.Delete(d2.Id);
                                    }
                                }
                                else
                                {
                                    if ((d1 as Dimension).Value > (d2 as Dimension).Value)
                                    {
                                        doc.Delete(d1.Id);
                                    }
                                    else
                                    {
                                        doc.Delete(d2.Id);
                                    }
                                }
                            }
                            if (l2 != null)
                            {
                                Element d3 = doc.Create.NewDimension(doc.ActiveView, l2, ra3);
                                Element d4 = doc.Create.NewDimension(doc.ActiveView, l2, ra4);
                                if (((cl2.GetEndPoint(0) + cl2.GetEndPoint(1)) / 2).X > point2.X)
                                {
                                    if ((d3 as Dimension).Value < (d4 as Dimension).Value)
                                    {
                                        doc.Delete(d3.Id);
                                    }
                                    else
                                    {
                                        doc.Delete(d4.Id);
                                    }
                                }
                                else
                                {
                                    if ((d3 as Dimension).Value > (d4 as Dimension).Value)
                                    {
                                        doc.Delete(d3.Id);
                                    }
                                    else
                                    {
                                        doc.Delete(d4.Id);
                                    }
                                }
                            }
                        }
                        else if ((min as FamilyInstance).FacingOrientation.Y == 1)
                        {
                            if (l1 != null)
                            {
                                Element d1 = doc.Create.NewDimension(doc.ActiveView, l1, ra1);
                                Element d2 = doc.Create.NewDimension(doc.ActiveView, l1, ra2);
                                if ((cl1.GetEndPoint(0) + cl1.GetEndPoint(1) / 2).Y > point1.Y)
                                {
                                    if ((d1 as Dimension).Value > (d2 as Dimension).Value)
                                    {
                                        doc.Delete(d1.Id);
                                    }
                                    else
                                    {
                                        doc.Delete(d2.Id);
                                    }
                                }
                                else
                                {
                                    if ((d1 as Dimension).Value < (d2 as Dimension).Value)
                                    {
                                        doc.Delete(d1.Id);
                                    }
                                    else
                                    {
                                        doc.Delete(d2.Id);
                                    }
                                }
                            }
                            if (l2 != null)
                            {
                                Element d3 = doc.Create.NewDimension(doc.ActiveView, l2, ra3);
                                Element d4 = doc.Create.NewDimension(doc.ActiveView, l2, ra4);
                                if (((cl2.GetEndPoint(0) + cl2.GetEndPoint(1)) / 2).Y > point2.Y)
                                {
                                    if ((d3 as Dimension).Value > (d4 as Dimension).Value)
                                    {
                                        doc.Delete(d3.Id);
                                    }
                                    else
                                    {
                                        doc.Delete(d4.Id);
                                    }
                                }
                                else
                                {
                                    if ((d3 as Dimension).Value < (d4 as Dimension).Value)
                                    {
                                        doc.Delete(d3.Id);
                                    }
                                    else
                                    {
                                        doc.Delete(d4.Id);
                                    }
                                }
                            }
                        }
                        else
                        {
                            if (l1 != null)
                            {
                                Element d1 = doc.Create.NewDimension(doc.ActiveView, l1, ra1);
                                Element d2 = doc.Create.NewDimension(doc.ActiveView, l1, ra2);
                                if (((cl1.GetEndPoint(0) + cl1.GetEndPoint(1)) / 2).Y > point1.X)
                                {
                                    if ((d1 as Dimension).Value < (d2 as Dimension).Value)
                                    {
                                        doc.Delete(d1.Id);
                                    }
                                    else
                                    {
                                        doc.Delete(d2.Id);
                                    }
                                }
                                else
                                {
                                    if ((d1 as Dimension).Value > (d2 as Dimension).Value)
                                    {
                                        doc.Delete(d1.Id);
                                    }
                                    else
                                    {
                                        doc.Delete(d2.Id);
                                    }
                                }
                            }
                            if (l2 != null)
                            {
                                Element d3 = doc.Create.NewDimension(doc.ActiveView, l2, ra3);
                                Element d4 = doc.Create.NewDimension(doc.ActiveView, l2, ra4);
                                if (((cl2.GetEndPoint(0) + cl2.GetEndPoint(1)) / 2).Y > point2.Y)
                                {
                                    if ((d3 as Dimension).Value < (d4 as Dimension).Value)
                                    {
                                        doc.Delete(d3.Id);
                                    }
                                    else
                                    {
                                        doc.Delete(d4.Id);
                                    }
                                }
                                else
                                {
                                    if ((d3 as Dimension).Value > (d4 as Dimension).Value)
                                    {
                                        doc.Delete(d3.Id);
                                    }
                                    else
                                    {
                                        doc.Delete(d4.Id);
                                    }
                                }
                            }
                        }
                    }
                    //������� ��� ground
                    foreach (Element element in ground)
                    {
                        Reference cfb = (element as FamilyInstance).GetReferences(FamilyInstanceReferenceType.CenterFrontBack).First();
                        Reference clr = (element as FamilyInstance).GetReferences(FamilyInstanceReferenceType.CenterLeftRight).First();
                        ReferenceArray ra1 = new ReferenceArray();
                        ReferenceArray ra2 = new ReferenceArray();
                        ReferenceArray ra3 = new ReferenceArray();
                        ReferenceArray ra4 = new ReferenceArray();
                        double dist1 = double.MaxValue;
                        double dist2 = double.MaxValue;
                        XYZ point = (element.Location as LocationPoint).Point;
                        Line x = null;
                        Line y = null;
                        Element clWall1 = null;
                        Element clWall2 = null;
                        XYZ linkP1 = null;
                        XYZ linkP2 = null;
                        foreach (Element wall in wallsInLinkList)
                        {
                            if ((linkDoc.GetElement(wall.LevelId) as Level).Elevation == (doc.GetElement(element.LevelId) as Level).Elevation)
                            {
                                if (Math.Abs(((wall.Location as LocationCurve).Curve as Line).Direction.X) == 1)

                                {
                                    Line linkP = Line.CreateBound(transform.OfPoint((wall.Location as LocationCurve).Curve.GetEndPoint(0)),
                                        transform.OfPoint((wall.Location as LocationCurve).Curve.GetEndPoint(1)));
                                    Curve linkc = (wall.Location as LocationCurve).Curve.CreateTransformed(transform);
                                    XYZ wallp = (transform.OfPoint((wall.Location as LocationCurve).Curve.GetEndPoint(0)) + transform.OfPoint((wall.Location as LocationCurve).Curve.GetEndPoint(1))) / 2;
                                    double n;
                                    double v;
                                    if (linkP.GetEndPoint(0).X > linkP.GetEndPoint(1).X)
                                    {
                                        v = linkP.GetEndPoint(0).X;
                                        n = linkP.GetEndPoint(1).X;
                                    }
                                    else
                                    {
                                        v = linkP.GetEndPoint(1).X;
                                        n = linkP.GetEndPoint(0).X;
                                    }

                                    foreach (Line lbound in boundary[sp])
                                    {
                                        if (((linkP.GetEndPoint(0) + linkP.GetEndPoint(1)) / 2).Y <= ((lbound.GetEndPoint(0) + lbound.GetEndPoint(1)) / 2).Y + 1.5 &&
                                            ((linkP.GetEndPoint(0) + linkP.GetEndPoint(1)) / 2).Y >= ((lbound.GetEndPoint(0) + lbound.GetEndPoint(1)) / 2).Y - 1.5 &&
                                            wallp.X <= v + 0.5 && wallp.X >= n + 0.5)
                                        {
                                            double dist1_ = linkP.Distance(point); //lp.DistanceTo(linkP);
                                            if (dist1_ < dist1)
                                            {
                                                dist1 = dist1_;
                                                clWall1 = wall;
                                                linkP1 = wallp;
                                            }
                                        }
                                    }

                                }
                                else if (Math.Abs(((wall.Location as LocationCurve).Curve as Line).Direction.Y) == 1)
                                {
                                    Line linkP = Line.CreateBound(transform.OfPoint((wall.Location as LocationCurve).Curve.GetEndPoint(0)),
                                        transform.OfPoint((wall.Location as LocationCurve).Curve.GetEndPoint(1)));
                                    Curve linkc = (wall.Location as LocationCurve).Curve.CreateTransformed(transform);
                                    XYZ wallp = (transform.OfPoint((wall.Location as LocationCurve).Curve.GetEndPoint(0)) + transform.OfPoint((wall.Location as LocationCurve).Curve.GetEndPoint(1))) / 2;
                                    double n;
                                    double v;
                                    if (linkP.GetEndPoint(0).Y > linkP.GetEndPoint(1).Y)
                                    {
                                        v = linkP.GetEndPoint(0).Y;
                                        n = linkP.GetEndPoint(1).Y;
                                    }
                                    else
                                    {
                                        v = linkP.GetEndPoint(1).Y;
                                        n = linkP.GetEndPoint(0).Y;
                                    }
                                    foreach (Line lbound in boundary[sp])
                                    {
                                        if (((linkP.GetEndPoint(0) + linkP.GetEndPoint(1)) / 2).X <= ((lbound.GetEndPoint(0) + lbound.GetEndPoint(1)) / 2).X + 1.5 &&
                                            ((linkP.GetEndPoint(0) + linkP.GetEndPoint(1)) / 2).X >= ((lbound.GetEndPoint(0) + lbound.GetEndPoint(1)) / 2).X - 1.5 &&
                                            wallp.Y <= v + 0.5 && wallp.Y >= n + 0.5)
                                        {
                                            double dist2_ = linkP.Distance(point); //lp.DistanceTo(linkP);
                                            if (dist2_ < dist2)
                                            {
                                                dist2 = dist2_;
                                                clWall2 = wall;
                                                linkP2 = (transform.OfPoint((wall.Location as LocationCurve).Curve.GetEndPoint(0)) + transform.OfPoint((wall.Location as LocationCurve).Curve.GetEndPoint(1))) / 2;
                                            }
                                        }
                                    }

                                }
                            }
                        }
                        Reference exteriorFaceRef1 = HostObjectUtils.GetSideFaces(clWall1 as Wall, ShellLayerType.Exterior).First<Reference>();
                        Reference exteriorFaceRef2 = HostObjectUtils.GetSideFaces(clWall1 as Wall, ShellLayerType.Interior).First<Reference>();

                        Reference linkToExteriorFaceRef1 = exteriorFaceRef1.CreateLinkReference(revitLinkInstance.First());
                        Reference linkToExteriorFaceRef2 = exteriorFaceRef2.CreateLinkReference(revitLinkInstance.First());
                        string sr = linkToExteriorFaceRef1.ConvertToStableRepresentation(doc);
                        var s = sr.Split(':');
                        sr = "";
                        foreach (string s1 in s)
                        {
                            if (s1.Contains("RVTLINK"))
                            {
                                sr += ":0:RVTLINK";
                            }
                            else
                            {
                                sr += ":" + s1;
                            }
                        }
                        sr = sr.Substring(1);
                        ra1.Append(Reference.ParseFromStableRepresentation(doc, sr));
                        sr = linkToExteriorFaceRef2.ConvertToStableRepresentation(doc);
                        s = sr.Split(':');
                        sr = "";
                        foreach (string s1 in s)
                        {
                            if (s1.Contains("RVTLINK"))
                            {
                                sr += ":0:RVTLINK";
                            }
                            else
                            {
                                sr += ":" + s1;
                            }
                        }
                        sr = sr.Substring(1);
                        ra2.Append(Reference.ParseFromStableRepresentation(doc, sr));

                        exteriorFaceRef1 = HostObjectUtils.GetSideFaces(clWall2 as Wall, ShellLayerType.Exterior).First<Reference>();
                        exteriorFaceRef2 = HostObjectUtils.GetSideFaces(clWall2 as Wall, ShellLayerType.Interior).First<Reference>();

                        linkToExteriorFaceRef1 = exteriorFaceRef1.CreateLinkReference(revitLinkInstance.First());
                        linkToExteriorFaceRef2 = exteriorFaceRef2.CreateLinkReference(revitLinkInstance.First());
                        sr = linkToExteriorFaceRef1.ConvertToStableRepresentation(doc);
                        s = sr.Split(':');
                        sr = "";
                        foreach (string s1 in s)
                        {
                            if (s1.Contains("RVTLINK"))
                            {
                                sr += ":0:RVTLINK";
                            }
                            else
                            {
                                sr += ":" + s1;
                            }
                        }
                        sr = sr.Substring(1);
                        ra3.Append(Reference.ParseFromStableRepresentation(doc, sr));
                        sr = linkToExteriorFaceRef2.ConvertToStableRepresentation(doc);
                        s = sr.Split(':');
                        sr = "";
                        foreach (string s1 in s)
                        {
                            if (s1.Contains("RVTLINK"))
                            {
                                sr += ":0:RVTLINK";
                            }
                            else
                            {
                                sr += ":" + s1;
                            }
                        }
                        sr = sr.Substring(1);
                        ra4.Append(Reference.ParseFromStableRepresentation(doc, sr));

                        ra1.Append(cfb);
                        ra2.Append(cfb);

                        Line l1 = null;
                        if (Math.Abs(linkP1.Y - point.Y) > 0.1)
                            if (point.X > linkP1.X)
                                l1 = Line.CreateBound(new XYZ(point.X - 1.5, point.Y, point.Z), new XYZ(point.X - 1.5, linkP1.Y, point.Z));
                            else
                                l1 = Line.CreateBound(new XYZ(point.X + 1.5, point.Y, point.Z), new XYZ(point.X + 1.5, linkP1.Y, point.Z));


                        ra3.Append(clr);
                        ra4.Append(clr);
                        Line l2 = null;
                        if (Math.Abs(linkP1.X - point.X) > 0.1)
                            if (point.Y > linkP1.Y)
                                l2 = Line.CreateBound(new XYZ(point.X, point.Y - 1.5, point.Z), new XYZ(linkP2.X, point.Y - 1.5, point.Z));
                            else
                                l2 = Line.CreateBound(new XYZ(point.X, point.Y + 1.5, point.Z), new XYZ(linkP2.X, point.Y + 1.5, point.Z));
                        if (l2 != null)
                        {
                            Element d1 = doc.Create.NewDimension(doc.ActiveView, l2, ra1);
                            Element d2 = doc.Create.NewDimension(doc.ActiveView, l2, ra2);
                            if ((d1 as Dimension).Value > (d2 as Dimension).Value)
                            {
                                doc.Delete(d1.Id);
                            }
                            else
                            {
                                doc.Delete(d2.Id);
                            }
                        }
                        if (l1 != null)
                        {
                            Element d3 = doc.Create.NewDimension(doc.ActiveView, l1, ra3);
                            Element d4 = doc.Create.NewDimension(doc.ActiveView, l1, ra4);
                            if ((d3 as Dimension).Value > (d4 as Dimension).Value)
                            {
                                doc.Delete(d3.Id);
                            }
                            else
                            {
                                doc.Delete(d4.Id);
                            }
                        }
                    }
                }


                tx.Commit();
            }
            return Result.Succeeded;
        }
    }
}
