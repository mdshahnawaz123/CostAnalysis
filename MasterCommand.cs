using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CostAnalysis
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class MasterCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            List<string> errors = new List<string>();
            int createdCount = 0;
            int scannedLoops = 0;
            int failedProjection = 0;

            try
            {
                View activeView = doc.ActiveView;

                // Guard: FilledRegion creation is not appropriate in 3D views
                if (activeView.ViewType == ViewType.ThreeD)
                {
                    message = "Active view is 3D. Switch to a plan/section/detail view to create filled regions.";
                    return Result.Failed;
                }

                // 1) get all floor instances in active view
                var floors = new FilteredElementCollector(doc, activeView.Id)
                    .OfCategory(BuiltInCategory.OST_Floors)
                    .WhereElementIsNotElementType()
                    .Cast<Floor>()
                    .ToList();

                if (floors.Count == 0)
                {
                    message = "No floor instances found in the active view.";
                    return Result.Failed;
                }

                // 2) pick a FilledRegionType to use (take first one available)
                var filledRegionType = new FilteredElementCollector(doc)
                    .OfClass(typeof(FilledRegionType))
                    .Cast<FilledRegionType>()
                    .FirstOrDefault();

                if (filledRegionType == null)
                {
                    message = "No FilledRegionType found in the document.";
                    return Result.Failed;
                }

                // 3) geometry extraction options
                Options geoOptions = new Options
                {
                    ComputeReferences = false,
                    DetailLevel = ViewDetailLevel.Fine
                };

                // target plane: active view plane
                Autodesk.Revit.DB.Plane viewPlane = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(activeView.ViewDirection, activeView.Origin);

                List<CurveLoop> projectedLoops = new List<CurveLoop>();

                // iterate floors, get face loops and project them
                foreach (var floor in floors)
                {
                    GeometryElement geomElement = floor.get_Geometry(geoOptions);
                    if (geomElement == null) continue;

                    foreach (GeometryObject geomObj in geomElement)
                    {
                        try
                        {
                            // Solid directly
                            if (geomObj is Solid solid)
                            {
                                if (solid.Faces != null && solid.Faces.Size > 0)
                                {
                                    foreach (Face face in solid.Faces)
                                    {
                                        var loops = face.GetEdgesAsCurveLoops();
                                        if (loops == null) continue;

                                        foreach (CurveLoop loop in loops)
                                        {
                                            scannedLoops++;
                                            CurveLoop proj = ProjectCurveLoopToPlane(loop, viewPlane, tessellationPerCurve: 12);
                                            if (proj != null && proj.Count() >= 3)
                                            {
                                                projectedLoops.Add(proj);
                                            }
                                            else failedProjection++;
                                        }
                                    }
                                }
                            }
                            // Geometry instance
                            else if (geomObj is GeometryInstance geomInst)
                            {
                                GeometryElement instGeom = geomInst.GetInstanceGeometry();
                                if (instGeom == null) continue;

                                foreach (GeometryObject instObj in instGeom)
                                {
                                    if (instObj is Solid instSolid)
                                    {
                                        if (instSolid.Faces != null && instSolid.Faces.Size > 0)
                                        {
                                            foreach (Face face in instSolid.Faces)
                                            {
                                                var loops = face.GetEdgesAsCurveLoops();
                                                if (loops == null) continue;

                                                foreach (CurveLoop loop in loops)
                                                {
                                                    scannedLoops++;
                                                    CurveLoop proj = ProjectCurveLoopToPlane(loop, viewPlane, tessellationPerCurve: 12);
                                                    if (proj != null && proj.Count() >= 3)
                                                    {
                                                        projectedLoops.Add(proj);
                                                    }
                                                    else failedProjection++;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        catch (Exception exGeom)
                        {
                            errors.Add($"Geometry processing error (floor id {floor.Id}): {exGeom.Message}");
                        }
                    }
                }

                if (projectedLoops.Count == 0)
                {
                    message = $"No valid projected loops found. Scanned loops: {scannedLoops}, failed projection: {failedProjection}.";
                    if (errors.Count > 0) message += "\nErrors: " + string.Join(" | ", errors);
                    return Result.Failed;
                }

                // 4) create filled regions
                using (Transaction tx = new Transaction(doc, "Create Filled Regions for Floors"))
                {
                    tx.Start();

                    int idx = 0;
                    foreach (var cl in projectedLoops)
                    {
                        idx++;
                        if (cl == null) continue;
                        if (cl.Count() < 3) continue;

                        try
                        {
                            FilledRegion fr = FilledRegion.Create(doc, filledRegionType.Id, activeView.Id, new List<CurveLoop> { cl });
                            if (fr != null) createdCount++;
                        }
                        catch (Exception exCreate)
                        {
                            errors.Add($"Create error for loop #{idx}: {exCreate.Message}");
                        }
                    }

                    tx.Commit();
                }

                // Summary
                string summary = $"Projected loops: {projectedLoops.Count}\nFilled regions created: {createdCount}\nScanned loops: {scannedLoops}\nFailed projections: {failedProjection}";
                if (errors.Count > 0) summary += "\n\nErrors:\n" + string.Join("\n", errors);

                TaskDialog.Show("Create Filled Regions — Summary", summary);
            }
            catch (Exception ex)
            {
                message = $"Unhandled exception: {ex.Message}\n{ex.StackTrace}";
                return Result.Failed;
            }

            return Result.Succeeded;
        }

        // ---------- Helper Methods ----------

        // Project a CurveLoop into a planar CurveLoop on targetPlane using tessellation (approximates curves with lines)
        private CurveLoop ProjectCurveLoopToPlane(CurveLoop inputLoop, Autodesk.Revit.DB.Plane targetPlane, int tessellationPerCurve = 8)
        {
            if (inputLoop == null || targetPlane == null) return null;

            List<XYZ> pts = new List<XYZ>();

            foreach (Curve c in inputLoop)
            {
                if (c == null) continue;

                IList<XYZ> tess = null;
                try
                {
                    tess = c.Tessellate();
                }
                catch
                {
                    tess = null;
                }

                // If basic tessellation returns only endpoints or fails, sample points along the curve
                if (tess == null || tess.Count <= 2)
                {
                    tess = new List<XYZ>();
                    for (int i = 0; i <= tessellationPerCurve; ++i)
                    {
                        double t = (double)i / tessellationPerCurve;
                        // Evaluate may throw for some curve types; guard it
                        try
                        {
                            tess.Add(c.Evaluate(t, true));
                        }
                        catch
                        {
                            // fallback to endpoints if Evaluate fails
                            try
                            {
                                tess.Add(c.GetEndPoint(0));
                                tess.Add(c.GetEndPoint(1));
                                break;
                            }
                            catch
                            {
                                // give up on this curve
                            }
                        }
                    }
                }

                foreach (XYZ p in tess)
                {
                    if (p != null)
                        pts.Add(ProjectPointOntoPlane(p, targetPlane));
                }
            }

            // prune consecutive duplicate points
            double tol = 1e-6;
            List<XYZ> clean = new List<XYZ>();
            for (int i = 0; i < pts.Count; ++i)
            {
                if (i == 0 || pts[i].DistanceTo(pts[i - 1]) > tol)
                    clean.Add(pts[i]);
            }

            if (clean.Count < 3) return null;

            // ensure closure
            if (clean[0].DistanceTo(clean[clean.Count - 1]) > tol)
                clean.Add(clean[0]);

            // build curve loop from straight segments
            CurveLoop outLoop = new CurveLoop();
            for (int i = 0; i < clean.Count - 1; ++i)
            {
                XYZ a = clean[i];
                XYZ b = clean[i + 1];
                if (a.DistanceTo(b) > tol)
                {
                    outLoop.Append(Line.CreateBound(a, b));
                }
            }

            if (outLoop.Count() < 3) return null;

            return outLoop;
        }

        // Project a point onto a Revit Plane
        private XYZ ProjectPointOntoPlane(XYZ p, Autodesk.Revit.DB.Plane plane)
        {
            if (p == null || plane == null) return p;
            double dist = plane.Normal.DotProduct(p - plane.Origin);
            return p - (plane.Normal.Multiply(dist));
        }
    }
}