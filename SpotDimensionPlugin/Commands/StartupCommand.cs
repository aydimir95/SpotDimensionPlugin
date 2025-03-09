using Autodesk.Revit.Attributes;
using Nice3point.Revit.Toolkit.External;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;

namespace SpotDimensionPlugin.Commands;

/// <summary>
///     External command entry point invoked from the Revit interface
/// </summary>
[UsedImplicitly]
[Transaction(TransactionMode.Manual)]
public class StartupCommand : ExternalCommand
{
    public override void Execute()
    {
        CreateSpotDimension();
    }

    private void CreateSpotDimension()
    {
        UIDocument uiDoc = new UIDocument(Document);
        Document doc = uiDoc.Document;
        View view = Document.ActiveView;

        // For non-2D views, create a workplane if missing.
        if (!(view.ViewType == ViewType.FloorPlan ||
              view.ViewType == ViewType.CeilingPlan ||
              view.ViewType == ViewType.Elevation ||
              view.ViewType == ViewType.Section))
        {
            if (view.SketchPlane == null)
            {
                using (Transaction trans = new Transaction(doc, "Set Workplane"))
                {
                    trans.Start();
                    Plane plane = Plane.CreateByNormalAndOrigin(XYZ.BasisZ, XYZ.Zero);
                    SketchPlane sketchPlane = SketchPlane.Create(doc, plane);
                    view.SketchPlane = sketchPlane;
                    trans.Commit();
                }
            }
        }
        else
        {
            // In 2D views the workplane is fixed internally.
            // To simulate it for point picking, compute it from the associated level.
            if (view.GenLevel == null)
            {
                TaskDialog.Show("Error", "The active view does not have an associated level.");
                return;
            }
            
            // In 2D views the workplane is fixed internally.
            // To allow user point picking, simulate that fixed workplane by computing it
            // from the associated level’s elevation and cut plane offset.
            Level level = doc.GetElement(view.GenLevel.Id) as Level;
            if (level != null)
            {
                // Since the built-in parameter isn't available, default offset to 0.
                double offset = 0.0;
                double z = level.Elevation + offset;

                using (Transaction trans = new Transaction(doc, "Set Fixed Workplane"))
                {
                    trans.Start();
                    Plane fixedPlane = Plane.CreateByNormalAndOrigin(XYZ.BasisZ, new XYZ(0, 0, z));
                    SketchPlane sp = SketchPlane.Create(doc, fixedPlane);
                    view.SketchPlane = sp;
                    trans.Commit();
                }
            }
            else
            {
                TaskDialog.Show("Error", "Could not retrieve the view's associated level.");
                return;
            }
        }

        // Now that the view has a modifiable workplane, PickPoint can succeed.
        XYZ pickedPoint;
        try
        {
            pickedPoint = uiDoc.Selection.PickPoint("Select a point for dimension placement");
        }
        catch (Autodesk.Revit.Exceptions.OperationCanceledException)
        {
            return;
        }

        // Replace the existing "Create Reference Plane" transaction block with this:
        Reference hostReference = null;
        using (Transaction trans = new Transaction(doc, "Create Reference Plane"))
        {
            trans.Start();

            // Get the view direction and up vector to ensure proper orientation
            XYZ viewDir = view.ViewDirection;
            XYZ viewUp = view.UpDirection;

            // Calculate a right vector that's perpendicular to both view direction and up
            XYZ rightDir = viewUp.CrossProduct(viewDir).Normalize();

            // Define non-collinear points that properly define a plane
            XYZ origin = pickedPoint;
            XYZ bubbleEnd = origin + rightDir * 1.0; // Scale factor for visibility
            XYZ freeEnd = origin + viewUp * 1.0;     // Scale factor for visibility

            try
            {
                // Create the reference plane with valid geometry
                ReferencePlane refPlane = doc.Create.NewReferencePlane(bubbleEnd, freeEnd, origin, view);
                hostReference = refPlane.GetReference();
                trans.Commit();
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Reference Plane Error", ex.Message);
                trans.RollBack();
                return;
            }
        }

        // Define additional geometry points for the spot elevation.
        XYZ bend = pickedPoint + new XYZ(0, 10, 0);
        XYZ end = pickedPoint + new XYZ(0, 20, 0);
        XYZ refPt = pickedPoint;

        using (Transaction trans = new Transaction(doc, "Create Spot Elevation"))
        {
            trans.Start();

            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(SpotDimensionType));
            SpotDimensionType spotDimType = collector.FirstElement() as SpotDimensionType;
            if (spotDimType == null)
            {
                TaskDialog.Show("Spot Elevation", "No Spot Elevation Type Found!");
                trans.RollBack();
                return;
            }

            SpotDimension newSpotElevation = doc.Create.NewSpotElevation(view, hostReference, pickedPoint, bend, end, refPt, true);
            TaskDialog.Show("Spot Dimension", $"Created new Spot Dimension with Id: {newSpotElevation.Id}");

            trans.Commit();
        }
    }
}
