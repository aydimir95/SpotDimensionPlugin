using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Nice3point.Revit.Toolkit.External;

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
        CreateSpotDimensionOnSelectedFace();
    }

    /// <summary>
    /// Creates a spot elevation on a user-selected face
    /// </summary>
    private void CreateSpotDimensionOnSelectedFace()
    {
        UIDocument uiDoc = new UIDocument(Document);
        Document doc = uiDoc.Document;
        View view = Document.ActiveView;

        // Verify we have a valid view
        if (view == null)
        {
            TaskDialog.Show("Error", "No active view found.");
            return;
        }

        // Allow the user to directly select a face
        Reference pickedFaceRef;
        try
        {
            pickedFaceRef = uiDoc.Selection.PickObject(
                ObjectType.Face,
                "Select a face for the spot elevation");
        }
        catch (Autodesk.Revit.Exceptions.OperationCanceledException)
        {
            return; // User canceled the selection
        }

        if (pickedFaceRef == null)
        {
            TaskDialog.Show("Error", "No face was selected.");
            return;
        }

        // Get the element that owns the face
        Element elem = doc.GetElement(pickedFaceRef.ElementId);

        // Get the picked point on the face
        XYZ pickedPoint = pickedFaceRef.GlobalPoint;

        // If the GlobalPoint is null, try to get a point on the face
        if (pickedPoint == null)
        {
            // Alternative approach to get a point on the face
            GeometryObject geoObject = elem.GetGeometryObjectFromReference(pickedFaceRef);
            Face face = geoObject as Face;
            if (face != null)
            {
                // Get a point in the middle of the face's bounding box
                BoundingBoxUV bbox = face.GetBoundingBox();
                UV midUV = (bbox.Min + bbox.Max) * 0.5;
                pickedPoint = face.Evaluate(midUV);
            }
            else
            {
                TaskDialog.Show("Error", "Could not determine a point on the selected face.");
                return;
            }
        }

        // Simple yes/no dialog for direction selection
        TaskDialog dirDialog = new TaskDialog("Leader Direction");
        dirDialog.MainInstruction = "Choose the direction for the spot elevation leader";
        dirDialog.MainContent = "Should the leader point to the LEFT?\n\nClick 'Yes' for left side, 'No' for right side.";
        dirDialog.CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No;
        TaskDialogResult result = dirDialog.Show();

        // Determine direction based on dialog result
        bool placeOnLeft = (result == TaskDialogResult.Yes);

        // Get view orientation vectors
        XYZ viewDir = view.ViewDirection;
        XYZ viewUp = view.UpDirection;
        XYZ viewRight = viewUp.CrossProduct(viewDir).Normalize();

        // Set direction factor based on user choice
        double directionFactor = placeOnLeft ? -1.0 : 1.0;

        // Create offset points for the leader with shoulder
        // First point bends horizontally and has a vertical shoulder component
        XYZ bendPoint = pickedPoint + (viewRight * directionFactor * 3.0) + (viewUp * 1.0);

        // End point continues in same horizontal direction at same height as bend point
        XYZ endPoint = pickedPoint + (viewRight * directionFactor * 7.0) + (viewUp * 1.0);

        // Create the spot elevation
        using (Transaction trans = new Transaction(doc, "Create Spot Elevation"))
        {
            trans.Start();

            try
            {
                // Get the spot elevation type (or use default)
                FilteredElementCollector collector = new FilteredElementCollector(doc);
                collector.OfClass(typeof(SpotDimensionType));
                SpotDimensionType spotDimType = collector.FirstElement() as SpotDimensionType;

                if (spotDimType == null)
                {
                    TaskDialog.Show("Warning", "No spot elevation type found. Using default.");
                }

                SpotDimension spotElevation = doc.Create.NewSpotElevation(
                    view, pickedFaceRef, pickedPoint, bendPoint, endPoint, pickedPoint, true);

                if (spotElevation != null)
                {
                    TaskDialog.Show("Success", $"Created spot elevation with ID: {spotElevation.Id}");
                    trans.Commit();
                }
                else
                {
                    TaskDialog.Show("Error", "Failed to create spot elevation.");
                    trans.RollBack();
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", $"Failed to create spot elevation: {ex.Message}");
                trans.RollBack();
            }
        }
    }
}