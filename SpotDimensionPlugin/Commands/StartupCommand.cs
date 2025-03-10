using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Nice3point.Revit.Toolkit.External;
using System;
using System.Collections.Generic;
using System.Linq;

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
        CreateSpotDimensionOnAllSectionViews();
    }

    /// <summary>
    /// Creates spot elevations on all section views for a user-selected face
    /// </summary>
    private void CreateSpotDimensionOnAllSectionViews()
    {
        UIDocument uiDoc = new UIDocument(Document);
        Document doc = uiDoc.Document;
        View currentView = uiDoc.ActiveView;

        // Verify we have a valid view
        if (currentView == null)
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
                "Select a face for spot elevations on all section views");
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
        ElementId elementId = elem.Id;

        // Get the picked point on the face (in global coordinates)
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
        dirDialog.MainInstruction = "Choose the direction for the spot elevation leaders";
        dirDialog.MainContent = "Should the leaders point to the LEFT?\n\nClick 'Yes' for left side, 'No' for right side.";
        dirDialog.CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No;
        TaskDialogResult result = dirDialog.Show();

        // Determine direction based on dialog result
        bool placeOnLeft = (result == TaskDialogResult.Yes);

        // Collect all section views in the document
        FilteredElementCollector viewCollector = new FilteredElementCollector(doc);
        viewCollector.OfClass(typeof(View));
        List<View> sectionViews = new List<View>();

        foreach (View view in viewCollector)
        {
            // Only include section views that are not templates and not system browser views
            if (view.ViewType == ViewType.Section &&
                !view.IsTemplate &&
                view.CanBePrinted)
            {
                sectionViews.Add(view);
            }
        }

        if (sectionViews.Count == 0)
        {
            TaskDialog.Show("Error", "No section views found in the document.");
            return;
        }

        // Keep track of success/failure counts
        int successCount = 0;
        List<string> failedViews = new List<string>();

        // Create a transaction for all the spot elevations
        using (Transaction trans = new Transaction(doc, "Create Spot Elevations in All Section Views"))
        {
            trans.Start();

            foreach (View sectionView in sectionViews)
            {
                try
                {
                    // Get orientation vectors for this section view
                    XYZ viewDir = sectionView.ViewDirection;
                    XYZ viewUp = sectionView.UpDirection;
                    XYZ viewRight = viewUp.CrossProduct(viewDir).Normalize();

                    // Set direction factor based on user choice
                    double directionFactor = placeOnLeft ? -1.0 : 1.0;

                    // Create offset points for the leader with shoulder
                    XYZ bendPoint = pickedPoint + (viewRight * directionFactor * 3.0) + (viewUp * 1.0);
                    XYZ endPoint = pickedPoint + (viewRight * directionFactor * 7.0) + (viewUp * 1.0);

                    // Make sure spot elevation category is visible in this view
                    Category spotCategory = Category.GetCategory(doc, BuiltInCategory.OST_SpotElevations);
                    if (spotCategory != null)
                    {
                        // Make the category visible in this view
                        sectionView.SetCategoryHidden(spotCategory.Id, false);
                    }

                    // Create a special reference for the spot elevation
                    SpotDimension spotElevation = null;

                    // APPROACH 1: Try to use the original reference directly
                    try
                    {
                        spotElevation = doc.Create.NewSpotElevation(
                            sectionView, pickedFaceRef, pickedPoint, bendPoint, endPoint, pickedPoint, true);
                    }
                    catch
                    {
                        // Reference not valid in this view - try another approach
                    }

                    // APPROACH 2: If that didn't work, try to create a face in this view's context
                    if (spotElevation == null)
                    {
                        Options options = new Options();
                        options.View = sectionView;
                        options.ComputeReferences = true;

                        GeometryElement geomElem = elem.get_Geometry(options);
                        if (geomElem != null)
                        {
                            // Try to find a face with a valid reference
                            foreach (GeometryObject geomObj in geomElem)
                            {
                                Solid solid = geomObj as Solid;
                                if (solid != null)
                                {
                                    foreach (Face face in solid.Faces)
                                    {
                                        try
                                        {
                                            if (face.Reference != null)
                                            {
                                                // Try creating with this face reference
                                                try
                                                {
                                                    spotElevation = doc.Create.NewSpotElevation(
                                                        sectionView, face.Reference, pickedPoint, bendPoint, endPoint, pickedPoint, true);

                                                    if (spotElevation != null)
                                                        break;
                                                }
                                                catch
                                                {
                                                    // This reference didn't work; try the next one
                                                }
                                            }
                                        }
                                        catch
                                        {
                                            // If we can't evaluate this face, just skip it
                                        }
                                    }

                                    if (spotElevation != null)
                                        break;
                                }
                            }
                        }
                    }

                    // APPROACH 3: If all else fails, try to find ANY valid face in the element
                    if (spotElevation == null)
                    {
                        // One last attempt - find ANY valid face on the element in this view
                        Options options = new Options();
                        options.View = sectionView;
                        options.ComputeReferences = true;

                        GeometryElement geomElem = elem.get_Geometry(options);
                        if (geomElem != null)
                        {
                            foreach (GeometryObject geomObj in geomElem)
                            {
                                Solid solid = geomObj as Solid;
                                if (solid != null && solid.Faces.Size > 0)
                                {
                                    foreach (Face face in solid.Faces)
                                    {
                                        if (face.Reference != null)
                                        {
                                            try
                                            {
                                                spotElevation = doc.Create.NewSpotElevation(
                                                    sectionView, face.Reference, pickedPoint, bendPoint, endPoint, pickedPoint, true);

                                                if (spotElevation != null)
                                                    break;
                                            }
                                            catch
                                            {
                                                // Try the next face
                                            }
                                        }
                                    }

                                    if (spotElevation != null)
                                        break;
                                }
                            }
                        }
                    }

                    if (spotElevation != null)
                    {
                        // Success! Try to make it more visible
                        successCount++;

                        try
                        {
                            // Set graphic overrides to make it more visible
                            OverrideGraphicSettings ogs = new OverrideGraphicSettings();
                            ogs.SetProjectionLineWeight(4); // Thicker lines
                            sectionView.SetElementOverrides(spotElevation.Id, ogs);

                            // Check for common parameters that might make it more visible
                            foreach (Parameter param in spotElevation.Parameters)
                            {
                                // Adjust leader settings if possible
                                if (param.Definition.Name.Contains("Leader") && !param.IsReadOnly)
                                {
                                    if (param.StorageType == StorageType.Double)
                                    {
                                        param.Set(5.0); // Make leader more prominent
                                    }
                                }

                                // Check for text size or visibility parameters
                                if (param.Definition.Name.Contains("Text") && !param.IsReadOnly)
                                {
                                    if (param.StorageType == StorageType.Double)
                                    {
                                        param.Set(3.0); // Make text larger
                                    }
                                }
                            }
                        }
                        catch
                        {
                            // If adjusting parameters fails, still consider it a success
                        }
                    }
                    else
                    {
                        failedViews.Add($"{sectionView.Name} (no valid face reference)");
                    }
                }
                catch (Exception ex)
                {
                    failedViews.Add($"{sectionView.Name} ({ex.Message})");
                }
            }

            trans.Commit();
        }

        // Show results dialog
        TaskDialog resultsDialog = new TaskDialog("Spot Elevation Results");
        resultsDialog.MainInstruction = $"Created {successCount} spot elevations out of {sectionViews.Count} section views";

        if (failedViews.Count > 0)
        {
            string failureMessage = "Failed in the following views:\n";
            foreach (string viewName in failedViews)
            {
                failureMessage += $"- {viewName}\n";
            }

            resultsDialog.MainContent = failureMessage +
                "\n\nIf you don't see any spot elevations, please check:\n" +
                "- Scroll/zoom out in the section views to find them\n" +
                "- Check if Annotations are visible in View Control Bar\n" +
                "- Verify Spot Elevations are enabled in Visibility/Graphics";
        }
        else if (successCount == 0)
        {
            resultsDialog.MainContent = "No spot elevations could be created.\n\n" +
                "Try selecting a different face that appears in multiple section views.";
        }
        else
        {
            resultsDialog.MainContent = "Spot elevations were created successfully.\n\n" +
                "If you don't see them immediately:\n" +
                "- Try scrolling/zoom out in your section views\n" +
                "- Check if Annotations are visible in View Control Bar\n" +
                "- Check Visibility/Graphics settings for Spot Elevations\n" +
                "- Try selecting 'Select All Instances' in Revit to find them";
        }

        // Add helpful buttons to navigate to views with spot elevations
        resultsDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1,
            "Show All Section Views in Project Browser");

        TaskDialogResult tdResult = resultsDialog.Show();

        // If user clicks to show all section views
        if (tdResult == TaskDialogResult.CommandLink1)
        {
            // Request Revit to show Views in project browser
            List<ElementId> viewIds = new List<ElementId>();
            foreach (View view in sectionViews)
            {
                viewIds.Add(view.Id);
            }
            uiDoc.ShowElements(viewIds);
        }
    }
}