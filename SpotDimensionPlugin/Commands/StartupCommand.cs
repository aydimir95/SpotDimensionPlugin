using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Nice3point.Revit.Toolkit.External;
using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using Newtonsoft.Json;

namespace SpotDimensionPlugin.Commands
{
    /// <summary>
    /// External command entry point invoked from the Revit interface.
    /// </summary>
    [UsedImplicitly]
    [Transaction(TransactionMode.Manual)]
    public class StartupCommand : ExternalCommand
    {
        // Log list for face matching results.
        private List<FaceMatchLog> _faceMatchLogs = new List<FaceMatchLog>();

        // Mapping of face directions to local direction vectors.
        // Only four options are provided.
        private Dictionary<string, XYZ> _faceDirections = new Dictionary<string, XYZ>
        {
            { "Left Face",   new XYZ(-1, 0, 0) },
            { "Right Face",  new XYZ( 1, 0, 0) },
            { "Front Face",  new XYZ( 0, 1, 0) },
            { "Back Face",   new XYZ( 0,-1, 0) }
        };

        public override void Execute()
        {
            CreateSpotDimensionOnPreselectedElements();
        }

        /// <summary>
        /// Creates spot elevations on pre-selected elements in all section views
        /// using a face-matching transform approach and forcing the view to Fine detail.
        /// Logs intermediate face evaluation values to JSON.
        /// </summary>
        private void CreateSpotDimensionOnPreselectedElements()
        {
            UIDocument uiDoc = new UIDocument(Document);
            Document doc = uiDoc.Document;
            View currentView = uiDoc.ActiveView;

            // 1. Verify active view.
            if (currentView == null)
            {
                TaskDialog.Show("Error", "No active view found.");
                return;
            }

            // 2. Get pre-selected elements.
            ICollection<ElementId> preselectedIds = uiDoc.Selection.GetElementIds();
            if (preselectedIds.Count == 0)
            {
                TaskDialog.Show("Error", "No elements selected. Please select elements first.");
                return;
            }
            List<Element> selectedElements = new List<Element>();
            foreach (ElementId id in preselectedIds)
            {
                Element elem = doc.GetElement(id);
                if (elem != null)
                    selectedElements.Add(elem);
            }
            if (selectedElements.Count == 0)
            {
                TaskDialog.Show("Error", "No valid elements found in selection.");
                return;
            }

            // Notify user.
            TaskDialog.Show("Spot Elevations",
                $"Creating spot elevations for {selectedElements.Count} pre-selected elements.\n\n" +
                "You'll be asked to choose a face direction for matching.");

            // 3. Prompt the user to choose a face direction.
            string chosenFaceDirection = PromptForFaceDirection();
            if (chosenFaceDirection == null)
            {
                TaskDialog.Show("Error", "No face direction chosen. Aborting.");
                return;
            }
            XYZ localDir = _faceDirections[chosenFaceDirection];

            // 4. Prompt user for leader side.
            bool placeOnLeft = PromptForLeaderSide();

            // 5. Collect all section views.
            FilteredElementCollector viewCollector = new FilteredElementCollector(doc).OfClass(typeof(View));
            List<View> sectionViews = viewCollector
                .Cast<View>()
                .Where(v => v.ViewType == ViewType.Section && !v.IsTemplate && v.CanBePrinted)
                .ToList();
            if (sectionViews.Count == 0)
            {
                TaskDialog.Show("Error", "No section views found in the document.");
                return;
            }

            int successCount = 0;
            Dictionary<string, int> successPerView = new Dictionary<string, int>();
            List<string> failedItems = new List<string>();

            // 6. Declare localPoint at method scope.
            XYZ localPoint = null;

            // 7. Get the template face by prompting user to pick a face.
            Reference pickedFaceRef;
            try
            {
                PreselectedElementsFilter filter = new PreselectedElementsFilter(preselectedIds);
                pickedFaceRef = uiDoc.Selection.PickObject(
                    ObjectType.Face,
                    filter,
                    "Select a face on one of the pre-selected elements");
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return;
            }
            if (pickedFaceRef == null)
            {
                TaskDialog.Show("Error", "No face was selected.");
                return;
            }
            Element templateElem = doc.GetElement(pickedFaceRef.ElementId);
            GeometryObject geoObj = templateElem.GetGeometryObjectFromReference(pickedFaceRef);
            Face selectedFace = geoObj as Face;
            if (selectedFace == null)
            {
                TaskDialog.Show("Error", "Could not analyze the selected face.");
                return;
            }
            // 8. Get the global pick point.
            XYZ pickedPoint = pickedFaceRef.GlobalPoint;
            if (pickedPoint == null)
            {
                BoundingBoxUV faceBB = selectedFace.GetBoundingBox();
                UV midUV = (faceBB.Min + faceBB.Max) * 0.5;
                pickedPoint = selectedFace.Evaluate(midUV);
            }
            // Compute the face normal.
            BoundingBoxUV bbox = selectedFace.GetBoundingBox();
            UV centerUV = (bbox.Min + bbox.Max) * 0.5;
            XYZ templateNormal = selectedFace.ComputeNormal(centerUV).Normalize();

            // 9. If the template element is a FamilyInstance, get its transform and set localPoint.
            FamilyInstance templateInstance = templateElem as FamilyInstance;
            string templateFamilyName = null;
            ElementId templateTypeId = null;
            Transform templateTransform = null;
            if (templateInstance != null)
            {
                templateFamilyName = templateInstance.Symbol.Family.Name;
                templateTypeId = templateInstance.GetTypeId();
                templateTransform = templateInstance.GetTransform();
                localPoint = templateTransform.Inverse.OfPoint(pickedPoint);
            }

            // 10. Begin a transaction group.
            using (TransactionGroup tg = new TransactionGroup(doc, "Create Spot Elevations (StopAfterSuccess)"))
            {
                tg.Start();

                foreach (View sectionView in sectionViews)
                {
                    using (Transaction trans = new Transaction(doc, $"Spot Elevations in {sectionView.Name}"))
                    {
                        trans.Start();
                        successPerView[sectionView.Name] = 0;

                        // ***** Force Detail Level to Fine *****
                        Parameter detailParam = sectionView.get_Parameter(BuiltInParameter.VIEW_DETAIL_LEVEL);
                        if (detailParam != null && !detailParam.IsReadOnly)
                        {
                            detailParam.Set(3); // 3 = Fine
                        }
                        // ***** Optionally force discipline to Coordination if available *****
                        Parameter discParam = sectionView.get_Parameter(BuiltInParameter.VIEW_DISCIPLINE);
                        if (discParam != null && !discParam.IsReadOnly)
                        {
                            try { discParam.Set(4095); } catch { }
                        }
                        // **********************************************

                        // Ensure Spot Elevations are visible.
                        Category spotCat = Category.GetCategory(doc, BuiltInCategory.OST_SpotElevations);
                        if (spotCat != null && sectionView.GetCategoryHidden(spotCat.Id))
                        {
                            try { sectionView.SetCategoryHidden(spotCat.Id, false); }
                            catch (Exception ex)
                            {
                                failedItems.Add($"Failed to unhide Spot Elevations in {sectionView.Name}: {ex.Message}");
                            }
                        }

                        // Orientation for leader lines.
                        XYZ viewDir = sectionView.ViewDirection.Normalize();
                        XYZ viewUp = sectionView.UpDirection.Normalize();
                        XYZ viewRight = viewUp.CrossProduct(viewDir).Normalize();
                        double directionFactor = placeOnLeft ? -1.0 : 1.0;

                        // Pre-calculate default leader offsets.
                        XYZ defaultBendPoint = pickedPoint + (viewRight * directionFactor * 3.0) + (viewUp * 1.0);
                        XYZ defaultEndPoint = pickedPoint + (viewRight * directionFactor * 7.0) + (viewUp * 1.0);

                        // Loop over each selected element.
                        foreach (Element elem in selectedElements)
                        {
                            SpotDimension spotElevation = null;
                            bool approachUsed = false;
                            try
                            {
                                string elemCat = elem.Category?.Name ?? "NoCategory";
                                string elemInfo = $"ID:{elem.Id.Value} ({elem.GetType().Name}, {elemCat})";

                                // Quick geometry check.
                                Options gOpt = new Options { IncludeNonVisibleObjects = true };
                                GeometryElement ge = elem.get_Geometry(gOpt);
                                if (ge == null)
                                {
                                    failedItems.Add($"No geometry for {elemInfo} in {sectionView.Name}");
                                    continue;
                                }

                                // Approach 1: If it's exactly the template element, use the picked face.
                                if (elem.Id == templateElem.Id)
                                {
                                    approachUsed = true;
                                    try
                                    {
                                        spotElevation = doc.Create.NewSpotElevation(
                                            sectionView,
                                            pickedFaceRef,
                                            pickedPoint,
                                            defaultBendPoint,
                                            defaultEndPoint,
                                            pickedPoint,
                                            true
                                        );
                                    }
                                    catch (Exception ex)
                                    {
                                        failedItems.Add($"Template element {elemInfo} in {sectionView.Name}: {ex.Message}");
                                    }
                                    if (spotElevation != null)
                                    {
                                        successCount++;
                                        successPerView[sectionView.Name]++;
                                        StyleSpotElevation(doc, sectionView, spotElevation, failedItems, elemInfo);
                                        continue; // Stop further attempts for this element.
                                    }
                                }

                                // Approach 1B: Face-matching transform approach.
                                if (templateInstance != null && elem is FamilyInstance fi)
                                {
                                    // Relax type check: only compare family names.
                                    bool sameFamily = (fi.Symbol.Family.Name == templateFamilyName);
                                    if (sameFamily && localPoint != null)
                                    {
                                        approachUsed = true;
                                        try
                                        {
                                            Transform instTransform = fi.GetTransform();
                                            XYZ instanceGlobalPoint = instTransform.OfPoint(localPoint);

                                            Options optGeom = new Options
                                            {
                                                View = sectionView,
                                                ComputeReferences = true,
                                                IncludeNonVisibleObjects = true,
                                                DetailLevel = ViewDetailLevel.Fine
                                            };

                                            // Use local-direction based face search.
                                            Face bestFace = FindFaceByLocalDirection(doc, fi, sectionView, localDir);
                                            if (bestFace != null && bestFace.Reference != null)
                                            {
                                                BoundingBoxUV faceBB = bestFace.GetBoundingBox();
                                                UV midUV = (faceBB.Min + faceBB.Max) * 0.5;
                                                XYZ faceCenter = bestFace.Evaluate(midUV);

                                                XYZ bendP = faceCenter + (viewRight * directionFactor * 3.0) + (viewUp * 1.0);
                                                XYZ endP = faceCenter + (viewRight * directionFactor * 7.0) + (viewUp * 1.0);

                                                spotElevation = doc.Create.NewSpotElevation(
                                                    sectionView,
                                                    bestFace.Reference,
                                                    faceCenter,
                                                    bendP,
                                                    endP,
                                                    faceCenter,
                                                    true
                                                );
                                                if (spotElevation != null)
                                                {
                                                    failedItems.Add($"Debug: Created spot on {fi.Id.Value} via transform approach in {sectionView.Name}");
                                                    successCount++;
                                                    successPerView[sectionView.Name]++;
                                                    StyleSpotElevation(doc, sectionView, spotElevation, failedItems, elemInfo);
                                                    continue; // Stop further attempts for this element.
                                                }
                                            }
                                            else
                                            {
                                                failedItems.Add($"Face search for {elemInfo} in {sectionView.Name}: No suitable face found.");
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            failedItems.Add($"Transform approach for {elemInfo}: {ex.Message}");
                                        }
                                    }
                                }

                                if (approachUsed && spotElevation == null)
                                    failedItems.Add($"Element {elem.Id.Value} in {sectionView.Name}: No spot created after local-direction approach.");
                                else if (!approachUsed)
                                    failedItems.Add($"Element {elem.Id.Value} in {sectionView.Name}: No valid approach attempted.");
                            }
                            catch (Exception ex)
                            {
                                failedItems.Add($"Element error in {sectionView.Name}: {ex.GetType().Name}: {ex.Message}");
                            }
                        }

                        trans.Commit();
                    }
                }
                tg.Assimilate();
            }

            // Write face-match logs to JSON.
            WriteLogsToJson(_faceMatchLogs);

            // Summarize results.
            TaskDialog resultsDialog = new TaskDialog("Spot Elevation Results");
            resultsDialog.MainInstruction = $"Created {successCount} spot elevations across {sectionViews.Count} section views";
            string summaryMsg = "Results by view:\n";
            foreach (var pair in successPerView)
                summaryMsg += $"- {pair.Key}: {pair.Value}\n";
            if (failedItems.Count > 0)
            {
                summaryMsg += "\nFailures/Warnings (first 20):\n";
                int limit = Math.Min(failedItems.Count, 20);
                for (int i = 0; i < limit; i++)
                    summaryMsg += $"- {failedItems[i]}\n";
                if (failedItems.Count > limit)
                    summaryMsg += $"... plus {failedItems.Count - limit} more.\n";
            }
            summaryMsg += "\nIf you don't see spot elevations:\n" +
                          "- Zoom out in the section views\n" +
                          "- Check if Annotations are visible\n" +
                          "- Verify Spot Elevations in Visibility/Graphics";
            resultsDialog.MainContent = summaryMsg;
            resultsDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Show All Section Views in Project Browser");
            TaskDialogResult tdRes = resultsDialog.Show();
            if (tdRes == TaskDialogResult.CommandLink1)
            {
                List<ElementId> viewIds = sectionViews.Select(v => v.Id).ToList();
                uiDoc.ShowElements(viewIds);
            }
        }

        /// <summary>
        /// Implements a local-direction based face search.
        /// Inverts the instance transform and compares each face's local normal with the chosen local direction.
        /// Logs all intermediate evaluation details to _faceMatchLogs.
        /// </summary>
        private Face FindFaceByLocalDirection(Document doc, FamilyInstance fi, View sectionView, XYZ localDir)
        {
            Options opt = new Options
            {
                View = sectionView,
                ComputeReferences = true,
                IncludeNonVisibleObjects = true,
                DetailLevel = ViewDetailLevel.Fine
            };
            GeometryElement geoElem = fi.get_Geometry(opt);
            if (geoElem == null)
            {
                _faceMatchLogs.Add(new FaceMatchLog
                {
                    InstanceId = fi.Id.Value,
                    ViewName = sectionView.Name,
                    BestScore = -999,
                    FaceFound = false,
                    Note = "No geometry found",
                    EvaluationDetails = new List<FaceEvaluationDetail>()
                });
                return null;
            }

            Transform invTransform = fi.GetTransform().Inverse;
            double bestAlignment = -999;
            Face chosenFace = null;
            List<FaceEvaluationDetail> evalDetails = new List<FaceEvaluationDetail>();
            int faceCounter = 0;

            foreach (GeometryObject g in geoElem)
            {
                Solid sol = g as Solid;
                if (sol == null || sol.Volume == 0)
                    continue;
                foreach (Face face in sol.Faces)
                {
                    if (face.Reference == null)
                        continue;
                    try
                    {
                        BoundingBoxUV bb = face.GetBoundingBox();
                        UV midUV = new UV((bb.Min.U + bb.Max.U) * 0.5, (bb.Min.V + bb.Max.V) * 0.5);
                        XYZ normalWorld = face.ComputeNormal(midUV).Normalize();
                        XYZ normalLocal = invTransform.OfVector(normalWorld).Normalize();
                        double alignment = normalLocal.DotProduct(localDir);
                        evalDetails.Add(new FaceEvaluationDetail
                        {
                            FaceIndex = faceCounter,
                            DotProduct = alignment
                        });
                        faceCounter++;
                        if (alignment > bestAlignment)
                        {
                            bestAlignment = alignment;
                            chosenFace = face;
                        }
                    }
                    catch
                    {
                        // Skip errors.
                    }
                }
            }

            _faceMatchLogs.Add(new FaceMatchLog
            {
                InstanceId = fi.Id.Value,
                ViewName = sectionView.Name,
                BestScore = bestAlignment,
                FaceFound = (chosenFace != null),
                Note = (chosenFace != null ? "Success" : "No face found with positive alignment"),
                EvaluationDetails = evalDetails
            });

            return chosenFace;
        }

        /// <summary>
        /// Prompts the user to choose a face direction.
        /// Only four options are available: Left Face, Right Face, Front Face, Back Face.
        /// </summary>
        private string PromptForFaceDirection()
        {
            TaskDialog td = new TaskDialog("Choose Face Direction");
            td.MainInstruction = "Pick a face direction:";
            td.MainContent = "Options: Left Face, Right Face, Front Face, Back Face.";
            td.CommonButtons = TaskDialogCommonButtons.Cancel;
            td.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Left Face");
            td.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Right Face");
            td.AddCommandLink(TaskDialogCommandLinkId.CommandLink3, "Front Face");
            td.AddCommandLink(TaskDialogCommandLinkId.CommandLink4, "Back Face");
            TaskDialogResult result = td.Show();
            if (result == TaskDialogResult.CommandLink1) return "Left Face";
            if (result == TaskDialogResult.CommandLink2) return "Right Face";
            if (result == TaskDialogResult.CommandLink3) return "Front Face";
            if (result == TaskDialogResult.CommandLink4) return "Back Face";
            return null;
        }

        /// <summary>
        /// Prompts the user for leader side (left or right).
        /// </summary>
        private bool PromptForLeaderSide()
        {
            TaskDialog td = new TaskDialog("Leader Direction");
            td.MainInstruction = "Should leaders point LEFT? (Yes) or RIGHT? (No)";
            td.CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No;
            return (td.Show() == TaskDialogResult.Yes);
        }

        /// <summary>
        /// Applies styling (graphic overrides, parameter adjustments) to the spot elevation.
        /// </summary>
        private void StyleSpotElevation(Document doc, View view, SpotDimension spotElevation, List<string> failedItems, string elementInfo)
        {
            try
            {
                OverrideGraphicSettings ogs = new OverrideGraphicSettings();
                ogs.SetProjectionLineWeight(4);
                view.SetElementOverrides(spotElevation.Id, ogs);
                foreach (Parameter p in spotElevation.Parameters)
                {
                    if (p.Definition.Name.Contains("Leader") && !p.IsReadOnly && p.StorageType == StorageType.Double)
                        p.Set(5.0);
                    if (p.Definition.Name.Contains("Text") && !p.IsReadOnly && p.StorageType == StorageType.Double)
                        p.Set(3.0);
                }
            }
            catch (Exception ex)
            {
                failedItems.Add($"Spot styling error on {elementInfo}: {ex.Message}");
            }
        }

        /// <summary>
        /// Writes the face-match logs to a JSON file in the user's Application Data folder.
        /// </summary>
        private void WriteLogsToJson(List<FaceMatchLog> logs)
        {
            try
            {
                string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                string logsDir = Path.Combine(appData, "RevitGPT_Logs");
                if (!Directory.Exists(logsDir))
                    Directory.CreateDirectory(logsDir);
                string logFile = Path.Combine(logsDir, "SpotElevationFacesFound.json");

                // If no logs were generated, add a dummy entry.
                if (logs.Count == 0)
                {
                    logs.Add(new FaceMatchLog
                    {
                        InstanceId = 0,
                        ViewName = "None",
                        BestScore = 0,
                        FaceFound = false,
                        Note = "No face evaluation was performed.",
                        EvaluationDetails = new List<FaceEvaluationDetail>()
                    });
                }

                string json = JsonConvert.SerializeObject(logs, Formatting.Indented);
                File.WriteAllText(logFile, json, System.Text.Encoding.UTF8);
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Logging Error", $"Failed to write logs: {ex.Message}");
            }
        }

        /// <summary>
        /// Data class for logging face matching results.
        /// </summary>
        public class FaceMatchLog
        {
            public long InstanceId { get; set; }
            public string ViewName { get; set; }
            public double BestScore { get; set; }
            public bool FaceFound { get; set; }
            public string Note { get; set; }
            public List<FaceEvaluationDetail> EvaluationDetails { get; set; }
        }

        /// <summary>
        /// Data class for logging evaluation details for each face.
        /// </summary>
        public class FaceEvaluationDetail
        {
            public int FaceIndex { get; set; }
            public double DotProduct { get; set; }
        }

        /// <summary>
        /// Filter that only allows picking references on pre-selected elements.
        /// </summary>
        private class PreselectedElementsFilter : ISelectionFilter
        {
            private ICollection<ElementId> _preselectedIds;
            public PreselectedElementsFilter(ICollection<ElementId> preselectedIds)
            {
                _preselectedIds = preselectedIds;
            }
            public bool AllowElement(Element elem)
            {
                return _preselectedIds.Contains(elem.Id);
            }
            public bool AllowReference(Reference reference, XYZ position)
            {
                return _preselectedIds.Contains(reference.ElementId);
            }
        }
    }
}