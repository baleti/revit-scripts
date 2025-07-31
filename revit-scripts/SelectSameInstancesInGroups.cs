using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class SelectSameInstancesInGroups : IExternalCommand
{
    // Structure to hold element info with relative centroid
    private class ElementInfo
    {
        public XYZ RelativeCentroid { get; set; }
        public string Signature { get; set; }
        public ElementId OriginalId { get; set; }
        public string Description { get; set; }
    }
    
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        try
        {
            var uidoc = commandData.Application.ActiveUIDocument;
            var doc = uidoc.Document;
            
            // Get currently selected elements using the extension method from SelectionModeManager
            var selectedIds = uidoc.GetSelectionIds();
            
            if (!selectedIds.Any())
            {
                TaskDialog.Show("Selection", "Please select at least one element that belongs to a group.");
                return Result.Cancelled;
            }
            
            // Dictionary to map GroupType to list of element info (relative centroids and signatures)
            var groupTypeToElementInfos = new Dictionary<ElementId, List<ElementInfo>>();
            
            // Dictionary to cache group instances by type
            var groupTypeToInstances = new Dictionary<ElementId, List<Group>>();
            
            // Process selected elements
            foreach (var selectedId in selectedIds)
            {
                var element = doc.GetElement(selectedId);
                if (element == null) continue;
                
                // Skip if not in a group
                if (element.GroupId == ElementId.InvalidElementId) continue;
                
                var group = doc.GetElement(element.GroupId) as Group;
                if (group == null) continue;
                
                var groupTypeId = group.GetTypeId();
                
                // Get element centroid relative to group origin
                var groupOrigin = GetGroupOrigin(group);
                var elementCentroid = GetElementCentroid(element);
                
                if (elementCentroid == null || groupOrigin == null) continue;
                
                // Calculate relative position
                var relativeCentroid = elementCentroid - groupOrigin;
                
                // Create element info
                var elementInfo = new ElementInfo
                {
                    RelativeCentroid = relativeCentroid,
                    Signature = CreateElementSignature(element),
                    OriginalId = selectedId,
                    Description = GetElementDescription(element)
                };
                
                // Store element info for this group type
                if (!groupTypeToElementInfos.ContainsKey(groupTypeId))
                {
                    groupTypeToElementInfos[groupTypeId] = new List<ElementInfo>();
                }
                
                // Check if we already have this element (avoid duplicates)
                bool alreadyAdded = groupTypeToElementInfos[groupTypeId].Any(ei => 
                    ArePointsEqual(ei.RelativeCentroid, elementInfo.RelativeCentroid) &&
                    ei.Signature == elementInfo.Signature);
                
                if (!alreadyAdded)
                {
                    groupTypeToElementInfos[groupTypeId].Add(elementInfo);
                }
                
                // Cache group instances for this type
                if (!groupTypeToInstances.ContainsKey(groupTypeId))
                {
                    groupTypeToInstances[groupTypeId] = new FilteredElementCollector(doc)
                        .OfClass(typeof(Group))
                        .Cast<Group>()
                        .Where(g => g.GetTypeId() == groupTypeId)
                        .ToList();
                }
            }
            
            if (!groupTypeToElementInfos.Any())
            {
                TaskDialog.Show("Selection", "No selected elements belong to groups.");
                return Result.Cancelled;
            }
            
            // Collect all matching elements
            var allElementIds = new HashSet<ElementId>(selectedIds);
            var diagnostics = new StringBuilder();
            int totalMatches = 0;
            int totalMismatches = 0;
            
            // For each group type with selected elements
            foreach (var kvp in groupTypeToElementInfos)
            {
                var groupTypeId = kvp.Key;
                var elementInfos = kvp.Value;
                var groupInstances = groupTypeToInstances[groupTypeId];
                var groupTypeName = doc.GetElement(groupTypeId).Name;
                
                // For each group instance of this type
                foreach (var groupInstance in groupInstances)
                {
                    var groupOrigin = GetGroupOrigin(groupInstance);
                    if (groupOrigin == null) continue;
                    
                    var memberIds = groupInstance.GetMemberIds();
                    var groupDiagnostics = new StringBuilder();
                    bool hasIssues = false;
                    
                    // For each element info (relative centroid to find)
                    foreach (var targetInfo in elementInfos)
                    {
                        // Calculate expected world position
                        var expectedWorldPosition = groupOrigin + targetInfo.RelativeCentroid;
                        
                        // Find element at this position
                        ElementId foundId = null;
                        Element foundElement = null;
                        double closestDistance = double.MaxValue;
                        
                        foreach (var memberId in memberIds)
                        {
                            var memberElement = doc.GetElement(memberId);
                            if (memberElement == null) continue;
                            
                            var memberCentroid = GetElementCentroid(memberElement);
                            if (memberCentroid == null) continue;
                            
                            var distance = memberCentroid.DistanceTo(expectedWorldPosition);
                            if (distance < closestDistance && distance < 0.01) // Within ~3mm tolerance
                            {
                                closestDistance = distance;
                                foundId = memberId;
                                foundElement = memberElement;
                            }
                        }
                        
                        if (foundId != null)
                        {
                            allElementIds.Add(foundId);
                            totalMatches++;
                            
                            // Check if signature matches
                            var foundSignature = CreateElementSignature(foundElement);
                            if (foundSignature != targetInfo.Signature)
                            {
                                hasIssues = true;
                                groupDiagnostics.AppendLine($"  - Position {FormatPoint(targetInfo.RelativeCentroid)}: Expected {targetInfo.Description}, Found {GetElementDescription(foundElement)}");
                                totalMismatches++;
                            }
                        }
                        else
                        {
                            hasIssues = true;
                            groupDiagnostics.AppendLine($"  - Position {FormatPoint(targetInfo.RelativeCentroid)}: No element found (Expected {targetInfo.Description})");
                        }
                    }
                    
                    if (hasIssues)
                    {
                        diagnostics.AppendLine($"\nGroup Instance: {groupInstance.Name} (Id: {groupInstance.Id})");
                        diagnostics.AppendLine($"Group Type: {groupTypeName}");
                        diagnostics.Append(groupDiagnostics.ToString());
                    }
                }
            }
            
            // Set the new selection
            uidoc.SetSelectionIds(allElementIds.ToList());
            
            // Show diagnostics
            var summary = new StringBuilder();
            summary.AppendLine($"Selected {allElementIds.Count} elements total.");
            summary.AppendLine($"Successfully matched {totalMatches} element positions across all groups.");
            
            if (totalMismatches > 0 || diagnostics.Length > 0)
            {
                summary.AppendLine($"\nFound {totalMismatches} type mismatches where elements at the same position have different types.");
                if (diagnostics.Length > 0)
                {
                    summary.AppendLine("\nDetailed diagnostics:");
                    summary.Append(diagnostics.ToString());
                }
                
                summary.AppendLine("\n--- Note ---");
                summary.AppendLine("Elements were selected based on their position (centroid) within the group.");
                summary.AppendLine("Type mismatches indicate that elements at the same position have been replaced or modified.");
            }
            else
            {
                summary.AppendLine("All elements matched expected types at their positions.");
            }
            
            // Use a custom dialog for longer text
            var td = new TaskDialog("Selection Diagnostics")
            {
                MainContent = summary.ToString(),
                MainIcon = totalMismatches > 0 ? TaskDialogIcon.TaskDialogIconWarning : TaskDialogIcon.TaskDialogIconInformation,
                CommonButtons = TaskDialogCommonButtons.Ok,
                TitleAutoPrefix = false
            };
            td.Show();
            
            return Result.Succeeded;
        }
        catch (Exception ex)
        {
            message = ex.Message;
            return Result.Failed;
        }
    }
    
    private XYZ GetGroupOrigin(Group group)
    {
        // Groups don't have a direct origin, but we can use the bounding box center
        // or the location point if it exists
        if (group.Location is LocationPoint locPoint)
        {
            return locPoint.Point;
        }
        
        // Fallback to bounding box center
        var bb = group.get_BoundingBox(null);
        if (bb != null)
        {
            return (bb.Min + bb.Max) * 0.5;
        }
        
        return XYZ.Zero;
    }
    
    private XYZ GetElementCentroid(Element element)
    {
        // Try to get location point first (for family instances)
        if (element.Location is LocationPoint locPoint)
        {
            return locPoint.Point;
        }
        
        // For linear elements, get midpoint
        if (element.Location is LocationCurve locCurve)
        {
            var curve = locCurve.Curve;
            return curve.Evaluate(0.5, true);
        }
        
        // For other elements, use bounding box center
        var bb = element.get_BoundingBox(null);
        if (bb != null)
        {
            return (bb.Min + bb.Max) * 0.5;
        }
        
        return null;
    }
    
    private string CreateElementSignature(Element element)
    {
        var categoryId = element.Category?.Id.IntegerValue ?? -1;
        var typeId = element.GetTypeId().IntegerValue;
        return $"{categoryId}_{typeId}";
    }
    
    private string GetElementDescription(Element element)
    {
        var category = element.Category?.Name ?? "Unknown Category";
        var typeName = "Unknown Type";
        
        var typeElement = element.Document.GetElement(element.GetTypeId());
        if (typeElement != null)
        {
            typeName = typeElement.Name;
        }
        
        return $"{category}: {typeName}";
    }
    
    private bool ArePointsEqual(XYZ p1, XYZ p2, double tolerance = 0.001)
    {
        if (p1 == null || p2 == null) return false;
        return p1.DistanceTo(p2) < tolerance;
    }
    
    private string FormatPoint(XYZ point)
    {
        return $"({point.X:F2}, {point.Y:F2}, {point.Z:F2})";
    }
}
