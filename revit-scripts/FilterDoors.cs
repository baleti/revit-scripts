﻿using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.ReadOnly)]
public class FilterDoors : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;

        // Get selected elements
        ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();
        if (!selectedIds.Any())
        {
            TaskDialog.Show("Warning", "Please select doors before running the command.");
            return Result.Failed;
        }

        // Filter for doors only
        List<Element> selectedDoors = selectedIds
            .Select(id => doc.GetElement(id))
            .Where(e => e.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Doors)
            .ToList();

        if (!selectedDoors.Any())
        {
            TaskDialog.Show("Warning", "No doors in current selection.");
            return Result.Failed;
        }

        List<Dictionary<string, object>> doorData = new List<Dictionary<string, object>>();
        List<string> propertyNames = new List<string>
        {
            "Family Name",
            "Instance Name",
            "Level",
            "Mark",
            "Head Height",
            "Comments",
            "Group",
            "FacingFlipped",
            "HandFlipped",
            "Width",
            "Height",
            "Room From",
            "Room To"
        };

        foreach (Element door in selectedDoors)
        {
            FamilyInstance doorInst = door as FamilyInstance;
            if (doorInst == null) continue;

            Dictionary<string, object> doorProperties = new Dictionary<string, object>();

            ElementType doorType = doc.GetElement(door.GetTypeId()) as ElementType;

            doorProperties["Family Name"] = doorType?.FamilyName ?? "";
            doorProperties["Instance Name"] = door.Name;
            doorProperties["Level"] = doc.GetElement(door.LevelId)?.Name ?? "";
            doorProperties["Mark"] = door.get_Parameter(BuiltInParameter.ALL_MODEL_MARK)?.AsString() ?? "";
            doorProperties["Head Height"] = door.get_Parameter(BuiltInParameter.INSTANCE_HEAD_HEIGHT_PARAM)?.AsDouble() ?? 0.0;
            doorProperties["Comments"] = door.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)?.AsString() ?? "";
            doorProperties["Group"] = door.GroupId != ElementId.InvalidElementId ? doc.GetElement(door.GroupId)?.Name : "";

            doorProperties["FacingFlipped"] = doorInst.FacingFlipped;
            doorProperties["HandFlipped"] = doorInst.HandFlipped;

            doorProperties["Width"] = doorType?.get_Parameter(BuiltInParameter.DOOR_WIDTH)?.AsDouble() ?? 0.0;
            doorProperties["Height"] = doorType?.get_Parameter(BuiltInParameter.DOOR_HEIGHT)?.AsDouble() ?? 0.0;

            if (doorInst.FromRoom != null)
                doorProperties["Room From"] = doorInst.FromRoom?.Name ?? "";
            if (doorInst.ToRoom != null)
                doorProperties["Room To"] = doorInst.ToRoom?.Name ?? "";

            foreach (Parameter param in door.Parameters)
            {
                if (param.IsShared && !propertyNames.Contains(param.Definition.Name))
                {
                    propertyNames.Add(param.Definition.Name);
                    doorProperties[param.Definition.Name] = param.AsString() ?? "";
                }
            }

            doorData.Add(doorProperties);
        }

        List<Dictionary<string, object>> selectedFromGrid = CustomGUIs.DataGrid(doorData, propertyNames, false);
        
        if (selectedFromGrid != null && selectedFromGrid.Any())
        {
            List<ElementId> finalSelection = selectedDoors
                .Where(d => selectedFromGrid.Any(s => 
                    d.Name == s["Instance Name"].ToString() && 
                    (d as FamilyInstance)?.Symbol.FamilyName == s["Family Name"].ToString()))
                .Select(d => d.Id)
                .ToList();
            
            uidoc.Selection.SetElementIds(finalSelection);
        }

        return Result.Succeeded;
    }
}
