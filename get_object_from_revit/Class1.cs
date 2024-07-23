using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using System.Collections.Generic;
using System.Text;
using System.IO;
using Newtonsoft.Json;
using System.Diagnostics;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using Autodesk.Revit.DB.Architecture;

namespace RevitAPI_test
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class MyFirstRevitAPI_Practice : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Start the stopwatch to measure execution time
            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();

            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.UI.Selection.Selection selElement = uidoc.Selection;
            Document doc = uidoc.Document;

            foreach (ElementId elementId in selElement.GetElementIds())
            {
                Element element = doc.GetElement(elementId);
                Dictionary<string, string> elementInfo = new Dictionary<string, string>
                {
                    ["Name"] = element.Name,
                    ["ID"] = element.Id.ToString(),
                    ["Category"] = element.Category?.Name ?? "Unknown Category"
                };

                bool hasElementType = false;

                if (element is FamilyInstance familyInstance)
                {
                    elementInfo["Family"] = familyInstance.Symbol.Family.Name;
                    hasElementType = true;

                    ElementType elementType = doc.GetElement(familyInstance.GetTypeId()) as ElementType;
                    if (elementType != null)
                    {
                        foreach (Parameter param in elementType.Parameters)
                        {
                            string paramName = param.Definition.Name;
                            string paramValue = param.AsValueString() ?? param.AsString();
                            if (!string.IsNullOrEmpty(paramValue))
                            {
                                elementInfo[paramName] = paramValue;
                            }
                        }
                    }
                }
                else if (element is Wall wall)
                {
                    elementInfo["Wall Type"] = wall.GetType().Name;
                    hasElementType = true;
                    foreach (Parameter param in wall.Parameters)
                    {
                        string paramName = param.Definition.Name;
                        string paramValue = param.AsValueString() ?? param.AsString();
                        if (!string.IsNullOrEmpty(paramValue))
                        {
                            elementInfo[paramName] = paramValue;
                        }
                    }
                }
                else if (element is Floor floor)
                {
                    elementInfo["Floor Type"] = floor.GetType().Name;
                    hasElementType = true;
                    foreach (Parameter param in floor.Parameters)
                    {
                        string paramName = param.Definition.Name;
                        string paramValue = param.AsValueString() ?? param.AsString();
                        if (!string.IsNullOrEmpty(paramValue))
                        {
                            elementInfo[paramName] = paramValue;
                        }
                    }
                }
                else if (element is Ceiling ceiling)
                {
                    elementInfo["Ceiling Type"] = ceiling.GetType().Name;
                    hasElementType = true;
                    foreach (Parameter param in ceiling.Parameters)
                    {
                        string paramName = param.Definition.Name;
                        string paramValue = param.AsValueString() ?? param.AsString();
                        if (!string.IsNullOrEmpty(paramValue))
                        {
                            elementInfo[paramName] = paramValue;
                        }
                    }
                }
                else if (element is Stairs stairs)
                {
                    elementInfo["Stairs Type"] = stairs.GetType().Name;
                    hasElementType = true;
                    foreach (Parameter param in stairs.Parameters)
                    {
                        string paramName = param.Definition.Name;
                        string paramValue = param.AsValueString() ?? param.AsString();
                        if (!string.IsNullOrEmpty(paramValue))
                        {
                            elementInfo[paramName] = paramValue;
                        }
                    }
                }
                else if (element.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Railings ||
                         element.Category.Id.IntegerValue == (int)BuiltInCategory.OST_StairsRailing)
                {
                    elementInfo["Type"] = "Handrail/Railing";
                    hasElementType = true;

                    foreach (Parameter param in element.Parameters)
                    {
                        string paramName = param.Definition.Name;
                        string paramValue = param.AsValueString() ?? param.AsString();
                        if (!string.IsNullOrEmpty(paramValue))
                        {
                            elementInfo[paramName] = paramValue;
                        }
                    }
                }

                StringBuilder displayInfo = new StringBuilder();
                foreach (var info in elementInfo)
                {
                    displayInfo.AppendLine($"{info.Key}: {info.Value}");
                }

                // Display element information in a message box
                // MessageBox.Show(displayInfo.ToString(), "Element Information");

                string json = JsonConvert.SerializeObject(elementInfo, Formatting.Indented);
                // Replace illegal characters in category name
                string sanitizedCategory = Regex.Replace(elementInfo["Category"], @"[<>:""/\\|?*]", "_");

                if (hasElementType)
                {
                    // Specify custom file save path
                    string fileName = $"{sanitizedCategory}__{elementInfo["ID"]}.json";
                    string customFilePath = Path.Combine(@"E:\test", fileName);

                    // Ensure the directory exists, create if not
                    string directoryPath = Path.GetDirectoryName(customFilePath);
                    if (!Directory.Exists(directoryPath))
                    {
                        Directory.CreateDirectory(directoryPath);
                    }
                    File.WriteAllText(customFilePath, json);
                }
                else
                {
                    // Specify custom file save path
                    string fileName = $"Fail(X)_{sanitizedCategory}__{elementInfo["ID"]}.json";
                    string customFilePath = Path.Combine(@"E:\test", fileName);

                    // Ensure the directory exists, create if not
                    string directoryPath = Path.GetDirectoryName(customFilePath);
                    if (!Directory.Exists(directoryPath))
                    {
                        Directory.CreateDirectory(directoryPath);
                    }
                    File.WriteAllText(customFilePath, json);
                }
            }

            // Stop the stopwatch
            stopwatch.Stop();

            // Show the elapsed time in a message box
            MessageBox.Show($"Execution time: {stopwatch.Elapsed.TotalSeconds} seconds");

            return Result.Succeeded;
        }
    }
}
