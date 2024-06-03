using DocumentFormat.OpenXml.Wordprocessing;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace TestExecution
{
    public class ExcelToXml
    {
        public void ReadExcel(string filePath)
        {
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                FileInfo fileInfo = new FileInfo(filePath);
                using (ExcelPackage package = new ExcelPackage(fileInfo))
                {
                    ExcelWorkbook workbook = package.Workbook;
                    if (workbook != null)
                    {
                        foreach (ExcelWorksheet sheet in workbook.Worksheets)
                        {
                            if (sheet.Name == "Conditional Tag Expression")
                            {
                                ProcessConditionalTagExpressionWorksheet(sheet);
                            }
                            else if (sheet.Name == "Variables")
                            {
                                ProcessVariablesWorksheet(sheet);
                            }
                        }
                    }
                }
                MessageBox.Show("Excel data converted to XML successfully!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error occurred: " + ex.Message);
            }
        }

        private void ProcessConditionalTagExpressionWorksheet(ExcelWorksheet sheet)
        {
            int rowCount = sheet.Dimension.Rows;
            int colCount = sheet.Dimension.Columns;

            for (int row = 2; row <= rowCount; row++)
            {
                string status = sheet.Cells[row, 1].Value?.ToString();
                string tag = sheet.Cells[row, 2].Value?.ToString();
                string targetFile = sheet.Cells[row, 7].Value?.ToString();

                if (string.IsNullOrEmpty(status) || string.IsNullOrEmpty(tag) || string.IsNullOrEmpty(targetFile))
                    continue;

                ProcessTag(status, tag, targetFile);
            }
        }

        private void ProcessVariablesWorksheet(ExcelWorksheet sheet)
        {
            int rowCount = sheet.Dimension.Rows;
            int colCount = sheet.Dimension.Columns;

            for (int row = 2; row <= rowCount; row++)
            {
                string targetValue = sheet.Cells[row, 1].Value?.ToString();
                string variable = sheet.Cells[row, 3].Value?.ToString();
                string targetFile = sheet.Cells[row, 8].Value?.ToString();

                if (string.IsNullOrEmpty(targetValue) || string.IsNullOrEmpty(variable) || string.IsNullOrEmpty(targetFile))
                    continue;

                ProcessVariable(targetValue, variable, targetFile);
            }
        }

        private void ProcessVariable(string targetValue, string variable, string targetFile)
        {
            XDocument doc = XDocument.Load(targetFile);
            //var xmlAttributes = doc.Descendants("Variable").Select(v => new { Name = v.Attribute("Name")?.Value, Value = v.Value }).ToList();
            var xmlAttributeVariables = doc.Descendants("Variable").ToList();
            // Remove the variable if it exists
            foreach (var variableElement in xmlAttributeVariables)
            {
                if (variableElement.Attribute("Name")?.Value == variable)
                {
                    variableElement.Remove();
                }
            }
            doc.Root.Element("Variables").Add(new XElement("Variable",
                    new XAttribute("Name", variable),
                    targetValue));
            doc.Save(targetFile);
        }

        private void ProcessTag(string status, string tag, string targetFile)
        {
            XmlDocument doc = new XmlDocument();
            doc.Load(targetFile);
            XmlNode node = doc.SelectSingleNode("//CatapultTarget");

            if (node != null && node.Attributes != null)
            {
                string conditionTagExpression = node.Attributes["ConditionTagExpression"]?.Value ?? string.Empty;

                List<string> excludeValues = ExtractValues(conditionTagExpression, "exclude");
                List<string> includeValues = ExtractValues(conditionTagExpression, "include");

                switch (status)
                {
                    case "Exclude":
                        conditionTagExpression = UpdateConditionTagExpression(conditionTagExpression, "exclude", tag, excludeValues, includeValues, addTag: true);
                        conditionTagExpression = UpdateConditionTagExpression(conditionTagExpression, "include", tag, includeValues, excludeValues, addTag: false);
                        break;
                    case "Include":
                        conditionTagExpression = UpdateConditionTagExpression(conditionTagExpression, "exclude", tag, excludeValues, includeValues, addTag: false);
                        conditionTagExpression = UpdateConditionTagExpression(conditionTagExpression, "include", tag, includeValues, excludeValues, addTag: true);
                        break;
                    case "Not Set":
                        conditionTagExpression = UpdateConditionTagExpression(conditionTagExpression, "exclude", tag, excludeValues, includeValues, addTag: false);
                        conditionTagExpression = UpdateConditionTagExpression(conditionTagExpression, "include", tag, includeValues, excludeValues, addTag: false);
                        break;
                }

                node.Attributes["ConditionTagExpression"].Value = conditionTagExpression;
                doc.Save(targetFile);
            }
        }

        private List<string> ExtractValues(string conditionTagExpression, string section)
        {
            string pattern = $@"{section}\[(.*?)\]";
            Match match = Regex.Match(conditionTagExpression, pattern);
            if (match.Success)
            {
                string content = match.Groups[1].Value;
                return Regex.Matches(content, "\"(.*?)\"").Cast<Match>().Select(m => m.Groups[1].Value).ToList();
            }
            return new List<string>();
        }

        private string UpdateConditionTagExpression(string conditionTagExpression, string section, string tag, List<string> currentValues, List<string> oppositeValues, bool addTag)
        {
            if (addTag)
            {
                if (!currentValues.Contains(tag))
                {
                    if (conditionTagExpression.Contains($"{section}["))
                    {
                        int startIndex = conditionTagExpression.IndexOf($"{section}[") + $"{section}[".Length;
                        int endIndex = conditionTagExpression.IndexOf("]", startIndex);
                        string existingValues = conditionTagExpression.Substring(startIndex, endIndex - startIndex).Trim();
                        if (!string.IsNullOrEmpty(existingValues))
                        {
                            existingValues += " or ";
                        }
                        existingValues += $"\"{tag}\"";
                        conditionTagExpression = conditionTagExpression.Substring(0, startIndex) + existingValues + conditionTagExpression.Substring(endIndex);
                    }
                }
            }
            else
            {
                if (currentValues.Contains(tag))
                {
                    if (conditionTagExpression.Contains($"{section}["))
                    {
                        int startIndex = conditionTagExpression.IndexOf($"{section}[") + $"{section}[".Length;
                        int endIndex = conditionTagExpression.IndexOf("]", startIndex);
                        string sectionContent = conditionTagExpression.Substring(startIndex, endIndex - startIndex);
                        var items = sectionContent.Split(new[] { " or " }, StringSplitOptions.None).Where(x => x.Trim() != $"\"{tag}\"").ToList();
                        string newSectionContent = string.Join(" or ", items);
                        conditionTagExpression = conditionTagExpression.Substring(0, startIndex) + newSectionContent + conditionTagExpression.Substring(endIndex);
                    }
                }
            }
            return conditionTagExpression;
        }
    }
}
