﻿using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using LicenseContext = OfficeOpenXml.LicenseContext;

namespace TestExecution
{
    public partial class Form1 : Form
    {
        private Dictionary<string, object> viewbag = new Dictionary<string, object>();
        public Form1()
        {
            InitializeComponent();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "Select File to Upload";
            openFileDialog.Filter = "All Files|*.*";
            DialogResult result = openFileDialog.ShowDialog();
            if (result == DialogResult.OK)
            {
                string fileName = openFileDialog.FileName;
            }
            ReadExcel(openFileDialog.FileName);
        }     
        private void ReadExcel(string filePath)
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
                                int rowCount = sheet.Dimension.Rows;
                                int colCount = sheet.Dimension.Columns;
                                for (int col = 1; col <= colCount - (colCount - 1); col++)
                                {
                                    for (int row = 2; row <= rowCount; row++)
                                    {
                                        string status = sheet.Cells[row, col].Value?.ToString();
                                        string tag = sheet.Cells[row, col + 1].Value?.ToString();
                                        string comment = sheet.Cells[row, col + 2].Value?.ToString();
                                        string change = sheet.Cells[row, col + 3].Value?.ToString();
                                        string twNotes = sheet.Cells[row, col + 4].Value?.ToString();
                                        string targetLanguage = sheet.Cells[row, col + 5].Value?.ToString();
                                        string targetFile = sheet.Cells[row, col + 6].Value?.ToString();
                                        string baseFolder = sheet.Cells[row, col + 7].Value?.ToString();
                                        string outputFile = sheet.Cells[row, col + 8].Value?.ToString();
                                        string type = sheet.Cells[row, col + 9].Value?.ToString();

                                        XmlDocument doc = new XmlDocument();
                                        doc.Load(targetFile);
                                        XmlNode node = doc.SelectSingleNode("//CatapultTarget");

                                        if (node != null && node.Attributes != null)
                                        {
                                            string attributeValue = node.Attributes["ConditionTagExpression"]?.Value;

                                            // Extract exclude values
                                            string excludePattern = @"exclude\[(.*?)\]";
                                            Match excludeMatch = Regex.Match(attributeValue, excludePattern);
                                            List<string> excludeValues = new List<string>();
                                            if (excludeMatch.Success)
                                            {
                                                string excludeContent = excludeMatch.Groups[1].Value;
                                                excludeValues = Regex.Matches(excludeContent, "\"(.*?)\"")
                                                                    .Cast<Match>()
                                                                    .Select(m => m.Groups[1].Value)
                                                                    .ToList();
                                            }

                                            // Extract include values
                                            string includePattern = @"include\[(.*?)\]";
                                            Match includeMatch = Regex.Match(attributeValue, includePattern);
                                            List<string> includeValues = new List<string>();
                                            if (includeMatch.Success)
                                            {
                                                string includeContent = includeMatch.Groups[1].Value;
                                                includeValues = Regex.Matches(includeContent, "\"(.*?)\"")
                                                                    .Cast<Match>()
                                                                    .Select(m => m.Groups[1].Value)
                                                                    .ToList();
                                            }

                                            XmlElement catapultTarget = (XmlElement)node;
                                            string conditionTagExpression = catapultTarget.GetAttribute("ConditionTagExpression");

                                            if (status == "Exclude")
                                            {
                                                // Add tag to exclude list if not present
                                                if (!excludeValues.Contains(tag))
                                                {
                                                    if (conditionTagExpression.Contains("exclude["))
                                                    {
                                                        int excludeStartIndex = conditionTagExpression.IndexOf("exclude[") + "exclude[".Length;
                                                        int excludeEndIndex = conditionTagExpression.IndexOf("]", excludeStartIndex);
                                                        string existingExcludes = conditionTagExpression.Substring(excludeStartIndex, excludeEndIndex - excludeStartIndex).Trim();
                                                        if (!string.IsNullOrEmpty(existingExcludes))
                                                        {
                                                            existingExcludes += " or ";
                                                        }
                                                        existingExcludes += $"\"{tag}\"";
                                                        conditionTagExpression = conditionTagExpression.Substring(0, excludeStartIndex) + existingExcludes + conditionTagExpression.Substring(excludeEndIndex);
                                                    }
                                                }

                                                // Remove tag from include list if present
                                                if (includeValues.Contains(tag))
                                                {
                                                    if (conditionTagExpression.Contains("include["))
                                                    {
                                                        int includeStartIndex = conditionTagExpression.IndexOf("include[") + "include[".Length;
                                                        int includeEndIndex = conditionTagExpression.IndexOf("]", includeStartIndex);
                                                        string includeSection = conditionTagExpression.Substring(includeStartIndex, includeEndIndex - includeStartIndex);
                                                        var includeItems = includeSection.Split(new[] { " or " }, StringSplitOptions.None)
                                                                                        .Where(x => x.Trim() != $"\"{tag}\"")
                                                                                        .ToList();
                                                        string newIncludeSection = string.Join(" or ", includeItems);
                                                        conditionTagExpression = conditionTagExpression.Substring(0, includeStartIndex) + newIncludeSection + conditionTagExpression.Substring(includeEndIndex);
                                                    }
                                                }
                                            }
                                            else if (status == "Include")
                                            {
                                                // Remove tag from exclude list if present
                                                if (excludeValues.Contains(tag))
                                                {
                                                    if (conditionTagExpression.Contains("exclude["))
                                                    {
                                                        int excludeStartIndex = conditionTagExpression.IndexOf("exclude[") + "exclude[".Length;
                                                        int excludeEndIndex = conditionTagExpression.IndexOf("]", excludeStartIndex);
                                                        string excludeSection = conditionTagExpression.Substring(excludeStartIndex, excludeEndIndex - excludeStartIndex);
                                                        var excludeItems = excludeSection.Split(new[] { " or " }, StringSplitOptions.None)
                                                                                        .Where(x => x.Trim() != $"\"{tag}\"")
                                                                                        .ToList();
                                                        string newExcludeSection = string.Join(" or ", excludeItems);
                                                        conditionTagExpression = conditionTagExpression.Substring(0, excludeStartIndex) + newExcludeSection + conditionTagExpression.Substring(excludeEndIndex);
                                                    }
                                                }
                                                // Add tag to include list if not present
                                                if (!includeValues.Contains(tag))
                                                {
                                                    if (conditionTagExpression.Contains("include["))
                                                    {
                                                        int includeStartIndex = conditionTagExpression.IndexOf("include[") + "include[".Length;
                                                        int includeEndIndex = conditionTagExpression.IndexOf("]", includeStartIndex);
                                                        string existingIncludes = conditionTagExpression.Substring(includeStartIndex, includeEndIndex - includeStartIndex).Trim();
                                                        if (!string.IsNullOrEmpty(existingIncludes))
                                                        {
                                                            existingIncludes += " or ";
                                                        }
                                                        existingIncludes += $"\"{tag}\"";
                                                        conditionTagExpression = conditionTagExpression.Substring(0, includeStartIndex) + existingIncludes + conditionTagExpression.Substring(includeEndIndex);
                                                    }
                                                }
                                            }

                                            catapultTarget.SetAttribute("ConditionTagExpression", conditionTagExpression);
                                            doc.Save(targetFile);
                                        }
                                    }
                                }
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
    }
}