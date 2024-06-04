using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using LicenseContext = OfficeOpenXml.LicenseContext;

namespace TestExecution
{
    public partial class TestExecutionForm : Form
    {
        public TestExecutionForm()
        {
            InitializeComponent();
            progressBar1.Visible = false;
            lblStatus.Visible = false;
        }
        private Task ProcessData(List<string> list,IProgress<ProgessReport> progress)
        {
            int index = 1;
            int totalProcess = list.Count;
            var progressReport = new ProgessReport();
            return Task.Run(() =>
            {
                for(int i = 0; i < totalProcess; i++)
                {
                    progressReport.PercentComplete = index++ * 100 / totalProcess;
                    progress.Report(progressReport);
                    Thread.Sleep(10);
                }
            });
        }
        private async void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "Select File to Upload";
            openFileDialog.Filter = "All Files|*.*";
            DialogResult result = openFileDialog.ShowDialog();
            if (result == DialogResult.OK)
            {
                string fileName = openFileDialog.FileName;
                ExcelToXml excelToXml = new ExcelToXml();
                excelToXml.ReadExcel(fileName);
            }
            buttonUpdateExcel.Enabled = false;
            buttonUpdateXml.Enabled = false;
            progressBar1.Visible = true;
            lblStatus.Visible = true;
            List<string> list = new List<string>();
            for (int i = 0; i < 1000; i++)
                list.Add(i.ToString());
            lblStatus.Text = "Working...";
            var progress = new Progress<ProgessReport>();
            progress.ProgressChanged += (o, report) =>
            {
                lblStatus.Text = string.Format("Processing...{0}%", report.PercentComplete);
                progressBar1.Value = report.PercentComplete;
                progressBar1.Update();
            };
            await ProcessData(list, progress);
            MessageBox.Show("Excel data converted to XML successfully!");
            progressBar1.Visible = false;
            buttonUpdateExcel.Enabled = true;
            buttonUpdateXml.Enabled = true;
            lblStatus.Visible = false;         
        }
        //private void ReadExcel(string filePath)
        //{
        //    try
        //    {
        //        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        //        FileInfo fileInfo = new FileInfo(filePath);
        //        using (ExcelPackage package = new ExcelPackage(fileInfo))
        //        {
        //            ExcelWorkbook workbook = package.Workbook;
        //            if (workbook != null)
        //            {
        //                foreach (ExcelWorksheet sheet in workbook.Worksheets)
        //                {
        //                    if (sheet.Name == "Conditional Tag Expression")
        //                    {
        //                        int rowCount = sheet.Dimension.Rows;
        //                        int colCount = sheet.Dimension.Columns;
        //                        for (int col = 1; col <= colCount - (colCount - 1); col++)
        //                        {
        //                            for (int row = 2; row <= rowCount; row++)
        //                            {
        //                                string status = sheet.Cells[row, col].Value?.ToString();
        //                                string tag = sheet.Cells[row, col + 1].Value?.ToString();
        //                                string comment = sheet.Cells[row, col + 2].Value?.ToString();
        //                                string change = sheet.Cells[row, col + 3].Value?.ToString();
        //                                string twNotes = sheet.Cells[row, col + 4].Value?.ToString();
        //                                string targetLanguage = sheet.Cells[row, col + 5].Value?.ToString();
        //                                string targetFile = sheet.Cells[row, col + 6].Value?.ToString();
        //                                string baseFolder = sheet.Cells[row, col + 7].Value?.ToString();
        //                                string outputFile = sheet.Cells[row, col + 8].Value?.ToString();
        //                                string type = sheet.Cells[row, col + 9].Value?.ToString();

        //                                XmlDocument doc = new XmlDocument();
        //                                doc.Load(targetFile);
        //                                XmlNode node = doc.SelectSingleNode("//CatapultTarget");

        //                                if (node != null && node.Attributes != null)
        //                                {
        //                                    string attributeValue = node.Attributes["ConditionTagExpression"]?.Value;

        //                                    /// <summary>
        //                                    // Extract exclude values
        //                                    /// <summary>
        //                                    /// 
        //                                    string excludePattern = @"exclude\[(.*?)\]";
        //                                    Match excludeMatch = Regex.Match(attributeValue, excludePattern);
        //                                    List<string> excludeValues = new List<string>();
        //                                    if (excludeMatch.Success)
        //                                    {
        //                                        string excludeContent = excludeMatch.Groups[1].Value;
        //                                        excludeValues = Regex.Matches(excludeContent, "\"(.*?)\"")
        //                                                            .Cast<Match>()
        //                                                            .Select(m => m.Groups[1].Value)
        //                                                            .ToList();
        //                                    }
        //                                    /// <summary>
        //                                    // Extract include values
        //                                    /// <summary>
        //                                    string includePattern = @"include\[(.*?)\]";
        //                                    Match includeMatch = Regex.Match(attributeValue, includePattern);
        //                                    List<string> includeValues = new List<string>();
        //                                    if (includeMatch.Success)
        //                                    {
        //                                        string includeContent = includeMatch.Groups[1].Value;
        //                                        includeValues = Regex.Matches(includeContent, "\"(.*?)\"")
        //                                                            .Cast<Match>()
        //                                                            .Select(m => m.Groups[1].Value)
        //                                                            .ToList();
        //                                    }

        //                                    XmlElement catapultTarget = (XmlElement)node;
        //                                    string conditionTagExpression = catapultTarget.GetAttribute("ConditionTagExpression");

        //                                    if (status == "Exclude")
        //                                    {
        //                                        /// <summary>
        //                                        // Add tag to exclude list if not present
        //                                        /// <summary>
        //                                        if (!excludeValues.Contains(tag))
        //                                        {
        //                                            if (conditionTagExpression.Contains("exclude["))
        //                                            {
        //                                                int excludeStartIndex = conditionTagExpression.IndexOf("exclude[") + "exclude[".Length;
        //                                                int excludeEndIndex = conditionTagExpression.IndexOf("]", excludeStartIndex);
        //                                                string existingExcludes = conditionTagExpression.Substring(excludeStartIndex, excludeEndIndex - excludeStartIndex).Trim();
        //                                                if (!string.IsNullOrEmpty(existingExcludes))
        //                                                {
        //                                                    existingExcludes += " or ";
        //                                                }
        //                                                existingExcludes += $"\"{tag}\"";
        //                                                conditionTagExpression = conditionTagExpression.Substring(0, excludeStartIndex) + existingExcludes + conditionTagExpression.Substring(excludeEndIndex);
        //                                            }
        //                                        }
        //                                        /// <summary>
        //                                        // Remove tag from include list if present
        //                                        /// <summary>
        //                                        if (includeValues.Contains(tag))
        //                                        {
        //                                            if (conditionTagExpression.Contains("include["))
        //                                            {
        //                                                int includeStartIndex = conditionTagExpression.IndexOf("include[") + "include[".Length;
        //                                                int includeEndIndex = conditionTagExpression.IndexOf("]", includeStartIndex);
        //                                                string includeSection = conditionTagExpression.Substring(includeStartIndex, includeEndIndex - includeStartIndex);
        //                                                var includeItems = includeSection.Split(new[] { " or " }, StringSplitOptions.None)
        //                                                                                .Where(x => x.Trim() != $"\"{tag}\"")
        //                                                                                .ToList();
        //                                                string newIncludeSection = string.Join(" or ", includeItems);
        //                                                conditionTagExpression = conditionTagExpression.Substring(0, includeStartIndex) + newIncludeSection + conditionTagExpression.Substring(includeEndIndex);
        //                                            }
        //                                        }
        //                                    }
        //                                    else if (status == "Include")
        //                                    {
        //                                        /// <summary>
        //                                        // Remove tag from exclude list if present
        //                                        /// <summary>
        //                                        if (excludeValues.Contains(tag))
        //                                        {
        //                                            if (conditionTagExpression.Contains("exclude["))
        //                                            {
        //                                                int excludeStartIndex = conditionTagExpression.IndexOf("exclude[") + "exclude[".Length;
        //                                                int excludeEndIndex = conditionTagExpression.IndexOf("]", excludeStartIndex);
        //                                                string excludeSection = conditionTagExpression.Substring(excludeStartIndex, excludeEndIndex - excludeStartIndex);
        //                                                var excludeItems = excludeSection.Split(new[] { " or " }, StringSplitOptions.None)
        //                                                                                .Where(x => x.Trim() != $"\"{tag}\"")
        //                                                                                .ToList();
        //                                                string newExcludeSection = string.Join(" or ", excludeItems);
        //                                                conditionTagExpression = conditionTagExpression.Substring(0, excludeStartIndex) + newExcludeSection + conditionTagExpression.Substring(excludeEndIndex);
        //                                            }
        //                                        }
        //                                        /// <summary>
        //                                        // Add tag to include list if not present
        //                                        /// <summary>
        //                                        if (!includeValues.Contains(tag))
        //                                        {
        //                                            if (conditionTagExpression.Contains("include["))
        //                                            {
        //                                                int includeStartIndex = conditionTagExpression.IndexOf("include[") + "include[".Length;
        //                                                int includeEndIndex = conditionTagExpression.IndexOf("]", includeStartIndex);
        //                                                string existingIncludes = conditionTagExpression.Substring(includeStartIndex, includeEndIndex - includeStartIndex).Trim();
        //                                                if (!string.IsNullOrEmpty(existingIncludes))
        //                                                {
        //                                                    existingIncludes += " or ";
        //                                                }
        //                                                existingIncludes += $"\"{tag}\"";
        //                                                conditionTagExpression = conditionTagExpression.Substring(0, includeStartIndex) + existingIncludes + conditionTagExpression.Substring(includeEndIndex);
        //                                            }
        //                                        }
        //                                    }
        //                                    else if (status == "Not Set")
        //                                    {
        //                                        /// <summary>
        //                                        // Remove tag from exclude list if present
        //                                        /// <summary>
        //                                        if (excludeValues.Contains(tag))
        //                                        {
        //                                            if (conditionTagExpression.Contains("exclude["))
        //                                            {
        //                                                int excludeStartIndex = conditionTagExpression.IndexOf("exclude[") + "exclude[".Length;
        //                                                int excludeEndIndex = conditionTagExpression.IndexOf("]", excludeStartIndex);
        //                                                string excludeSection = conditionTagExpression.Substring(excludeStartIndex, excludeEndIndex - excludeStartIndex);
        //                                                var excludeItems = excludeSection.Split(new[] { " or " }, StringSplitOptions.None)
        //                                                                                .Where(x => x.Trim() != $"\"{tag}\"")
        //                                                                                .ToList();
        //                                                string newExcludeSection = string.Join(" or ", excludeItems);
        //                                                conditionTagExpression = conditionTagExpression.Substring(0, excludeStartIndex) + newExcludeSection + conditionTagExpression.Substring(excludeEndIndex);
        //                                            }
        //                                        }
        //                                        /// <summary>
        //                                        // Remove tag from include list if present
        //                                        /// <summary>
        //                                        if (includeValues.Contains(tag))
        //                                        {
        //                                            if (conditionTagExpression.Contains("include["))
        //                                            {
        //                                                int includeStartIndex = conditionTagExpression.IndexOf("include[") + "include[".Length;
        //                                                int includeEndIndex = conditionTagExpression.IndexOf("]", includeStartIndex);
        //                                                string includeSection = conditionTagExpression.Substring(includeStartIndex, includeEndIndex - includeStartIndex);
        //                                                var includeItems = includeSection.Split(new[] { " or " }, StringSplitOptions.None)
        //                                                                                .Where(x => x.Trim() != $"\"{tag}\"")
        //                                                                                .ToList();
        //                                                string newIncludeSection = string.Join(" or ", includeItems);
        //                                                conditionTagExpression = conditionTagExpression.Substring(0, includeStartIndex) + newIncludeSection + conditionTagExpression.Substring(includeEndIndex);
        //                                            }
        //                                        }
        //                                    }

        //                                    catapultTarget.SetAttribute("ConditionTagExpression", conditionTagExpression);
        //                                    doc.Save(targetFile);
        //                                }
        //                            }
        //                        }
        //                    }
        //                }
        //            }
        //        }
        //        MessageBox.Show("Excel data converted to XML successfully!");
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show("Error occurred: " + ex.Message);
        //    }
        //}
    }
}
