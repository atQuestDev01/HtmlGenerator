using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace HtmlGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            string templatePath = "";
            string excelFile = "";
            string resultFile = "";

            if (args.Length > 0)
            {
                string configFile = args[0];
                string[] parameters = File.ReadAllText(configFile).Split(Environment.NewLine);

                templatePath = parameters[0].Trim();
                excelFile = parameters[1].Trim();
                resultFile = parameters[2].Trim();
            }

            List<Template> templates = GetTemplates(templatePath);
            FileInfo fi = new FileInfo(excelFile);

            using (ExcelPackage ep = new ExcelPackage(fi))
            {
                foreach (ExcelNamedRange currentNameRange in ep.Workbook.Names)
                {
                    List<Marker> markers = new List<Marker>();

                    string htmlResult = "";
                    CellAddress currentNameRangeAddress = ResolveAddress(currentNameRange.FullAddressAbsolute);
                    ExcelWorksheet currentWorksheet = currentNameRange.Worksheet;

                    object[,] cellInfo = (object[,])currentNameRange.Value;
                    string currentRow = "";

                    for (int row = 0; row < currentNameRange.Rows; row++)
                    {
                        currentRow += "<tr>" + Environment.NewLine;

                        #region Format the columns
                        string currentRowColumns = "";
                        string colValue = "";
                        bool nextCol = true;
                        int iColSpan = 1;

                        for (int col = 0; col < currentNameRange.Columns; col++)
                        {
                            CellAddress currentCellAddress = new CellAddress()
                            {
                                FromRow = currentNameRangeAddress.FromRow + row,
                                FromCol = currentNameRangeAddress.FromCol + col,
                                ToRow = currentNameRangeAddress.FromRow + row,
                                ToCol = currentNameRangeAddress.FromCol + col
                            };

                            ExcelRange currentCell = currentWorksheet.Cells[currentCellAddress.FromRow, currentCellAddress.FromCol];

                            do
                            {
                                colValue += cellInfo[row, col] + "";
                                if (col < currentNameRange.Columns - 1)
                                {
                                    if (cellInfo[row, col + 1] == null)
                                    {
                                        iColSpan++;
                                        col++;
                                    }
                                    else
                                    {
                                        nextCol = false;
                                    }
                                }
                                else
                                {
                                    nextCol = false;
                                }

                            } while (nextCol);

                            Marker markerType = null;
                            if (IsMarker(templates, colValue, out markerType))
                            {
                                markers.Add(markerType);
                            }

                            currentRowColumns += string.Format("<td colspan=\"{0}\" style=\"{1}\">{2}</td>", iColSpan, ResolveStyle(currentCell), colValue);
                            colValue = "";
                            iColSpan = 1;                            }
                        #endregion

                        currentRow += currentRowColumns + Environment.NewLine;
                        currentRow += "</tr>" + Environment.NewLine;
                    }

                    string currentTable = "";
                    currentTable += "<table id=\"formTable\">" + Environment.NewLine;
                    currentTable += currentRow + Environment.NewLine;
                    currentTable += "</table>" + Environment.NewLine;

                    string submitButton = GetHtmlControl(templatePath);
                    htmlResult = currentTable + Environment.NewLine + 
                                 submitButton + Environment.NewLine;

                    foreach (Marker marker in markers)
                    {
                        Template currentTemplate = templates.Where(m => m.MarkerType == marker.Type).FirstOrDefault();
                        if (currentTemplate != null)
                        {
                            htmlResult = htmlResult.Replace(marker.Value, currentTemplate.Content.Replace("_NAME_", marker.Id));
                        }
                    }
                    
                    string outputFilename = resultFile.Replace(".html", currentWorksheet.Name.Replace(".", "_") + ".html");
                    string javascript = GetScript(templatePath);
                    string htmlTemplate = GetHtmlBody(templatePath);

                    GenerateHtml(outputFilename, htmlTemplate, javascript, htmlResult);
                }
            }
        }


        static bool IsMarker(List<Template> templates, string value, out Marker markerType)
        {
            bool returnValue = false;
            markerType = null;

            foreach (Template currentTemplate in templates)
            {
                if (value.StartsWith(currentTemplate.MarkerType, StringComparison.InvariantCultureIgnoreCase))
                {
                    markerType = new Marker()
                    {
                        Value = value,
                        Type = currentTemplate.MarkerType,
                        Id = value.Substring(value.IndexOf("#") + 1).Trim()
                    };
                    returnValue = true;
                    break;
                }
            }
            return returnValue;
        }

        static List<Template> GetTemplates(string templateFolder)
        {
            List<Template> returnValue = new List<Template>();
            string[] templateFiles = Directory.GetFiles(templateFolder, "*.txt");

            for (int i = 0; i < templateFiles.Length; i++)
            {
                string fullPath = templateFiles[i];
                string filename = fullPath;
                filename = filename.Replace(templateFolder, "");

                if (filename.StartsWith("_") && filename.EndsWith("_.txt"))
                {
                    returnValue.Add(new Template()
                    {
                        MarkerType = filename.Replace(".txt", ""),
                        Content = File.ReadAllText(fullPath)
                    });
                }
            }

            return returnValue;
        }

        static string GetScript(string templateFolder)
        {
            string returnValue = "";
            string[] templateFiles = Directory.GetFiles(templateFolder, "scripts.js");

            if (templateFiles.Length > 0)
            {
                returnValue = File.ReadAllText(templateFiles[0]);
            }

            return returnValue;
        }

        static string GetHtmlBody(string templateFolder)
        {
            string returnValue = "";
            string[] templateFiles = Directory.GetFiles(templateFolder, "templateBody.html");

            if (templateFiles.Length > 0)
            {
                returnValue = File.ReadAllText(templateFiles[0]);
            }

            return returnValue;
        }

        static string GetHtmlControl(string templateFolder)
        {
            string returnValue = "";
            string[] templateFiles = Directory.GetFiles(templateFolder, "templateControl.html");

            if (templateFiles.Length > 0)
            {
                returnValue = File.ReadAllText(templateFiles[0]);
            }

            return returnValue;
        }

        static string GenerateHtml(string fileToSave, string htmlTemplate, string javascript, string htmlBody)
        {
            string returnValue = "";
            string finalContent = htmlTemplate;

            finalContent = finalContent.Replace("_SCRIPT_", javascript);
            finalContent = finalContent.Replace("_BODY_", htmlBody);

            try
            {
                File.WriteAllText(fileToSave, finalContent);
                returnValue = fileToSave;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return returnValue;
        }

        static string ResolveStyle(ExcelRange currentCell)
        {
            string returnValue = "";

            if (!string.IsNullOrEmpty(currentCell.Style.Fill.BackgroundColor.Rgb))
            {
                returnValue += string.Format("background-color:#{0};", currentCell.Style.Fill.BackgroundColor.Rgb.Substring(2));
            }
            if (!string.IsNullOrEmpty(currentCell.Style.Font.Color.Rgb))
            {
                returnValue += string.Format("color:#{0};", currentCell.Style.Font.Color.Rgb.Substring(2));
            }
            if (!string.IsNullOrEmpty(currentCell.Style.Font.Size + ""))
            {
                returnValue += string.Format("font-size:{0};", currentCell.Style.Font.Size + "px");
            }
            if (currentCell.Style.Font.Bold)
            {
                returnValue += string.Format("font-weight:bold;");
            }
            if (currentCell.Style.Font.Italic)
            {
                returnValue += string.Format("font-style:italic;");
            }
            if (currentCell.Style.Font.UnderLine)
            {
                returnValue += string.Format("text-decoration:underline;");
            }
            if (!string.IsNullOrEmpty(currentCell.Style.HorizontalAlignment + ""))
            {
                if ((currentCell.Style.HorizontalAlignment + "").Equals("center", StringComparison.InvariantCultureIgnoreCase))
                {
                    returnValue += string.Format("text-align:{0};", currentCell.Style.HorizontalAlignment + "");
                }
            }
            if (!string.IsNullOrEmpty(currentCell.Style.VerticalAlignment + ""))
            {
                returnValue += string.Format("vertical-align:{0};", currentCell.Style.VerticalAlignment + "");
            }

            return returnValue;
        }

        static CellAddress ResolveAddress(string fullAddress)
        {
            CellAddress returnValue = new CellAddress();

            if (!string.IsNullOrEmpty(fullAddress))
            {
                fullAddress = fullAddress.Substring(fullAddress.IndexOf("!") + 1);
                string fromAddress = fullAddress.Substring(0, fullAddress.IndexOf(":"));
                string toAddress = fullAddress.Substring(fullAddress.IndexOf(":") + 1);

                string[] fromAddressComponents = fromAddress.Split("$");
                returnValue.FromCol = Encoding.ASCII.GetBytes(fromAddressComponents[1])[0] - 64;
                returnValue.FromRow = int.Parse(fromAddressComponents[2]);

                string[] toAddressComponents = toAddress.Split("$");
                returnValue.ToCol = Encoding.ASCII.GetBytes(toAddressComponents[1])[0] - 64;
                returnValue.ToRow = int.Parse(toAddressComponents[2]);
            }

            return returnValue;
        }
    }

    public class CellAddress
    {
        public int FromRow { get; set; }
        public int FromCol { get; set; }
        public int ToRow { get; set; }
        public int ToCol { get; set; }
    }

    public class CellStyle
    {
        public string BackgroundColor { get; set; }
        public string FontColor { get; set; }
        public string FontSize { get; set; }
        public string HorizontalAlign { get; set; }
        public string VerticalAlign { get; set; }
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public bool Underline { get; set; }
    }

    public class Template
    {
        public string MarkerType { get; set; }
        public string Content { get; set; }
    }

    public class Marker
    {
        public string Id { get; set; }
        public string Type { get; set; }
        public string Value { get; set; }
    }
}
