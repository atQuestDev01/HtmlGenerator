using OfficeOpenXml;
using System;
using System.IO;
using System.Text;

namespace HtmlGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            Console.WriteLine("Loading Excel file.");

            FileInfo fi = new FileInfo(@"C:\\Users\\chuns\\OneDrive\\Documents\\GitHub\\HtmlGenerator\\Template01.xlsx");
            using (ExcelPackage ep = new ExcelPackage(fi))
            {
                foreach(ExcelWorksheet currentWorksheet in ep.Workbook.Worksheets)
                {
                    Console.WriteLine(currentWorksheet.Name);

                    //Get the name range
                    foreach (ExcelNamedRange currentNameRange in currentWorksheet.Workbook.Names)
                    {
                        CellAddress currentNameRangeAddress = ResolveAddress(currentNameRange.FullAddressAbsolute);

                        object[,] cellInfo = (object[,])currentNameRange.Value;

                        string currentRow = "";
                        for(int row = 0; row < currentNameRange.Rows; row++)
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

                                currentRowColumns += string.Format("<td colspan=\"{0}\" style=\"{1}\">{2}</td>", iColSpan, ResolveStyle(currentCell), colValue);
                                colValue = "";
                                iColSpan = 1;
                            }
                            #endregion

                            currentRow += currentRowColumns + Environment.NewLine;
                            currentRow += "</tr>" + Environment.NewLine;
                        }

                        string currentTable = "";
                        currentTable += "<table>" + Environment.NewLine;
                        currentTable += currentRow + Environment.NewLine;
                        currentTable += "</table>" + Environment.NewLine;

                        Console.WriteLine(currentTable);
                    }
                }
            }
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
}
