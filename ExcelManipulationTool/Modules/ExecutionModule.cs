using ClosedXML.Excel;
using ExcelManipulationTool.Classes;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelManipulationTool.Modules
{
    public class ExecutionModule
    {
        private List<ExampleClass> valuesFromSheet = new List<ExampleClass>();

        private Boolean testMode = false;

        public ExecutionModule(String fileToRead, Boolean testMode)
        {
            this.testMode = testMode;

            readValuesFromFile(fileToRead);
            createNewWorkbook();
        }

        public void createNewWorkbook()
        {
            XLWorkbook xLWorkbook = new XLWorkbook();
            IXLWorksheet excelWorksheet = xLWorkbook.AddWorksheet("Example Worksheet");

            List<String> headers = new List<String>();

            headers.Add("Column 1");
            headers.Add("Column 2");
            headers.Add("Column 3");

            int currentRow = 1;
            int currentColumn = 1;

            excelWorksheet.Cell(currentRow, currentColumn).Style.Border.LeftBorder = XLBorderStyleValues.Thick;

            foreach (String header in headers)
            {
                excelWorksheet.Cell(currentRow, currentColumn).Value = "'" + header;
                excelWorksheet.Cell(currentRow, currentColumn).Style.Fill.BackgroundColor = XLColor.FromArgb(47, 117, 181);
                excelWorksheet.Cell(currentRow, currentColumn).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                excelWorksheet.Cell(currentRow, currentColumn).Style.Font.FontColor = XLColor.White;
                excelWorksheet.Cell(currentRow, currentColumn).Style.Font.Bold = true;
                excelWorksheet.Cell(currentRow, currentColumn).Style.Border.TopBorder = XLBorderStyleValues.Thick;
                excelWorksheet.Cell(currentRow, currentColumn).Style.Border.BottomBorder = XLBorderStyleValues.Thick;
                excelWorksheet.Cell(currentRow, currentColumn).Style.Border.RightBorder = XLBorderStyleValues.Thin;

                currentColumn++;
            }

            excelWorksheet.Cell(currentRow, currentColumn).Style.Border.LeftBorder = XLBorderStyleValues.None;
            excelWorksheet.Cell(currentRow, currentColumn).Style.Border.LeftBorder = XLBorderStyleValues.Thick;

            foreach (ExampleClass exampleClass in valuesFromSheet)
            {
                currentColumn = 1;
                currentRow++;

                //Column 1
                excelWorksheet.Cell(currentRow, currentColumn).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                excelWorksheet.Cell(currentRow, currentColumn).Value = exampleClass.Column_1;
                excelWorksheet.Cell(currentRow, currentColumn).Style.Fill.BackgroundColor = XLColor.NoColor;
                excelWorksheet.Cell(currentRow, currentColumn).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                excelWorksheet.Cell(currentRow, currentColumn).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                currentColumn++;

                //Column 2
                excelWorksheet.Cell(currentRow, currentColumn).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                excelWorksheet.Cell(currentRow, currentColumn).Value = exampleClass.Column_2;
                excelWorksheet.Cell(currentRow, currentColumn).Style.Fill.BackgroundColor = XLColor.NoColor;
                excelWorksheet.Cell(currentRow, currentColumn).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                excelWorksheet.Cell(currentRow, currentColumn).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                currentColumn++;

                //Column 3
                excelWorksheet.Cell(currentRow, currentColumn).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                excelWorksheet.Cell(currentRow, currentColumn).Value = exampleClass.Column_3;
                excelWorksheet.Cell(currentRow, currentColumn).Style.Fill.BackgroundColor = XLColor.NoColor;
                excelWorksheet.Cell(currentRow, currentColumn).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                excelWorksheet.Cell(currentRow, currentColumn).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                currentColumn++;
            }

            ///////////////
            /// SAVE //////
            ///////////////

            String finalFilePath = "";

            if (testMode != true)
            {
                SaveFileDialog saveFileDialog1 = new SaveFileDialog();

                saveFileDialog1.Filter = "Excel Files (*.xlsx)|"; // or just "txt files (*.txt)|*.txt" if you only want to save text files

                saveFileDialog1.RestoreDirectory = true;

                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    finalFilePath = saveFileDialog1.FileName + ".xlsx";
                    xLWorkbook.SaveAs(finalFilePath);
                    System.Diagnostics.Process.Start(finalFilePath);
                }
            }
            else
            {
                finalFilePath = Path.GetTempPath() + Path.GetFileNameWithoutExtension(Path.GetRandomFileName()) + ".xlsx";
                xLWorkbook.SaveAs(finalFilePath);
                System.Diagnostics.Process.Start(finalFilePath);
            }
        }

        public void readValuesFromFile(String fileToRead)
        {
            XLWorkbook excelWorkbook = new XLWorkbook(fileToRead);
            IXLWorksheet excelWorksheet = excelWorkbook.Worksheet(1);
            IXLRange xlRange = excelWorksheet.RangeUsed();

            //In case some annoying cells are merged together, makes parsing easier
            xlRange.Unmerge();

            int rowCount = xlRange.RowCount();
            int colCount = xlRange.ColumnCount();

            IXLCell rangeToFindCell = xlRange.Search("RangeToFind").SingleOrDefault();
            int row = rangeToFindCell.Address.RowNumber;
            int column = rangeToFindCell.Address.ColumnNumber;

            Coordinate tableTopLeftCoordinate = new Coordinate(row + 1, column);
            Coordinate tableTopRightCoordinate = new Coordinate(row + 1, colCount);
            Coordinate tableBottomLeftCoordinate = new Coordinate(rowCount, column);
            Coordinate tableBottomRightCoordinate = new Coordinate(rowCount, colCount);

            int x_value = tableTopLeftCoordinate.X;

            while (x_value <= tableBottomLeftCoordinate.X)
            {
                ExampleClass recordToAdd = new ExampleClass();

                for (int y_value = tableTopLeftCoordinate.Y; y_value <= tableTopRightCoordinate.Y; y_value++)
                {
                    String valueOfCell = Convert.ToString(xlRange.Cell(x_value, y_value).Value);

                    if (y_value == 1)
                    {
                        recordToAdd.Column_1 = valueOfCell;
                    }

                    if (y_value == 2)
                    {
                        recordToAdd.Column_2 = valueOfCell;
                    }

                    if (y_value == 3)
                    {
                        recordToAdd.Column_3 = valueOfCell;
                    }
                }

                valuesFromSheet.Add(recordToAdd);

                x_value++;
            }

            excelWorkbook.Dispose();
        }
    }
}
