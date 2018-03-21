using System;
using Excel = Microsoft.Office.Interop.Excel; //if this is not working, you need to add a reference to the excel interop
using System.Reflection;

namespace Blueprint
{
    class WelderID
    {
        //excel variables
        //private Excel.Application oXL;
        private Excel._Workbook oWB;
        private Excel._Worksheet oSheet;
        private Excel.Range oRng;
        private int firstRow, lastRow, currentRow;

        //document variables
        private int sheet_count;
        private string drawing_number;
        private string revision_number;
        private string serial_number;


        //constructor - sets initial data
        public WelderID(Project project, Excel._Workbook currentWorkbook)
        {
            oWB = currentWorkbook;
            firstRow = 1;
            lastRow = 29;
            currentRow = 1;
            sheet_count = 1;

            drawing_number = project.drawingNumber;
            revision_number = project.revisionNumber;
            serial_number = project.serialNumber;
        }


        public void generate()
        {
            //start Excel and get Application object.
            //oXL = new Excel.Application();
            //oXL.Visible = true;

            //get a new workbook.
            //oWB = oXL.Workbooks.Add();

            //add stuff to sheet
            sheet1_setup();
            title_block();
            hop_welds();
        }


        private void sheet1_setup()
        {
            oSheet = oWB.Sheets.Add();
            //oSheet = oWB.ActiveSheet; //grab the next worksheet
            oSheet.Name = "WeldID"; //excel sheet name

            //set up excel view window so that it looks cool while populating
            //oXL.ActiveWindow.View = Excel.XlWindowView.xlPageBreakPreview;
            //oXL.ActiveWindow.Zoom = 80;

            //margins n shit - leave these alone
            oSheet.PageSetup.CenterHorizontally = true;
            oSheet.PageSetup.CenterVertically = true;
            oSheet.PageSetup.TopMargin = .25;
            oSheet.PageSetup.BottomMargin = .25;
            oSheet.PageSetup.LeftMargin = .25;
            oSheet.PageSetup.RightMargin = .25;

            //format every cell
            oRng = oSheet.Range[oSheet.Cells[firstRow, 1], oSheet.Cells[lastRow, 12]];
            oRng.RowHeight = 20;
            oRng.ColumnWidth = 8;
            oRng.Font.Size = 11;
            oRng.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            for (int i = 5; i <= lastRow; i++)
            {
                oRng = oSheet.Range[oSheet.Cells[i, 2], oSheet.Cells[i, 4]];
                oRng.Merge();
                oRng = oSheet.Range[oSheet.Cells[i, 5], oSheet.Cells[i, 6]];
                oRng.Merge();
                oRng = oSheet.Range[oSheet.Cells[i, 7], oSheet.Cells[i, 8]];
                oRng.Merge();
                oRng = oSheet.Range[oSheet.Cells[i, 9], oSheet.Cells[i, 10]];
                oRng.Merge();
                oRng = oSheet.Range[oSheet.Cells[i, 11], oSheet.Cells[i, 12]];
                oRng.Merge();
            }

            oRng = oSheet.Range[oSheet.Cells[3, 1], oSheet.Cells[lastRow, 12]];
            oRng.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlMedium;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlMedium;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlMedium;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlMedium;
        }


        private void title_block()
        {
            //LINE 1
            oRng = oSheet.Range[oSheet.Cells[currentRow, 1], oSheet.Cells[currentRow, 12]];
            oRng.Merge();
            oRng.Font.Bold = true;
            oRng.Font.Size = 14;
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            oRng.Value = "REFRIGERATION VALVES AND SYSTEMS VESSEL TRAVELER";
            currentRow++;

            //LINE 2
            oRng = oSheet.Range[oSheet.Cells[currentRow, 1], oSheet.Cells[currentRow, 12]];
            oRng.Merge();
            oRng.Font.Bold = true;
            oRng.Font.Size = 14;
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            oRng.Value = "WELDER DOCUMENTATION / LEAK TEST";
            currentRow++;

            //LINE 3
            oRng = oSheet.Range[oSheet.Cells[currentRow, 1], oSheet.Cells[currentRow, 4]];
            oRng.Merge();
            oRng.Font.Bold = true;
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            oRng.Value = "DRAWING NO. " + drawing_number;

            oRng = oSheet.Range[oSheet.Cells[currentRow, 5], oSheet.Cells[currentRow, 6]];
            oRng.Merge();
            oRng.Font.Bold = true;
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            oRng.Value = "REV NO. " + revision_number;

            oRng = oSheet.Range[oSheet.Cells[currentRow, 7], oSheet.Cells[currentRow, 9]];
            oRng.Merge();
            oRng.Font.Bold = true;
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            oRng.Value = "SERIAL NO. " + serial_number;

            oRng = oSheet.Range[oSheet.Cells[currentRow, 10], oSheet.Cells[currentRow, 12]];
            oRng.Merge();
            oRng.Font.Bold = true;
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            oRng.Value = "SHEET " + 1 + " OF " + sheet_count;
            currentRow++;

            //LINE 4
            oRng = oSheet.Range[oSheet.Cells[currentRow, 1], oSheet.Cells[currentRow, 4]];
            oRng.RowHeight = 42;
            oRng.Merge();
            oRng.WrapText = true;
            oRng.Font.Bold = true;
            oRng.Value = "Description of Weld";


            oRng = oSheet.Range[oSheet.Cells[currentRow, 5], oSheet.Cells[currentRow, 6]];
            oRng.Merge();
            oRng.WrapText = true;
            oRng.Font.Bold = true;
            oRng.Value = "Nozzle WD (as applicable)";

            oRng = oSheet.Range[oSheet.Cells[currentRow, 7], oSheet.Cells[currentRow, 8]];
            oRng.Merge();
            oRng.WrapText = true;
            oRng.Font.Bold = true;
            oRng.Value = "Welder(s) Stamp #";

            oRng = oSheet.Range[oSheet.Cells[currentRow, 9], oSheet.Cells[currentRow, 10]];
            oRng.Merge();
            oRng.WrapText = true;
            oRng.Font.Bold = true;
            oRng.Value = "Accept (Yes/No)";

            oRng = oSheet.Range[oSheet.Cells[currentRow, 11], oSheet.Cells[currentRow, 12]];
            oRng.Merge();
            oRng.WrapText = true;
            oRng.Font.Bold = true;
            oRng.Value = "Pressure Test Inspector (initial/date)";
            currentRow++;
        }

        private void fill_in_empty_spaces()
        {
            int bla = currentRow;
            for (int row_number = bla; row_number <= lastRow; row_number++)
            {
                oRng = oSheet.Cells[row_number, 1];
                oRng.Value = " ";
            }
        }


        private void add_item(string weldType, string weldDescription)
        {
            int rowOffset = currentRow * 30 - 102;

            oRng = oSheet.Cells[currentRow, 1];
            oRng.Value = weldType;
            oRng = oSheet.Cells[currentRow, 2];
            oRng.Value = weldDescription;

            currentRow++;
        }


        private void hop_welds()
        {
            add_item("C", "REF GIRTH");
            add_item("C", "NON REF GIRTH");
            currentRow++;
            add_item("N", "NOZ. A");
            add_item("N", "NOZ. B");
            add_item("N", "NOZ. C");
            add_item("N", "NOZ. D");
            add_item("N", "NOZ. E");
        }
    }
}
