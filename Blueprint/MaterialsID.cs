using System;
using Excel = Microsoft.Office.Interop.Excel; //if this is not working, you need to add a reference to the excel interop
using System.Reflection;


namespace Blueprint
{
    class MaterialsID
    {
        //excel variables
        //private Excel.Application oXL;
        private Excel._Workbook oWB;
        private Excel._Worksheet oSheet;
        private Excel.Range oRng;
        private Excel.OLEObjects objs;
        private Excel.OLEObject obj;
        private int firstRow, lastRow, currentRow;

        //document variables
        private int sheet_count;
        private string drawing_number;
        private string revision_number;
        private string serial_number;


        public MaterialsID(Project project, Excel._Workbook currentWorkbook)
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
            ////start Excel and get Application object.
            //oXL = new Excel.Application();
            //oXL.Visible = true;

            ////get a new workbook.
            //oWB = oXL.Workbooks.Add();

            //---SHEET1
            sheet1_setup();
            title_block1();
            hop_materials();
        }


        private void sheet1_setup()
        {
            //oSheet = oWB.ActiveSheet; //grab the first worksheet
            oSheet = oWB.Sheets.Add();
            oSheet.Name = "MaterialsID"; //excel sheet name
            objs = oSheet.OLEObjects(); //grab objects (like checkmarks) on that page

            ////set up excel view window so that it looks cool while populating
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

            for (int i = 6; i <= lastRow; i++)
            {
                oRng = oSheet.Range[oSheet.Cells[i, 1], oSheet.Cells[i, 8]];
                oRng.Merge();
                oRng = oSheet.Range[oSheet.Cells[i, 9], oSheet.Cells[i, 10]];
                oRng.Merge();
                oRng = oSheet.Range[oSheet.Cells[i, 11], oSheet.Cells[i, 12]];
                oRng.Merge();
            }

            oRng = oSheet.Range[oSheet.Cells[6, 1], oSheet.Cells[lastRow, 12]];
            oRng.RowHeight = 30;
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


            //oRng = oSheet1.Range[oSheet1.Cells[4, 12], oSheet1.Cells[lastRow, 13]];
            //oRng.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
            //oRng.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
            //oRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
        }


        private void title_block1()
        {
            //LINE 1
            oRng = oSheet.Range[oSheet.Cells[currentRow, 1], oSheet.Cells[currentRow, 8]];
            oRng.Merge();
            oRng.Font.Bold = true;
            oRng.Font.Size = 14;
            oRng.Value = "REFRIGERATION VALVES AND SYSTEMS MATERIALS DOCUMENTATION";

            oRng = oSheet.Cells[currentRow, 9];
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            oRng.Value = "SHT";

            oRng = oSheet.Cells[currentRow, 10];
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Value = 1;

            oRng = oSheet.Cells[currentRow, 11];
            oRng.Value = "OF";

            oRng = oSheet.Cells[currentRow, 12];
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Value = sheet_count;
            currentRow++;

            //LINE 2
            oRng = oSheet.Range[oSheet.Cells[currentRow, 1], oSheet.Cells[currentRow, 2]];
            oRng.Merge();
            oRng.Font.Bold = true;
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            oRng.Value = "DRAWING #";

            oRng = oSheet.Cells[currentRow, 3];
            oRng.Font.Bold = true;
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Value = drawing_number;

            oRng = oSheet.Range[oSheet.Cells[currentRow, 5], oSheet.Cells[currentRow, 6]];
            oRng.Merge();
            oRng.Font.Bold = true;
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            oRng.Value = "REV.";

            oRng = oSheet.Cells[currentRow, 7];
            oRng.Font.Bold = true;
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Value = revision_number;

            oRng = oSheet.Range[oSheet.Cells[currentRow, 9], oSheet.Cells[currentRow, 10]];
            oRng.Merge();
            oRng.Font.Bold = true;
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            oRng.Value = "SERIAL #";

            oRng = oSheet.Cells[currentRow, 11];
            oRng.Font.Bold = true;
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Value = serial_number;
            currentRow++;

            //LINE 3
            oRng = oSheet.Cells[currentRow, 1];
            oRng.RowHeight = 4;
            currentRow++;

            //LINE 4
            oRng = oSheet.Range[oSheet.Cells[currentRow, 1], oSheet.Cells[currentRow, 12]];
            oRng.Merge();
            oRng.RowHeight = 30;
            oRng.WrapText = true;
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            oRng.Value = "VERIFY MATERIAL SIZE, SCH, THICKNESS, RATING OF PRESSURE BOUNDARY ITEMS & ITEMS WELDED TO THE PRESSURE BOUNDARY";
            oRng.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlMedium;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlMedium;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlMedium;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlMedium;
            currentRow++;

            //LINE 5
            oRng = oSheet.Cells[currentRow, 1];
            oRng.RowHeight = 4;
            currentRow++;

            //LINE 6
            oRng = oSheet.Range[oSheet.Cells[currentRow, 1], oSheet.Cells[currentRow, 8]];
            oRng.RowHeight = 30;
            oRng.Font.Bold = true;
            oRng.Value = "ITEM";
            

            oRng = oSheet.Range[oSheet.Cells[currentRow, 9], oSheet.Cells[currentRow, 10]];
            oRng.WrapText = true;
            oRng.Font.Bold = true;
            oRng.Value = "IDENTIFICATION REQ'D";

            oRng = oSheet.Range[oSheet.Cells[currentRow, 11], oSheet.Cells[currentRow, 12]];
            oRng.WrapText = true;
            oRng.Font.Bold = true;
            oRng.Value = "IDENTIFICATION ON ITEM";
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


        private void add_item(string itemDescription, string idType, bool hasCheckboxes)
        {
            int rowOffset = currentRow * 30 -102;

            oRng = oSheet.Cells[currentRow, 1];
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            oRng.Value = " " + itemDescription;
            oRng = oSheet.Cells[currentRow, 9];
            oRng.Value = idType;

            if (hasCheckboxes)
            {
                //add check boxes
                obj = objs.Add("Forms.Checkbox.1", Missing.Value, Missing.Value, false, false, Missing.Value, Missing.Value, 240, rowOffset, 120, 16);
                obj.Object.Caption = "SPECIAL M'TL REQUIREMENTS";
                obj.Object.Value = false;
                obj.Object.Font.Size = 6.5;

                obj = objs.Add("Forms.Checkbox.1", Missing.Value, Missing.Value, false, false, Missing.Value, Missing.Value, 240, rowOffset + 12, 120, 16);
                obj.Object.Caption = "NORMALIZED M'TL";
                obj.Object.Value = false;
                obj.Object.Font.Size = 6.5;
            }

            currentRow++;
        }


        private void hop_materials()
        {
            add_item("DATA PLATE BRACKET", "COLOR", false);
            add_item("BACKING RINGS", "COLOR", false);
            add_item("SHELL PIPE", "I.D.", true);
            add_item("REFERENCE END HEAD", "HT. NO.", true);
            add_item("NON-REFERENCE END HEAD", "HT. NO.", true);
            add_item("SADDLES", "COLOR", true);
            add_item("NOZZ. A PIPE", "I.D.", true);
            add_item("NOZZ. B PIPE", "I.D.", true);
            add_item("NOZZ. C PIPE", "I.D.", true);
            add_item("NOZZ. D PIPE", "I.D.", true);
            add_item("NOZZ. E PIPE", "I.D.", true);
        }
    }
}

