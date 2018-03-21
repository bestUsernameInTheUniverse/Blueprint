using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;


namespace Blueprint
{
    class Traveler
    {
        //excel variables
        //private Excel.Application oXL;
        private Excel._Workbook oWB;
        private Excel._Worksheet oSheet;
        private Excel.Range oRng;
        private Excel.OLEObjects objs;
        private Excel.OLEObject obj;
        private int firstRow, lastRow, currentRow;

        //traveler variables
        private int sheet_count;
        private string drawing_number;
        private string revision_number;
        private string serial_number;
        private string special_note;


        //constructor
        public Traveler(Project project, Excel._Workbook currentWorkbook)
        {
            oWB = currentWorkbook;
            firstRow = 1;
            lastRow = 39;
            currentRow = 1;
            sheet_count = 1;

            drawing_number = project.drawingNumber;
            revision_number = project.revisionNumber;
            serial_number = project.serialNumber;
            special_note = "MAKE SURE NOT TO MESS UP THE THING WITH THE STUFF";
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

            line_hold_points();
            blank_lines(1);

            line_pipe_plasma("22\"", "8\"", "80");
            blank_lines(1);

            line_layout_head1();
            line_layout_head2();
            blank_lines(1);

            line_internal_exam_hd_shl();
            line_internal_exam_dip_tube();
            line_internal_exam_cleanliness();
            blank_lines(1);

            line_final_vt();
            line_dimensional();
            blank_lines(1);

            line_soap_bubble_preliminary();
            blank_lines(1);

            line_die_stamp();
            line_pressure_test("520", "WATER");
            line_evacuation_test();
            last_signatures();


            //---SHEET2
            //oSheet2 = oWB.Worksheets.Add(Missing.Value, oWB.Sheets[oWB.Sheets.Count]);
        }


        private void blank_lines(int number)
        {
            currentRow += number;
        }


        private void sheet1_setup()
        {
            oSheet = oWB.Sheets.Add();
            ////grab the first worksheet
            //oSheet = oWB.ActiveSheet;
            objs = oSheet.OLEObjects();

            //format page
            oSheet.Name = "Traveler";
            //oXL.ActiveWindow.View = Excel.XlWindowView.xlPageBreakPreview;
            //oXL.ActiveWindow.Zoom = 80;
            oSheet.PageSetup.CenterHorizontally = true;
            oSheet.PageSetup.CenterVertically = true;
            oSheet.PageSetup.TopMargin = .25;
            oSheet.PageSetup.BottomMargin = .25;
            oSheet.PageSetup.LeftMargin = .25;
            oSheet.PageSetup.RightMargin = .25;

            oRng = oSheet.Range[oSheet.Cells[firstRow, 1], oSheet.Cells[lastRow, 13]];
            oRng.ColumnWidth = 8;
            oRng.RowHeight = 20;
            oRng.Font.Size = 11;
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            oRng.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

            oRng = oSheet.Cells[1, 9];
            oRng.ColumnWidth = 5;
            oRng = oSheet.Cells[1, 11];
            oRng.ColumnWidth = 2;
            oRng = oSheet.Cells[1, 12];
            oRng.ColumnWidth = 5;

            for (int row_number = 4; row_number <= 39; row_number++)
            {
                oRng = oSheet.Range[oSheet.Cells[row_number, 1], oSheet.Cells[row_number, 8]];
                oRng.Merge();
                oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                oRng.Value = " ";
            }

            oRng = oSheet.Range[oSheet.Cells[4, 9], oSheet.Cells[lastRow, 9]];
            oRng.Font.Size = 22;
            oRng = oSheet.Range[oSheet.Cells[4, 12], oSheet.Cells[lastRow, 12]];
            oRng.Font.Size = 22;

            oRng = oSheet.Range[oSheet.Cells[1, 1], oSheet.Cells[3, 13]];
            oRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlMedium;

            oRng = oSheet.Range[oSheet.Cells[1, 1], oSheet.Cells[lastRow, 13]];
            oRng.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlMedium;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlMedium;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlMedium;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlMedium;

            oRng = oSheet.Range[oSheet.Cells[4, 1], oSheet.Cells[lastRow, 10]];
            oRng.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;

            oRng = oSheet.Range[oSheet.Cells[4, 12], oSheet.Cells[lastRow, 13]];
            oRng.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
        }


        private void title_block1()
        {
            //LINE 1
            oRng = oSheet.Range[oSheet.Cells[currentRow, 1], oSheet.Cells[currentRow, 8]];
            oRng.Merge();
            oRng.Font.Bold = true;
            oRng.Font.Size = 14;
            oRng.Value = "REFRIGERATION VALVES AND SYSTEMS VESSEL TRAVELER";

            oRng = oSheet.Cells[currentRow, 9];
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            oRng.Value = "SHT";

            oRng = oSheet.Cells[currentRow, 10];
            oRng.Value = 1;

            oRng = oSheet.Range[oSheet.Cells[currentRow, 11], oSheet.Cells[currentRow, 12]];
            oRng.Merge();
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            oRng.Value = "OF";

            oRng = oSheet.Cells[currentRow, 13];
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            oRng.Value = sheet_count;

            currentRow++;

            //LINE 2
            oRng = oSheet.Range[oSheet.Cells[currentRow, 1], oSheet.Cells[currentRow, 2]];
            oRng.Merge();
            oRng.Font.Bold = true;
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            oRng.Value = "DRAWING #";

            oRng = oSheet.Range[oSheet.Cells[currentRow, 3], oSheet.Cells[currentRow, 4]];
            oRng.Merge();
            oRng.Font.Bold = true;
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            oRng.Value = drawing_number;

            oRng = oSheet.Range[oSheet.Cells[currentRow, 5], oSheet.Cells[currentRow, 6]];
            oRng.Merge();
            oRng.Font.Bold = true;
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            oRng.Value = "REV.";

            oRng = oSheet.Range[oSheet.Cells[currentRow, 7], oSheet.Cells[currentRow, 8]];
            oRng.Merge();
            oRng.Font.Bold = true;
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            oRng.Value = revision_number;

            oRng = oSheet.Range[oSheet.Cells[currentRow, 9], oSheet.Cells[currentRow, 10]];
            oRng.Merge();
            oRng.Font.Bold = true;
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            oRng.Value = "SERIAL #";

            oRng = oSheet.Range[oSheet.Cells[currentRow, 11], oSheet.Cells[currentRow, 13]];
            oRng.Merge();
            oRng.Font.Bold = true;
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            oRng.Value = serial_number;

            currentRow++;

            //LINE 3
            oRng = oSheet.Range[oSheet.Cells[currentRow, 1], oSheet.Cells[currentRow, 6]];
            oRng.Merge();

            obj = objs.Add("Forms.Checkbox.1", Missing.Value, Missing.Value, false, false, Missing.Value, Missing.Value, 5, 41, 120, 16);
            obj.Object.Caption = "SPECIAL M'TL REQUIREMENTS";
            obj.Object.Value = false;
            obj.Object.Font.Size = 8;

            obj = objs.Add("Forms.Checkbox.1", Missing.Value, Missing.Value, false, false, Missing.Value, Missing.Value, 160, 41, 90, 16);
            obj.Object.Caption = "NORMALIZED M'TL";
            obj.Object.Value = false;
            obj.Object.Font.Size = 8;

            obj = objs.Add("Forms.Checkbox.1", Missing.Value, Missing.Value, false, false, Missing.Value, Missing.Value, 262, 41, 15, 16);
            obj.Object.Caption = "";
            obj.Object.Value = true;
            obj.Object.Font.Size = 8;

            oRng = oSheet.Range[oSheet.Cells[currentRow, 7], oSheet.Cells[currentRow, 13]];
            oRng.Merge();
            oRng.Font.Size = 8;
            oRng.WrapText = true;
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            oRng.Value = special_note;

            currentRow++;

            //LINE 4
            oRng = oSheet.Cells[currentRow, 1];
            oRng.Value = "DESCRIBE SPECIFIC ITEMS TO BE INSPECTED";
            oRng.Font.Bold = true;

            oRng = oSheet.Cells[currentRow, 9];
            oRng.Value = "X";

            oRng = oSheet.Cells[currentRow, 10];
            oRng.Font.Size = 10;
            oRng.Value = "QC HOLD";

            oRng = oSheet.Cells[currentRow, 12];
            oRng.Value = "X";

            oRng = oSheet.Cells[currentRow, 13];
            oRng.Font.Size = 10;
            oRng.Value = "A/I HOLD";

            currentRow++;
        }


        private void line_hold_points()
        {
            oSheet.Cells[currentRow, 1].Value = "HOLD POINTS SET BY";
            oSheet.Cells[currentRow, 9].Value = "X";
            currentRow++;

            oSheet.Cells[currentRow, 1].Value = "DRAWINGS REVIEWED/ACCEPTED";
            oSheet.Cells[currentRow, 9].Value = "X";
            currentRow++;

            oSheet.Cells[currentRow, 1].Value = "CALCULATIONS REVIEWED/ON FILE";
            oSheet.Cells[currentRow, 9].Value = "X";
            currentRow++;

            obj = objs.Add("Forms.Checkbox.1", Missing.Value, Missing.Value, false, false, Missing.Value, Missing.Value, 250, 122, 100, 16);
            obj.Object.Caption = "STANDARD CALCS USED";
            obj.Object.Value = true;
            obj.Object.Font.Size = 8;
        }


        private void line_plasma(int itemNumber)
        {
            oSheet.Cells[currentRow, 1].Value = "PLASMA SHL #" + itemNumber + " L/O & CUT:  LENGTH___________  WIDTH__________  THK_________";
            oRng = oSheet.Range[oSheet.Cells[currentRow, 1], oSheet.Cells[currentRow, 8]];
            oRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            oSheet.Cells[currentRow + 1, 1].Value = "OPERATOR ____________________  DATE _______________";

            merge_signatures(2);
            oSheet.Cells[currentRow, 9].Value = "X";

            currentRow = currentRow + 2;
        }

        private void line_pipe_plasma(string length, string nps, string schedule)
        {
            oSheet.Cells[currentRow, 1].Value = "PLASMA PIPE SHL L/O & CUT:  LENGTH " + length + "   NPS " + nps + "   SCH " + schedule;
            oRng = oSheet.Range[oSheet.Cells[currentRow, 1], oSheet.Cells[currentRow, 8]];
            oRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            oSheet.Cells[currentRow + 1, 1].Value = "OPERATOR ____________________  DATE _______________";

            merge_signatures(2);
            oSheet.Cells[currentRow, 9].Value = "X";

            currentRow = currentRow + 2;
        }


        private void line_shell(int itemNumber)
        {
            oSheet.Cells[currentRow, 1].Value = "SHELL # " + itemNumber + " RADIUS__________ DIA__________";
            oSheet.Cells[currentRow, 9].Value = "X";
            currentRow++;
        }


        private void line_layout_head1()
        {
            oSheet.Cells[currentRow, 1].Value = "L/O & FIT-UP PRIOR TO ROOT PASS - REFERENCE END HD";
            oSheet.Cells[currentRow, 9].Value = "X";
            currentRow++;
        }


        private void line_layout_shell_long()
        {
            oSheet.Cells[currentRow, 1].Value = "L/O & FIT-UP PRIOR TO ROOT PASS - LG SEAM #_____";
            oSheet.Cells[currentRow, 9].Value = "X";
            currentRow++;
        }


        private void line_layout_shell_girth()
        {
            oSheet.Cells[currentRow, 1].Value = "L/O & FIT-UP PRIOR TO ROOT PASS - GIRTH SEAM #_____";
            oSheet.Cells[currentRow, 9].Value = "X";
            currentRow++;
        }


        private void line_layout_head2()
        {
            oSheet.Cells[currentRow, 1].Value = "L/O & FIT-UP PRIOR TO ROOT PASS - NON-REFERENCE END HD";
            oSheet.Cells[currentRow, 9].Value = "X";
            currentRow++;
        }


        private void line_supports()
        {
            oSheet.Cells[currentRow, 1].Value = "SUPPORTS";
            oSheet.Cells[currentRow, 9].Value = "X";
            currentRow++;
        }


        private void line_layouts()
        {
            oSheet.Cells[currentRow, 1].Value = "L/O SHELL";
            oSheet.Cells[currentRow, 9].Value = "X";
            currentRow++;

            oSheet.Cells[currentRow, 1].Value = "L/O REFERENCE END HD";
            oSheet.Cells[currentRow, 9].Value = "X";
            currentRow++;

            oSheet.Cells[currentRow, 1].Value = "L/O NON-REFERENCE END HD";
            oSheet.Cells[currentRow, 9].Value = "X";
            currentRow++;
        }


        private void line_internal_exam_hd_shl()
        {
            oSheet.Cells[currentRow, 1].Value = "INTERNAL EXAM-SHELL & HEADS & ASSOCIATED WELDS";
            oSheet.Cells[currentRow, 9].Value = "X";
            currentRow++;
        }


        private void line_internal_exam_dip_tube()
        {
            oSheet.Cells[currentRow, 1].Value = "INTERNAL EXAM-DIP TUBE";
            oSheet.Cells[currentRow, 9].Value = "X";
            currentRow++;
        }


        private void line_internal_exam_cleanliness()
        {
            oSheet.Cells[currentRow, 1].Value = "INTERNAL EXAM-CLEANLINESS PRIOR TO CLOSURE";
            oSheet.Cells[currentRow, 9].Value = "X";
            currentRow++;
        }


        private void line_final_vt()
        {
            oSheet.Cells[currentRow, 1].Value = "FINAL VT - OD";
            oSheet.Cells[currentRow, 9].Value = "X";
            currentRow++;
        }


        private void line_dimensional()
        {
            oSheet.Cells[currentRow, 1].Value = "DIMENSIONAL";
            oSheet.Cells[currentRow, 9].Value = "X";
            currentRow++;
        }


        private void line_bubble_test(double designPressure, double multiplier)
        {
            int x = (int)Math.Ceiling(designPressure * multiplier);
            int remainder = x % 5;
            int testPressure;

            if (remainder > 0) testPressure = x + 5 - remainder;
            else testPressure = x;


            oRng = oSheet.Range[oSheet.Cells[currentRow, 1], oSheet.Cells[currentRow + 2, 8]];
            oRng.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            oRng.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            oRng.UnMerge();

            oRng = oSheet.Range[oSheet.Cells[currentRow, 6], oSheet.Cells[currentRow + 2, 6]];
            oRng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;

            oRng = oSheet.Range[oSheet.Cells[currentRow, 1], oSheet.Cells[currentRow, 3]];
            oRng.Merge();
            oRng.Value = "SOAP BUBBLE TEST COIL";

            oRng = oSheet.Range[oSheet.Cells[currentRow, 4], oSheet.Cells[currentRow, 5]];
            oRng.Merge();
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            oRng.Value = "TEST WITH:";

            oSheet.Cells[currentRow, 6].Value = "AIR";

            oSheet.Cells[currentRow, 7].Value = "GAGE #";

            oRng = oSheet.Range[oSheet.Cells[currentRow + 1, 1], oSheet.Cells[currentRow + 1, 3]];
            oRng.Merge();
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            oRng.Value = "PNEUMATIC TEST PRESSURE:";

            oSheet.Cells[currentRow + 1, 4].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            oSheet.Cells[currentRow + 1, 4].Value = testPressure;

            oRng = oSheet.Range[oSheet.Cells[currentRow + 1, 5], oSheet.Cells[currentRow + 1, 6]];
            oRng.Merge();
            oRng.Value = "PSIG";

            oSheet.Cells[currentRow + 1, 7].Value = "GAGE #";

            oRng = oSheet.Range[oSheet.Cells[currentRow + 2, 1], oSheet.Cells[currentRow + 2, 3]];
            oRng.Merge();
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            oRng.Value = "HOLD TEST PRESSURE:";

            oSheet.Cells[currentRow + 2, 4].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            oSheet.Cells[currentRow + 2, 4].Value = 45;

            oRng = oSheet.Range[oSheet.Cells[currentRow + 2, 5], oSheet.Cells[currentRow + 2, 6]];
            oRng.Merge();
            oRng.Value = "MINUTES";

            merge_signatures(3);
            oSheet.Cells[currentRow, 9].Value = "X";

            currentRow = currentRow + 3;
        }


        private void line_soap_bubble_preliminary()
        {
            oSheet.Cells[currentRow, 1].Value = "SOAP BUBBLE TEST (PRELIMINARY)";
            oSheet.Cells[currentRow, 9].Value = "X";
            currentRow++;
        }


        private void line_die_stamp()
        {
            oSheet.Cells[currentRow, 1].Value = "DIE STAMP \"RVS\" & SERIAL NO. ON SHELL AT N.P.";
            oSheet.Cells[currentRow, 9].Value = "X";
            currentRow++;
        }


        private void line_pressure_test(string testPressure, string testMedium)
        {
            //int x = (int)Math.Ceiling(designPressure * multiplier);
            //int remainder = x % 5;
            //int testPressure;

            //if (remainder > 0) testPressure = x + 5 - remainder;
            //else testPressure = x;


            oRng = oSheet.Range[oSheet.Cells[currentRow, 1], oSheet.Cells[currentRow + 2, 8]];
            oRng.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            oRng.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            oRng.UnMerge();

            oRng = oSheet.Range[oSheet.Cells[currentRow, 6], oSheet.Cells[currentRow + 2, 6]];
            oRng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;

            oRng = oSheet.Range[oSheet.Cells[currentRow, 1], oSheet.Cells[currentRow, 3]];
            oRng.Merge();
            oRng.Value = "VESSEL";

            oRng = oSheet.Range[oSheet.Cells[currentRow, 4], oSheet.Cells[currentRow, 5]];
            oRng.Merge();
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            oRng.Value = "TEST WITH:";

            oSheet.Cells[currentRow, 6].Value = testMedium;

            oSheet.Cells[currentRow, 7].Value = "GAGE #";

            oRng = oSheet.Range[oSheet.Cells[currentRow + 1, 1], oSheet.Cells[currentRow + 1, 3]];
            oRng.Merge();
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            if (testMedium.Equals("AIR")) oRng.Value = "PNEUMATIC TEST PRESSURE:";
            else oRng.Value = "HYDRO TEST PRESSURE:";

            oSheet.Cells[currentRow + 1, 4].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            oSheet.Cells[currentRow + 1, 4].Value = testPressure;

            oRng = oSheet.Range[oSheet.Cells[currentRow + 1, 5], oSheet.Cells[currentRow + 1, 6]];
            oRng.Merge();
            oRng.Value = "PSIG";

            oSheet.Cells[currentRow + 1, 7].Value = "GAGE #";

            oRng = oSheet.Range[oSheet.Cells[currentRow + 2, 1], oSheet.Cells[currentRow + 2, 3]];
            oRng.Merge();
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            oRng.Value = "HOLD TEST PRESSURE:";

            oSheet.Cells[currentRow + 2, 4].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            oSheet.Cells[currentRow + 2, 4].Value = 45;

            oRng = oSheet.Range[oSheet.Cells[currentRow + 2, 5], oSheet.Cells[currentRow + 2, 6]];
            oRng.Merge();
            oRng.Value = "MINUTES";

            merge_signatures(3);
            oSheet.Cells[currentRow, 9].Value = "X";

            currentRow = currentRow + 3;
        }


        private void line_evacuation_test()
        {
            oRng = oSheet.Range[oSheet.Cells[currentRow, 1], oSheet.Cells[currentRow + 3, 8]];
            oRng.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            oRng.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            oRng.UnMerge();

            //oRng = oSheet1.Range[oSheet1.Cells[currentRow, 6], oSheet1.Cells[currentRow + 3, 8]];
            //oRng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;


            oRng = oSheet.Range[oSheet.Cells[currentRow, 1], oSheet.Cells[currentRow, 3]];
            oRng.Merge();
            oRng.Value = "EVACUATION TEST";

            oRng = oSheet.Range[oSheet.Cells[currentRow, 4], oSheet.Cells[currentRow, 5]];
            oRng.Merge();
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            oRng.Value = "TEST DATE:";

            oRng = oSheet.Range[oSheet.Cells[currentRow, 6], oSheet.Cells[currentRow, 8]];
            oRng.Merge();
            oRng.Value = "____________________";

            oRng = oSheet.Range[oSheet.Cells[currentRow + 1, 1], oSheet.Cells[currentRow + 1, 2]];
            oRng.Merge();
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            oRng.Value = "AMBIENT TEMP:";

            oRng = oSheet.Range[oSheet.Cells[currentRow + 1, 3], oSheet.Cells[currentRow + 1, 4]];
            oRng.Merge();
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            oRng.Value = "______________";

            oRng = oSheet.Range[oSheet.Cells[currentRow + 1, 5], oSheet.Cells[currentRow + 1, 8]];
            oRng.Merge();
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            oRng.Value = "°F";

            oRng = oSheet.Range[oSheet.Cells[currentRow + 2, 1], oSheet.Cells[currentRow + 2, 2]];
            oRng.Merge();
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            oRng.Value = "START TIME:";

            oRng = oSheet.Range[oSheet.Cells[currentRow + 2, 3], oSheet.Cells[currentRow + 2, 4]];
            oRng.Merge();
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            oRng.Value = "______________";

            oRng = oSheet.Range[oSheet.Cells[currentRow + 2, 5], oSheet.Cells[currentRow + 2, 6]];
            oRng.Merge();
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            oRng.Value = "INCHES Hg:";

            oRng = oSheet.Range[oSheet.Cells[currentRow + 2, 7], oSheet.Cells[currentRow + 2, 8]];
            oRng.Merge();
            oRng.Value = "______________";

            oRng = oSheet.Range[oSheet.Cells[currentRow + 3, 1], oSheet.Cells[currentRow + 3, 2]];
            oRng.Merge();
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            oRng.Value = "FINISH TIME:";

            oRng = oSheet.Range[oSheet.Cells[currentRow + 3, 3], oSheet.Cells[currentRow + 3, 4]];
            oRng.Merge();
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            oRng.Value = "______________";

            oRng = oSheet.Range[oSheet.Cells[currentRow + 3, 5], oSheet.Cells[currentRow + 3, 6]];
            oRng.Merge();
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            oRng.Value = "INCHES Hg:";

            oRng = oSheet.Range[oSheet.Cells[currentRow + 3, 7], oSheet.Cells[currentRow + 3, 8]];
            oRng.Merge();
            oRng.Value = "______________";

            merge_signatures(4);
            oSheet.Cells[currentRow, 9].Value = "X";

            currentRow = currentRow + 4;
        }


        private void last_signatures()
        {
            oSheet.Cells[currentRow, 1].Value = "DATA PLATE ATTACHED & STAMP ACCEPTED";
            oSheet.Cells[currentRow, 9].Value = "X";
            currentRow++;

            oSheet.Cells[currentRow, 1].Value = "DOCUMENTATION REVIEWED & ACCEPTED";
            oSheet.Cells[currentRow, 9].Value = "X";
            currentRow++;

            oSheet.Cells[currentRow, 1].Value = "DATA REPORT SIGNED";
            oSheet.Cells[currentRow, 9].Value = "X";
            currentRow++;
        }


        private void merge_signatures(int rowCount)
        {
            if (rowCount > 1)
            {
                rowCount = rowCount - 1;

                oRng = oSheet.Range[oSheet.Cells[currentRow, 9], oSheet.Cells[currentRow + rowCount, 9]];
                oRng.Merge();

                oRng = oSheet.Range[oSheet.Cells[currentRow, 10], oSheet.Cells[currentRow + rowCount, 10]];
                oRng.Merge();

                oRng = oSheet.Range[oSheet.Cells[currentRow, 11], oSheet.Cells[currentRow + rowCount, 11]];
                oRng.Merge();

                oRng = oSheet.Range[oSheet.Cells[currentRow, 12], oSheet.Cells[currentRow + rowCount, 12]];
                oRng.Merge();

                oRng = oSheet.Range[oSheet.Cells[currentRow, 13], oSheet.Cells[currentRow + rowCount, 13]];
                oRng.Merge();
            }
        }

    }
}
