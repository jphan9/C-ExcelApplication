using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace CSharpExcelAPplication
{
    class ExcelApplication 
    {
        public void writeToExcelLow(string fileName, string openFileName)
        {
            MainWindow main = new MainWindow();

            Excel.Application oXL;
            Excel._Workbook oWB;
            Excel._Worksheet oSheet;
            Excel.Range oRng;

            List<string> listOfDataA;
            List<string> listOfDataB;
            List<string> listOfDataC;

            int cIndex1 = 1, cIndex2;
            int index = 1;

            // Start Excel and get Application object.
            oXL = new Excel.Application();
            oXL.Visible = true;

            // Get a new workbook.
            //oWB = (Excel._Workbook)(oXL.Workbooks.Add(System.Reflection.Missing.Value));
            oWB = (Excel._Workbook)(oXL.Workbooks.Open(@openFileName));
            //oSheet = (Excel._Worksheet)oWB.ActiveSheet;
            oSheet = (Excel._Worksheet)oWB.Sheets[2];

            // Assign new name to worksheet.
            oSheet.Name = "Data Low";

            // Delete everything in that sheet. 
            oSheet.Cells.ClearContents();

            listOfDataA = main.readInFile(fileName, 0);
            listOfDataB = main.readInFile(fileName, 1);
            listOfDataC = main.readInFile(fileName, 2);

            int rowIndexA = 1;
            int rowIndexB = 1;
            int rowIndexC = 1;

            for (int i = 0; i < listOfDataA.Count; i++)
            {
                oSheet.Cells[rowIndexA, "A"] = listOfDataA[i];
                rowIndexA++;
            }

            for (int i = 0; i < listOfDataB.Count; i++)
            {
                oSheet.Cells[rowIndexB, "B"] = listOfDataB[i];
                rowIndexB++;
            }

            for (int i = 0; i < listOfDataC.Count; i++)
            {
                oSheet.Cells[rowIndexC, "C"] = listOfDataC[i];
                rowIndexC++;
            }

            //gets the total number of cells used in Column C
            int columnTotalC = oSheet.UsedRange.Columns["C:C", Type.Missing].rows.count;

            for (int i = 0; i < columnTotalC / 4; i++)
            {
                cIndex2 = cIndex1 + 3;
                oSheet.Range["G" + Convert.ToString(index)].Formula = "=Average(C" + Convert.ToString(cIndex1) + ":C" + Convert.ToString(cIndex2);
                cIndex1 = cIndex2 + 1;
                index++;
            }
        }

        public void writeToExcelMed(string fileName)
        {
            Excel.Application oXL;
            Excel._Workbook oWB;
            Excel._Worksheet oSheet;
            Excel.Range oRng;

            MainWindow main = new MainWindow();

            List<string> listOfDataA;
            List<string> listOfDataB;
            List<string> listOfDataC;

            int cIndex1 = 1, cIndex2;
            int index = 1;

            // Start Excel and get Application object.
            oXL = new Excel.Application();
            oXL.Visible = true;

            // Get a new workbook.
            oWB = (Excel._Workbook)(oXL.Workbooks.Add(System.Reflection.Missing.Value));
            //var xlSheets = oWB.Sheets as Excel.Sheets;
            //oSheet = (Excel._Worksheet)xlSheets.Add(xlSheets[1], Type.Missing, Type.Missing, Type.Missing);
            oSheet = (Excel._Worksheet)oWB.Sheets[2];
            // Assign new name to worksheet.
            oSheet.Name = "Data Med";

            listOfDataA = main.readInFile(fileName, 0);
            listOfDataB = main.readInFile(fileName, 1);
            listOfDataC = main.readInFile(fileName, 2);

            int rowIndexA = 1;
            int rowIndexB = 1;
            int rowIndexC = 1;

            for (int i = 0; i < listOfDataA.Count; i++)
            {
                oSheet.Cells[rowIndexA, "A"] = listOfDataA[i];
                rowIndexA++;
            }

            for (int i = 0; i < listOfDataB.Count; i++)
            {
                oSheet.Cells[rowIndexB, "B"] = listOfDataB[i];
                rowIndexB++;
            }

            for (int i = 0; i < listOfDataC.Count; i++)
            {
                oSheet.Cells[rowIndexC, "C"] = listOfDataC[i];
                rowIndexC++;
            }

            //gets the total number of cells used in Column C
            int columnTotalC = oSheet.UsedRange.Columns["C:C", Type.Missing].rows.count;

            for (int i = 0; i < columnTotalC / 4; i++)
            {
                cIndex2 = cIndex1 + 3;
                oSheet.Range["G" + Convert.ToString(index)].Formula = "=Average(C" + Convert.ToString(cIndex1) + ":C" + Convert.ToString(cIndex2);
                cIndex1 = cIndex2 + 1;
                index++;
            }
        }
    }
}
