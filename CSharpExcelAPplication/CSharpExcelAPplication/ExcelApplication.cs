﻿using System;
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
        // Function that opens all the files.
        public void openExcel(string openFileName, string fileName, string fileName1, string fileName2, string fileName3)
        {
            Excel.Application oXL;
            Excel._Workbook oWB;
            Excel._Worksheet oSheet;

            // Start Excel and get Application object.
            oXL = new Excel.Application();
            oXL.Visible = true;

            // Get a new workbook.
            oWB = (Excel._Workbook)(oXL.Workbooks.Open(@openFileName));


            // Check to see if the user selected a file for the data. If the user does select a file check to write that data to the appropriate tab.     
            if (fileName != "No File Selected")
            {
                try
                {
                    oSheet = (Excel._Worksheet)oWB.Sheets["Data Low"];
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    oSheet = (Excel._Worksheet)oWB.Worksheets.Add(Type.Missing, oWB.Sheets[1], Type.Missing, Type.Missing);
                }
                writeToExcel(fileName, oSheet, "Data Low");
            }

            if (fileName1 != "No File Selected")
            {
                try
                {
                    oSheet = (Excel._Worksheet)oWB.Sheets["Data Med"];
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    oSheet = (Excel._Worksheet)oWB.Worksheets.Add(Type.Missing, oWB.Sheets[2], Type.Missing, Type.Missing);
                }
                writeToExcel(fileName1, oSheet, "Data Med");
            }

            if (fileName2 != "No File Selected")
            {
                try
                {
                    oSheet = (Excel._Worksheet)oWB.Sheets["Data High"];
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    oSheet = (Excel._Worksheet)oWB.Worksheets.Add(Type.Missing, oWB.Sheets[3], Type.Missing, Type.Missing);
                }
                writeToExcel(fileName2, oSheet, "Data High");
            }

            if (fileName3 != "No File Selected")
            {
                try
                {
                    oSheet = (Excel._Worksheet)oWB.Sheets["Data Ultra"];
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    oSheet = (Excel._Worksheet)oWB.Worksheets.Add(Type.Missing, oWB.Sheets[4], Type.Missing, Type.Missing);
                }
                writeToExcel(fileName3, oSheet, "Data Ultra");
            }

        }

        //Function that writes all the data from the data files to the template layout. 
        public void writeToExcel(string fileName, Excel._Worksheet oSheet, string sheetName)
        {
            readFiles readFile = new readFiles();

            List<string> listOfDataA;
            List<string> listOfDataB;
            List<string> listOfDataC;

            int cIndex1 = 1, cIndex2;
            int index = 1;

            // Assign new name to worksheet.
            oSheet.Name = sheetName;

            // Delete everything in that sheet. 
            oSheet.Cells.ClearContents();

            // clearing the columns for each sheet. 
            oSheet.Range["A1", "J1"].EntireColumn.Clear();

            // Assign each list to the file that is read in.
            listOfDataA = readFile.readInFile(fileName, 0);
            listOfDataB = readFile.readInFile(fileName, 1);
            listOfDataC = readFile.readInFile(fileName, 2);

            int rowIndexA = 1;
            int rowIndexB = 1;
            int rowIndexC = 1;

            // write the data from the lostofData to each cell. 
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

            // Write Average and Max values of column G
            oSheet.Cells[1, "I"] = "Average: ";
            oSheet.Cells[2, "I"] = "Max: ";
            oSheet.Range["J1"].Formula = "=Average(G:G)";
            oSheet.Range["J2"].Formula = "=Max(G:G)";
        }
    }
}
