using System.Windows.Input;

namespace Synergique_Activity_Formatter.Core
{ using System;
    using System.Collections.Generic;
    using System.Text;
    using Bytescout.Spreadsheet;
    using System.Diagnostics;
    using System.IO;
    using System.Linq;
    using System.Windows.Media.Animation;
    using Synergique_Activity_Formatter.Core;
    
    public class ReadNewSheet
    {
        public string newSheetPath;
        
        public void BrowseFile()
        {
            // Configure open file dialog box
            var dialog = new Microsoft.Win32.OpenFileDialog();
            dialog.FileName = "Document"; // Default file name
            dialog.DefaultExt = ".xlsx"; // Default file extension
            dialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";// Filter files by extension

// Show open file dialog box
            bool? result = dialog.ShowDialog();

// Process open file dialog box results
            if (result == true)
            {
                // Open document
                newSheetPath= dialog.FileName;
                ReadData();
            }
        }
         public void ReadData()
        {
            string prevCellTxt = "";
            string currentCellTxt = "";
            int salesCol = 3;
            int onHandCol = 4;
            int currentLine = 2;// row of first item(SkipHeading)
            bool finishFlag = false;
            Spreadsheet document = new Spreadsheet();
            document.LoadFromFile(newSheetPath);
            Worksheet incomingData = document.Workbook.Worksheets.ByName("Data1");
            while (!finishFlag)
            {
                var currentCell = incomingData.Cell(currentLine, 0);
                currentCellTxt = currentCell.ValueAsString;
                
                //item Name
                Item newItem = new Item();
                newItem.Name = currentCellTxt;
                Console.WriteLine(newItem.Name);
                while (currentCellTxt!="March")
                {
                    currentLine++;
                    currentCell = incomingData.Cell(currentLine, 0);
                    currentCellTxt = currentCell.ValueAsString;
                }
                //Item salesArray
                float[] sales = new float[12];
                for (int i = 0; i < 12; i++)
                {
                    currentCell = incomingData.Cell(currentLine, salesCol);
                    sales[i] = currentCell.ValueAsInteger;
                    currentLine++;
                }
                newItem.Sales = sales;
                foreach (var val in newItem.Sales)
                {
                    Console.WriteLine(val);
                }
                newItem.AverageSales = sales.Sum() / 12;
                Console.WriteLine("Average = "+ newItem.AverageSales);
                //currentStock
                //currentLine++;//unposted
                //currentCell = incomingData.Cell(currentLine, onHandCol);
               // newItem.CurrentStock = currentCell.ValueAsInteger;
                
                currentLine+=2;//should be next inventory item
                currentCell = incomingData.Cell(currentLine, 0);
                if (currentCell.ValueAsString == "")
                {
                    finishFlag = true;
                    break;
                }
            }
        }
    }
}