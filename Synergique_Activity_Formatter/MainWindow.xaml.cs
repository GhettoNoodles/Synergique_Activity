using System.Windows;
using System.Windows.Navigation;
using System;
using System.Collections.Generic;
using System.Text;
using Bytescout.Spreadsheet;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Media.Animation;
using Bytescout.Spreadsheet.Constants;
using Synergique_Activity_Formatter.Core;
using System.Text.Json;

namespace Synergique_Activity_Formatter
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {
        public string newSheetPath;
        private List<Item> _items = new List<Item>();

        public MainWindow()
        {
            InitializeComponent();
            //  ReadData("C:\\Users\\jacqu\\RiderProjects\\Synergique_Activity\\Synergique_Activity_Formatter\\Copy of Activity Sumary 04-04-2023 Excel - Copy.xlsx");
        }


        private void Browse_OnClick(object sender, RoutedEventArgs e)
        {
            // Configure open file dialog box
            var dialog = new Microsoft.Win32.OpenFileDialog();
            dialog.FileName = "Document"; // Default file name
            dialog.DefaultExt = ".xlsx"; // Default file extension
            dialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm"; // Filter files by extension

// Show open file dialog box
            bool? result = dialog.ShowDialog();

// Process open file dialog box results
            if (result == true)
            {
                // Open document
                newSheetPath = dialog.FileName;
                ReadData();
            }
        }

        public void ReadData()
        {
            string currentCellTxt = "";
            int salesCol = 3;
            int currentLine = 2; // row of first item(SkipHeading)
            Spreadsheet document = new Spreadsheet();
            document.LoadFromFile(newSheetPath);
            Worksheet incomingData = document.Workbook.Worksheets.ByName("Data1");
            int TimeoutCounter = 0;
            while (TimeoutCounter < 15000)
            {
                TimeoutCounter++;
                var currentCell = incomingData.Cell(currentLine, 0);
                currentCellTxt = currentCell.ValueAsString;

                //item Name
                Item newItem = new Item();
                newItem.Name = currentCellTxt;
                Console.WriteLine(newItem.Name);

                //invoiced
                currentLine++;
                currentCell = incomingData.Cell(currentLine, 0);
                newItem.LastInvoiced = currentCell.ValueAsString;
                Console.WriteLine(newItem.LastInvoiced);
                //Purchased
                currentLine++;
                currentCell = incomingData.Cell(currentLine, 0);
                newItem.LastPurchased = currentCell.ValueAsString;
                Console.WriteLine(newItem.LastPurchased);

                //Sales
                while (currentCellTxt != "March")
                {
                    currentLine++;
                    currentCell = incomingData.Cell(currentLine, 0);
                    currentCellTxt = currentCell.ValueAsString;
                }

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

                //AverageSales
                newItem.AverageSales = (float)Math.Ceiling(sales.Sum() / 12f);
                Console.WriteLine("Average = " + newItem.AverageSales);
                _items.Add(newItem);
                currentLine += 2; //should be next inventory item
                currentCell = incomingData.Cell(currentLine, 0);
                if (currentCell.ValueAsString == "")
                {
                    break;
                }
            }
        }

        private void CompareAndSerializeJsons(List<Item>items)
        {
            
        }
        

        private void Save_OnClick(object sender, RoutedEventArgs e)
        {
            // Create new Spreadsheet
            Spreadsheet outputDoc = new Spreadsheet();

            // add new worksheet
            Worksheet orderSheet = outputDoc.Workbook.Worksheets.Add("FormulaDemo");
            
            //Columns for initial calculations
            orderSheet.Cell("A1").Value =
                "Reve Holdings SA (Pty) Ltd - Equipment  |  Prepared by:  Reve Holdings SA (Pty) Ltd";
            orderSheet.Rows[0].BottomBorderStyle = LineStyle.Thick;
            orderSheet.Columns[0].Width = 520;
            orderSheet.Cell("B1").Value = "12 Month Average";
            orderSheet.Columns[1].Width = 70;
            orderSheet.Columns[1].AlignmentHorizontal = AlignmentHorizontal.Left;
            orderSheet.Cell("C1").Value = "Minimum Level";
            orderSheet.Columns[2].Width = 70;
            orderSheet.Columns[2].AlignmentHorizontal = AlignmentHorizontal.Left;
            orderSheet.Cell("D1").Value = "Maximum Level";
            orderSheet.Columns[3].Width = 70;
            orderSheet.Columns[3].AlignmentHorizontal = AlignmentHorizontal.Left;
            orderSheet.Rows[0].Height = 40;
            orderSheet.Rows[0].Wrap = true;
            orderSheet.Rows[0].AlignmentVertical = AlignmentVertical.Centered;

            //Adding values for each item
            var nameLine = 1;
            foreach (var item in _items)
            {
                orderSheet.Cell(nameLine, 0).Value = item.Name;

                orderSheet.Cell(nameLine + 1, 0).Value = item.LastInvoiced;

                orderSheet.Cell(nameLine + 2, 0).Value = item.LastPurchased;
                orderSheet.Cell(nameLine + 2, 1).Value = item.AverageSales;
                orderSheet.Cell(nameLine + 2, 2).Value = item.AverageSales*3;
                orderSheet.Cell(nameLine + 2, 3).Value = item.AverageSales*6;
                orderSheet.Rows[nameLine + 2].BottomBorderStyle = LineStyle.Thin;
                
                
                nameLine += 3;//next item
                
            }

            // delete output file if exists already
            if (File.Exists("Output.xls"))
            {
                File.Delete("Output.xls");
            }

            // Save document
            outputDoc.SaveAs("Output.xls");

            // Close Spreadsheet
            outputDoc.Close();

            // open generated XLS document in default program
            Process.Start("Output.xls");
        }
    }
}