using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows;
using Bytescout.Spreadsheet;
using Bytescout.Spreadsheet.Constants;

namespace Synergique_Activity_Formatter.Core
{
    public class ExcelWriter
    {
        public List<Item> WriteOutput(int currentMonth, List<Item> _recentDataItems, List<Item> _oldDataItems)
        {
            
            if (currentMonth == 0)
            {
                MessageBox.Show("Please enter a valid Month");
                return null;
            }

            // Create new Spreadsheet
            Spreadsheet outputDoc = new Spreadsheet();

            // add new worksheet
            Worksheet orderSheet = outputDoc.Workbook.Worksheets.Add("FormulaDemo");

            //Columns for initial calculations
            orderSheet.Cell("A1").Value =
                "Reve Holdings SA (Pty) Ltd - Equipment  |  Prepared by:  Reve Holdings SA (Pty) Ltd";
            orderSheet.Rows[0].BottomBorderStyle = LineStyle.Thick;
            orderSheet.Columns[0].Width = 520;
            orderSheet.Cell("B1").Value = "6 Month Average";
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
            int itemCounter = 0;
            foreach (var item in _recentDataItems)
            {
                orderSheet.Cell(nameLine, 0).Value = item.Name;

                orderSheet.Cell(nameLine + 1, 0).Value = item.LastInvoiced;

                orderSheet.Cell(nameLine + 2, 0).Value = item.LastPurchased;

                int[] sales = new int[6];
                //combines old data and new data
                if (currentMonth < 6)
                {
                    for (int i = 0; i < currentMonth; i++)
                    {
                        sales[i] = (int)item.Sales[i];
                    }

                    int saleItemCounter = currentMonth;
                    for (int i = 11; i > 11 - 6 + currentMonth; i--)
                    {
                        sales[saleItemCounter] = (int)_oldDataItems[itemCounter].Sales[i];
                        saleItemCounter++;
                    }
                }
                else // uses most recent 6 months
                {
                    int saleItemCounter = 0;
                    for (int i = currentMonth - 6; i < currentMonth; i++)
                    {
                        sales[saleItemCounter] = (int)item.Sales[i];
                        saleItemCounter++;
                    }
                }

                var avgSales = Math.Ceiling(sales.Sum() / 6f);
                orderSheet.Cell(nameLine + 2, 1).Value = avgSales;
                orderSheet.Cell(nameLine + 2, 2).Value = avgSales * 3;
                orderSheet.Cell(nameLine + 2, 3).Value = avgSales * 6;
                orderSheet.Rows[nameLine + 2].BottomBorderStyle = LineStyle.Thin;
                nameLine += 3; //next item
                itemCounter++;
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
            return _recentDataItems;
        }
    }
}