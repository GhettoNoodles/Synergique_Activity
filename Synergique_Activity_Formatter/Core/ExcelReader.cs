using System;
using System.Collections.Generic;
using System.Linq;
using Bytescout.Spreadsheet;

namespace Synergique_Activity_Formatter.Core
{
    public class ExcelReader
    {
        public List<Item> ReadData(List<Item> savedItems, string dataPath,JsonManager jsonManager,bool newData)
        {
            List<Item> newlyReadItems = new List<Item>();
            string currentCellTxt = "";
            int salesCol = 3;
            int currentLine = 2; // row of first item(SkipHeading)
            Spreadsheet document = new Spreadsheet();
            document.LoadFromFile(dataPath);
            Worksheet incomingData = document.Workbook.Worksheets.ByName("Data1");
            int timeoutCounter = 0;
            while (timeoutCounter < 15000)
            {
                timeoutCounter++;
                var currentCell = incomingData.Cell(currentLine, 0);
                currentCellTxt = currentCell.ValueAsString;

                //item Name
                Item newItem = new Item();
                newItem.Name = currentCellTxt;

                //invoiced
                currentLine++;
                currentCell = incomingData.Cell(currentLine, 0);
                newItem.LastInvoiced = currentCell.ValueAsString;
                //Purchased
                currentLine++;
                currentCell = incomingData.Cell(currentLine, 0);
                newItem.LastPurchased = currentCell.ValueAsString;

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
                //AverageSales
                newItem.AverageSales = (float)Math.Ceiling(sales.Sum() / 12f);
               // Console.WriteLine("Average = " + newItem.AverageSales);
                newlyReadItems.Add(newItem);
                currentLine += 2; //should be next inventory item
                currentCell = incomingData.Cell(currentLine, 0);
                if (currentCell.ValueAsString == "")
                {
                    break;
                }
            }

            if (newData)
            {
                var index = 0;
                
                while (index<newlyReadItems.Count)
                {
                    if (newlyReadItems[index].Name!=savedItems[index].Name)
                    {
                       var newItem = newlyReadItems[index];
                        newItem.Sales= new float[]{0,0,0,0,0,0,0,0,0,0,0,0};
                        savedItems.Insert(index,newItem);
                    }
                    index++;
                }
                jsonManager.SerializeToJson(savedItems,"oldItems.txt");
                Console.WriteLine(savedItems.Count + ";" + newlyReadItems.Count);
            }
            return newlyReadItems;
        }
    }
}