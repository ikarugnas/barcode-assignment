using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace Barcode_Assignment
{
    class Program
    {
        static void Main(string[] args)
        {
            Program myP = new Program();
            myP.Start();
        }

        private void Start()
        {
            //Dynamic interpretation was attempted but failed. File wouldn't recognize on my PC despite being in the right directory.
            //excel.Workbooks.Open(filepath) would not accept dynamic filepath format.
            //basePath is in the \bin\ folder. All you have to do for the code to run is to change the base path to the \bin\ folder on the device.
            string basePath = @"C:\Users\Enes\source\repos\Barcode Assignment\Barcode Assignment\bin\";
            string inputPath = basePath + "input.xlsx";
            string outputPath = basePath + "output.xlsx";
            string MAP1Path = basePath + @"MAP1\";
            string MAP2Path = basePath + @"MAP2\";

            List<StoreItem> items = ReadExcel(inputPath);
            items = TransferPictures(items, MAP1Path, MAP2Path);
            CreateOutputFile(items, outputPath);

            Logger.WriteLog($"Files succesfully extracted, copied from {MAP1Path} to {MAP2Path}, and created in new excel file named output.xlsx at {outputPath}.");
        }

        private List<StoreItem> ReadExcel(string filepath)
        {
            List<StoreItem> items = new List<StoreItem>();
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            excel.Visible = false;
            Workbook wb;
            Worksheet ws;

            wb = excel.Workbooks.Open(filepath);
            ws = wb.Worksheets[1];

            for (int i = 2; i < ws.UsedRange.Rows.Count; i++)
            {
                StoreItem item = new StoreItem();
                int coll = 1;
                Range cell = ws.Cells[i, coll];
                if (cell.Value.ToString().Length != 13)
                {
                    Logger.WriteLog($"Wrong barcode format detected at: [{i}, {coll}]");
                    continue;
                }
                Logger.WriteLog($"Barcode found: {cell.Value}");
                item.Barcode = cell.Value;
                items.Add(item);
            }

            excel.Workbooks.Close();
            excel.Quit();
            return items;
        }
        
        private List<StoreItem> TransferPictures(List<StoreItem> items, string MAP1Path, string MAP2Path)
        {
            string[] files = System.IO.Directory.GetFiles(MAP1Path);
            int copiedFiles = 0;

            if (!Directory.Exists(MAP2Path))
            {
                Directory.CreateDirectory(MAP2Path);
            }

            foreach (StoreItem item in items)
            {
                item.FileNames = new List<string>();
                foreach (string file in files)
                {
                    if (file.Contains(item.Barcode.ToString()))
                    {
                        string fileName = System.IO.Path.GetFileName(file);
                        string destinationFile = MAP2Path + fileName;
                        try
                        {
                            

                            System.IO.File.Copy(file, destinationFile);
                            item.FileNames.Add(fileName);
                            Logger.WriteLog($"{fileName} was successfully added as a file to {destinationFile}");
                            copiedFiles++;
                        }
                        catch (System.IO.IOException e)
                        {
                            
                            Logger.WriteLog(e.Message);
                            item.FileNames.Add(fileName);
                        }

                    }
                }
            }

            Logger.WriteLog($"Copied {copiedFiles} files successfully from {MAP1Path} to {MAP2Path}.");
            return items;
        }

        private void CreateOutputFile(List<StoreItem> items, string outputPath)
        {
            /*Code to delete previous output file.
              If this is not present the program will always
              ask for permission to overwrite the file if it exists.*/
            //if(File.Exists(outputPath))
            //{
            //    File.Delete(outputPath);
            //}

            int largestAmountOfFiles = LargestAmountOfFiles(items);

            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            excel.Visible = false;
            Workbook wb;
            Worksheet ws;

            wb = excel.Workbooks.Add();
            ws = wb.Worksheets[1];

            GenerateFirstRow(ws, largestAmountOfFiles);
            GenerateBarcodeAndFileData(ws, items);

            try
            {
                wb.SaveAs(outputPath);
                wb.Close();
                excel.Quit();
                Logger.WriteLog($"Saved file successfully to {outputPath}");
            }
            catch (System.Runtime.InteropServices.COMException e)
            {
                Logger.WriteLog(e.Message);
                wb.Close();
                excel.Quit();
            }
        }

        private int LargestAmountOfFiles(List<StoreItem> items)
        {
            int largestAmountOfFiles = 0;

            foreach (StoreItem item in items)
            {
                if (item.FileNames.Count > largestAmountOfFiles)
                {
                    largestAmountOfFiles = item.FileNames.Count;
                }
            }
            Logger.WriteLog($"Largest amount of Files detected is {largestAmountOfFiles}.");

            return largestAmountOfFiles;
        }

        private void GenerateFirstRow(Worksheet ws, int largestAmountOfFiles)
        {
            //Generation of the first row.
            Range initialCell = ws.Range["A1:A1"];
            initialCell.Value = "Ean";
            for (int i = 1; i <= largestAmountOfFiles; i++)
            {
                int iPlusOne = i + 1;
                Range topRow = ws.Cells[1, iPlusOne];
                topRow.Value = "plaatje" + i;
            }

            Logger.WriteLog("First row generated successfully");
        }

        private void GenerateBarcodeAndFileData(Worksheet ws, List<StoreItem> items)
        {
            /*For loop where int i goes through all the items and then loops through the FileNames list of the
              store item, whereafter it adds this to its designated cell.*/
            for (int i = 1; i <= items.Count; i++)
            {
                Range firstCol = ws.Cells[i + 1, 1];
                firstCol.Value = items[i - 1].Barcode;

                for (int j = 0; j <= items[i - 1].FileNames.Count - 1; j++)
                {
                    Range targetCell = ws.Cells[i + 1, j + 2];
                    targetCell.Value = items[i - 1].FileNames[j];
                }
            }

            Logger.WriteLog("Data added successfully");
        }
    }
}