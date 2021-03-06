﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Syncfusion.XlsIO;
using Syncfusion.Windows.Shared;
using Microsoft.Win32;
using System.IO;

namespace FinishStartFees
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>

    public partial class MainWindow : Window
    {
        //  public int[] variousCols =new int[2];
        public Dictionary<string, int> variousCols = new Dictionary<string, int>();
        //   public Dictionary<string, int> file1Cols = new Dictionary<string, int>();
        //  public Dictionary<string, int> file2Cols = new Dictionary<string, int>();
        feeColsStruct file1Cols = new feeColsStruct();
        feeColsStruct file2Cols = new feeColsStruct();

        public OpenFileDialog feeFile1;
        public MainWindow()
        {
            InitializeComponent();
        }
        protected virtual bool IsFileLocked(string fileName)
        {
            FileInfo file = new FileInfo(fileName);
            FileStream stream = null;

            try
            {
                stream = file.Open(FileMode.Open, FileAccess.Read, FileShare.None);
            }
            catch (IOException)
            {
                //the file is unavailable because it is:
                //still being written to
                //or being processed by another thread
                //or does not exist (has already been processed)
                return true;
            }
            finally
            {
                if (stream != null)
                    stream.Close();
            }

            //file is not locked
            return false;
        }
        // return 2 ints for the columns for fees and balance taken as the 2 doubles on a "RECIBO DEL" row and adjusted for zero base column numbers
        //public void getvariousCols(IWorksheet sheet) {
        //    // return 2 ints for the columns for fees and balance taken as the 2 doubles on a "RECIBO DEL" row and adjusted for zero base column numbers
        //    IRange reciboRow = sheet.FindFirst("RECIBO DEL", ExcelFindType.Text).EntireRow;
        //    IRange[] resultAll = reciboRow.FindAll(0, ExcelFindType.Number);
        //    if (resultAll.Length != 2)
        //    {
        //        MessageBox.Show("Problem Not two doubles on the chosen row for RECIBO DEL analysis on row number" + reciboRow.Row, "Problem ..  aborting");
        //        Application.Current.MainWindow.Close();

        //    }
        //    // get columns for fees, balance and activity date column
        //    //convert column number to zero based
        //    variousCols.feeCol = resultAll[0].Column-1;
        //    variousCols["balanceCol"] = resultAll[1].Column-1;
        //    // the date column is the first text field of the row
        //    variousCols["dateCol"] = reciboRow.FindFirst("01/01", ExcelFindType.Text).Column - 1;
        //    //now find columns for property index and property name
        //    variousCols.propIndexCol = sheet.FindFirst("4300", ExcelFindType.Text).Column - 1;

        //    variousCols.propNameCol = sheet.FindFirst("PARC.", ExcelFindType.Text).Column-1;

        //    variousCols.sumaCol = sheet.FindFirst("Suma total", ExcelFindType.Text).Column - 1;


        //    return; 
        //}
        // get the column number (-1) for the initial date column -  Usually A  i.e. zero


        private void bCompare_Click(object sender, RoutedEventArgs e)
        {
            using (new WaitCursor())
            {
                IWorksheet sheet;
                //string currdir;
                // //New instance of XlsIO is created.[Equivalent to launching MS Excel with no workbooks open].
                // //The instantiation process consists of two steps.

                // //Step 1 : Instantiate the spreadsheet creation engine.
                // ExcelEngine excelEngine = new ExcelEngine();

                ////Step 2 : Instantiate the excel application object.
                //IApplication application = excelEngine.Excel;
                //currdir = System.IO.Directory.GetCurrentDirectory();
                //IWorkbook workbook = application.Workbooks.Open(@"..\..\MAYORES A 31.12.15.XLS");
                //IWorksheet sheet = workbook.Worksheets[0];

                //IRange result;
                //IRange[] resultAll, cells1;
                // int saldoCol, feeCol, paymentCol, balanceCol, recibodelCol;
                //   int[] miscCols;



                FeesSheet sheet1 = new FeesSheet();
                sheet1.fileName = FeesSheet.fileName1;
                //  sheet = sheet1.getSheetfromFile(@"..\..\MAYORES A 31.12.15.XLS");
                sheet = sheet1.getSheetfromFile(sheet1.fileName);
                sheet1.getvariousCols(sheet);
                // getvariousCols(sheet);
                file1Cols = sheet1.feeCols;
                sheet1.scanFeeFile(sheet);

                FeesSheet sheet2 = new FeesSheet();
                sheet2.fileName = FeesSheet.fileName2;
                sheet = sheet2.getSheetfromFile(sheet2.fileName);
                sheet2.getvariousCols(sheet);
                // getvariousCols(sheet);
                file2Cols = sheet2.feeCols;
                sheet2.scanFeeFile(sheet);


                //sheet.GetText(13, 14);
                //// Now run through file1 get all the start and finish row numbers for each property, used later to find Asiento ownership
                //IRange one =sheet.Range["A13:z13"];
                //int R1, C1, R2, C2;
                //R1 = 11; R2 = 22;
                //C1 = 1;C2 = 26;
                ////one = sheet.Range[15, 1, 15, 41];
                //one = sheet.Range[R1,C1,R2,C2];
                //IRange arange =sheet.Range[15, 4];
                //result = sheet.FindFirst("ASIENTO DE APERTURA", ExcelFindType.Text);




                //double asientoValue = result.EntireRow.Cells[variousCols.balanceCol].Number;
                bool discrepancies = false;
                string resultText = "Properties with Discrepancies\n\n";
                foreach (long prop in sheet1.fileScan2.Keys)
                {
                    if (!sheet1.fileScan2[prop].finishBalance.Equals(sheet2.fileScan2[prop].startBalance))
                    {
                        discrepancies = true;
                        resultText += "\n " + prop + "\t" + sheet1.fileScan2[prop].propName + "\n Close Balance \t" + sheet1.fileScan2[prop].finishBalance + "\t Opening Balance \t" +
                             sheet2.fileScan2[prop].startBalance + "\n";
                        long propp = prop;
                        ResultsScroll.Content = resultText;
                    }


                }
                if (!discrepancies) { resultText = "No Discrepancies"; }
                MessageBox.Show("Finished");
            }

        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog feeFile1 = new OpenFileDialog();
            feeFile1.Title = "Open FeeFile";
            if (feeFile1.ShowDialog() == true) FeesSheet.fileName1 = feeFile1.FileName;
            textBox1.Text = System.IO.Path.GetFileName(feeFile1.FileName);
        }

        private void button2_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog feeFile2 = new OpenFileDialog();
            feeFile2.Title = "Open FeeFile2";
            if (feeFile2.ShowDialog() == true) FeesSheet.fileName2 = feeFile2.FileName;
            textBox2.Text = System.IO.Path.GetFileName(feeFile2.FileName);
        }
        /**********
         * For feefile1  analyse and produce output containing a summary for each property
         * 
         *  Header with FileName and 1st and last dates covered by the file
         *
         *  Each line is property ID, Name, Brought Forward Amt, Fees Total, Paid Total, Carried Forward Amt
         * ************/
        private void bSummarise_Click(object sender, RoutedEventArgs e)
        {
            using (new WaitCursor())
            {


                IWorksheet sheet, sheetOut;
                FeesSheet sheet1 = new FeesSheet();
                sheet1.fileName = FeesSheet.fileName1;
                using (new WaitCursor())
                {
                    // very long task
                }
                sheet = sheet1.getSheetfromFile(sheet1.fileName);
                sheet1.getvariousCols(sheet);

                file1Cols = sheet1.feeCols;
                sheet1.scanFeeFile(sheet);

                ExcelEngine excelEngine = new ExcelEngine();
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Excel2013;
                string outFeeFileName;

                string currdir = System.IO.Path.GetDirectoryName(sheet1.fileName);
                string shortFileName = System.IO.Path.GetFileNameWithoutExtension(sheet1.fileName);

                outFeeFileName = currdir + "\\" + shortFileName + "-Summarised.xlsx";
                IWorkbook workbook = application.Workbooks.Create(1);
                sheetOut = workbook.Worksheets[0];

                object[,] array2 = sheet1.summariseFeeFile(sheet1);
                // object[] array = new object[4] { "Total Income", "Actual Expense", "Expected Expenses", "Profit" };


                sheetOut.ImportArray(array2, 1, 1);




                sheetOut.PageSetup.PrintTitleColumns = "$1:$1";
                sheetOut.Range["A1:F1"].AutofitColumns();
                sheetOut.AutofitColumn(2);
                sheetOut.PageSetup.CenterHeader = @"&""Gothic,bold""Summary of " + shortFileName;
                sheetOut.PageSetup.RightFooter = @"&""Gothic,bold""&D";
                sheetOut.PageSetup.CenterFooter = @"&""Gothic,bold""Page &P";
                sheetOut.PageSetup.Orientation = ExcelPageOrientation.Landscape;
                sheetOut.PageSetup.TopMargin = 0.59;
                sheetOut.PageSetup.BottomMargin = 0.59;
                sheetOut.PageSetup.HeaderMargin = 0.32;
                sheetOut.PageSetup.FooterMargin = 0.32;
                // Check if File is open already in Excel


                while (IsFileLocked(outFeeFileName))
                {
                    MessageBoxResult result = MessageBox.Show("Output file " + outFeeFileName + "\n is open in another application" + "\nPlease Close it", "Attention", MessageBoxButton.OKCancel);
                    if (result == MessageBoxResult.Cancel)
                    {
                        Application.Current.Shutdown();
                        return;
                    }

                }


                workbook.SaveAs(outFeeFileName);
                workbook.Close();
                excelEngine.Dispose();
                MessageBox.Show("Finished writing Summary file\n " + outFeeFileName, "Success");


            }
        }
    }


}
