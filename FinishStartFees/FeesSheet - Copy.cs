﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Syncfusion.XlsIO;
using Syncfusion.Windows.Shared;

namespace FinishStartFees
{  public struct DateandValue
    { public string txnDate { get; set; }
        public decimal txnAmt { get; set; }

    }
    public struct feeColsStruct
    {
        public int balanceCol { get; set; }
        public int txnDateCol { get; set; }
        public int txnSeqCol { get; set; }
        public int feeCol { get; set; }
        public int payCol { get; set; }
        public int sumaCol { get; set; }
        public int propIndexCol { get; set; }
        public int propNameCol { get; set; }
   //     public feeColsStruct(int p1, int p2, int p3) {
    //        balanceCol 
    //    }
    }

    class FeesSheet
    {
       // public Dictionary<string, int> feeCols { get; set; }
        public Dictionary<long, Dictionary<string, Object>> fileScan { get; set; }//= new Dictionary<string, Object>();

        public  string fileName { get; set; }
        public static string fileName1 { get; set; }
        public static string fileName2 { get; set; }
        public feeColsStruct feeCols = new feeColsStruct();
        public FeesSheet()        {
            // feeCols = new Dictionary<string, int>();
            
            feeCols.balanceCol = 0;
            feeCols.feeCol = 0;
            feeCols.payCol = 0;
            feeCols.propIndexCol = 0;
            feeCols.propNameCol = 0;
            feeCols.sumaCol = 0;
            feeCols.txnDateCol = 0;
            feeCols.txnSeqCol = 0;



            fileScan = new Dictionary<long, Dictionary<string, Object>>();
            // indexed on propertyID  43000..... etc, 
            //internal dictionary for propertyName , startRow, endRow, startbalance, endbalance

        }

        public void getvariousCols(IWorksheet sheet)
        {
            // return  ints for the columns for fees and balance taken as the 2 decimals on a "RECIBO DEL" row and adjusted for zero base column numbers
            IRange reciboRow = sheet.FindFirst("RECIBO DEL", ExcelFindType.Text).EntireRow;
            IRange[] resultAll = reciboRow.FindAll(0, ExcelFindType.Number); 
            if (resultAll.Length != 2)
            {
                MessageBox.Show("Problem Not two decimals on the chosen row for RECIBO DEL analysis on row number" + reciboRow.Row, "Problem ..  aborting");
                Application.Current.MainWindow.Close();
            }
            // get columns for fees, payments, balance and activity date column
            //convert column number to zero based
          
            feeCols.feeCol = resultAll[0].Column;
            feeCols.balanceCol = resultAll[1].Column;
            // the date column is the first text field of the row
            feeCols.txnDateCol = reciboRow.FindFirst("01/01", ExcelFindType.Text).Column;
            //The txnSeqCol is the next but 1 number after the date.
            //feeCols.txnSeqCol = findnumberaftertext(reciboRow, "01/01")[1];
            // to find the next but 1 number after date its space padded so have to trim before testing for number
            int numberNumber = 1;
            int aNumber;
            foreach (IRange aCell in reciboRow.Cells)
            {
                
                { if(((aCell.HasString)||(aCell.HasNumber)) &&(    int.TryParse(aCell.DisplayText.Trim(),out aNumber))) {


                    if (numberNumber == 2)
                    {
                        feeCols.txnSeqCol = aCell.Column;
                        break;
                    }
                    
                    numberNumber++;
               } }
            }

            //now find columns for property index and property name
            feeCols.propIndexCol = sheet.FindFirst("4300", ExcelFindType.Text).Column;


            //feeCols.propNameCol = sheet.FindFirst("COBRO REC", ExcelFindType.Text).Column;// TODO: make a function, find numberaftertext
            feeCols.payCol = findnumberaftertext(sheet, "COBRO REC")[0];
            feeCols.propNameCol = sheet.FindFirst("PARC.", ExcelFindType.Text).Column;
            //Don't actually need to look for the balance cos already acquired it - 
            feeCols.sumaCol = sheet.FindFirst("Suma total", ExcelFindType.Text).Column;


            return;
        }

        //Find the column numbers of any numbers in the columns after the search text
        public int[] findnumberaftertext(IWorksheet sheet, string texttofind)
        {
            int[] thenumbersCol = new int[2];
            // thenumbersCol[0] = 99;

            //find row containing the text
            IRange anumberrow = sheet.FindFirst(texttofind, ExcelFindType.Text);
            // find the column for the text so as ToString look after it 1-based
            int itsatColumn = anumberrow.Column;
            int itsatRow = anumberrow.Row;
            // get the cells for the row and find a number after the column satisfying the search
            IRange[] thecells = anumberrow[itsatRow, itsatColumn + 1, itsatRow, 255].Cells;
            //   IRange[] somenumbers =anumberrow[itsatRow,itsatColumn+1,itsatRow,255].FindAll(".", ExcelFindType.Text);
            int indx = 0;
            foreach (IRange acell in thecells)
            {
                if (acell.HasNumber)
                {
                    thenumbersCol[indx] = acell.Column;
                    indx++;

                }
            }
            return thenumbersCol;
        }
        public IWorksheet getSheetfromFile(string feefileName)
        {   //IWorksheet sheet;
            fileName = feefileName;
            string currdir;
            //New instance of XlsIO is created.[Equivalent to launching MS Excel with no workbooks open].
            //The instantiation process consists of two steps.

            //Step 1 : Instantiate the spreadsheet creation engine.
            ExcelEngine excelEngine = new ExcelEngine();

            //Step 2 : Instantiate the excel application object.
            IApplication application = excelEngine.Excel;
            currdir = System.IO.Directory.GetCurrentDirectory();
            IWorkbook workbook = application.Workbooks.OpenReadOnly(feefileName);


            return workbook.Worksheets[0];
        }
        public void scanFeeFile(IWorksheet sheet)
        {// scan through the file acquiring propertID, propertyName, start balances and finish balances
         // first run through file1 get all the start and finish row numbers for each property, used later to find Asiento ownership
         // set the sheet range to only include the PropertyID and search for then all
            int lastRow = sheet.Range.LastRow;
            int propIndexCol = feeCols.propIndexCol;
            int  propNameCol = (feeCols.propNameCol-1);
            IRange[] propIndices = sheet.Range[1, propIndexCol, lastRow, propIndexCol].FindAll("4300", ExcelFindType.Text);
            //filescan dictionary 
            // indexed on propertyID  43000..... etc, 
            //internal dictionary for propertyName , startRow, endRow, startbalance, endbalance
            long aa;
            //Run through the file getting all the propertyindex rows and saving the ID as keys to propIndices Dictionary
            foreach (IRange prop in propIndices)
            {
                if (long.TryParse(prop.DisplayText, out aa))
                {
                    fileScan.Add(aa, new Dictionary<string, object>());
                }
                else
                {
                    MessageBox.Show("Problem PropertyIndex not convertible to long type converting " + prop.DisplayText, "Problem ..  aborting");
                    Application.Current.MainWindow.Close();
                }


                fileScan[aa].Add("startRow", prop.Row + 1);
                fileScan[aa].Add("propName", prop.EntireRow.Cells[propNameCol].DisplayText);

            }
            //Now scan for name of property in startRow and end of each property which is given by first sumaCol after the row given by startRow
            //then we have start and end of each property entry and we can get all the financial data for each property
            IRange foundRow;

            foreach (long prop in fileScan.Keys)
            {  //First get the endRows and and later go through again for starting balance, all fees and all payments and endbalance


                foundRow = sheet.Range[(int)fileScan[prop]["startRow"], feeCols.sumaCol, lastRow, feeCols.balanceCol].FindFirst("Suma total", ExcelFindType.Text);//.EntireRow;
                fileScan[prop]["endRow"] = foundRow.Row;
            }
            //startRow depends on whether there is a ASIENTO DE APERTURA Row
            // If no ASIENTA startRow is 1 more than the propert row
            //if there is an ASSIENTA row then startRow is 1 more than that

            foreach (long prop in fileScan.Keys)
            {  //go through again for starting balance (if any) , all fees and all payments and endbalance
                ///look for any opening balance
                foundRow = sheet.Range[(int)fileScan[prop]["startRow"], 1, (int)fileScan[prop]["endRow"], feeCols.balanceCol].FindFirst("ASIENTO DE APERTURA", ExcelFindType.Text);//.EntireRow;
                if ((foundRow != null))
                {
                    //  fileScan[prop]["startbalance"] =                        foundRow.Row;
                    fileScan[prop]["startBalance"] = (decimal) sheet[foundRow.Row, feeCols.balanceCol].Number;
                    //startRow depends on whether there is a ASIENTO DE APERTURA Row
                    // If no ASIENTA startRow is 1 more than the propert row
                    //if there is an ASSIENTA row then startRow is 1 more than that
                    fileScan[prop]["startRow"] = foundRow.Row + 1;
                }
                else
                {
                    fileScan[prop]["startBalance"] = (decimal)0.0;
                }
                //Now all fees and all payments and endbalance
                // ******Some versions do not have the sum of fees, payments and balances  so recalc them myself

                // foundRow = sheet.Range[(int)fileScan[prop]["startRow"], 1, (int)fileScan[prop]["endRow"], feeCols.balanceCol].FindFirst("Suma total", ExcelFindType.Text);//.EntireRow;
                fileScan[prop]["feeTotal"]  = sumofNumberCols(sheet, prop, feeCols.feeCol);
                fileScan[prop]["payTotal"] = sumofNumberCols(sheet, prop, feeCols.payCol);
                fileScan[prop]["finishBalance"] = (decimal)fileScan[prop]["startBalance"] +(decimal) fileScan[prop]["feeTotal"] - (decimal) fileScan[prop]["payTotal"];

                fileScan[prop].Add("fees", new Dictionary<string, decimal>());
                fileScan[prop].Add("payments", new Dictionary<string, decimal>());
                fileScan[prop]["fees"] = detailNumberCols(sheet, prop, feeCols.feeCol);
                fileScan[prop]["payments"] = detailNumberCols(sheet, prop, feeCols.payCol);

                // foundRow = sheet[foundRow.Row, feeCols.balanceCol];//.Number;
                //if (foundRow.IsBlank) { fileScan[prop]["endbalance"] = 0; }
                //else { fileScan[prop]["endbalance"] = foundRow.Number; }


            }

            //Now have startrow, endrow,startbalance,endbalance for each property

            //Now get the sum of all the fees and all the balance for each property from the 


            return;
        }


        /*********
       * get all the fees and payments entries, in detail 
       * create a dictionary to add into fileScan[property] after the return
       * the key of the dictionary is the transaction date
         ***************/

        public List<DateandValue> detailNumberCols(IWorksheet sheet, long property, int colNum) {
            //Dictionary<string, decimal> theNumbers = new Dictionary<string, decimal>();
            List<DateandValue> theNumbers = new List<  DateandValue >();
           DateandValue nums = new DateandValue();

            string txnDate;
            IRange[] setofNumberCols = sheet.Range[(int)fileScan[property]["startRow"],
                 colNum, (int)fileScan[property]["endRow"] - 1, colNum].Cells;
            int txnDateCol = (int)  feeCols.txnDateCol-1;   // cells is zero based
           // int txnSeqCol= (int)feeCols.txnSeqCol - 1;   // cells is zero based
            foreach (IRange numberCol in setofNumberCols)
            {
                if (numberCol.HasNumber)
                {
                    txnDate = numberCol.EntireRow.Cells[txnDateCol].DisplayText;// +":"+ numberCol.EntireRow.Cells[txnSeqCol].DisplayText;  // cells is zero based
                    nums.txnDate = txnDate;
                    nums.txnAmt = ( decimal) numberCol.Number;
                    theNumbers.Add( nums);
                }

            }

            return theNumbers;
                    }

        /*********
         * get all the fees and payments entries,
         *  as totals
        ***************/

        public decimal sumofNumberCols(IWorksheet sheet, long property, int colNum) // TODO: also save individual entries for the simolify task leading to OCM exctract
        {
            IRange[]  setofNumberCols = sheet.Range[(int)fileScan[property]["startRow"], 
                colNum, (int)fileScan[property]["endRow"]-1, colNum].Cells;
            decimal colTotal = 0;
            foreach (IRange numberCol in setofNumberCols)
            {
                if (numberCol.HasNumber) colTotal += (decimal) numberCol.Number;

            }
            return colTotal;
        }
    }
}
