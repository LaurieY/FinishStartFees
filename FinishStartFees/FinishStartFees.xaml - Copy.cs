using System;
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


        private void button_Click(object sender, RoutedEventArgs e)
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

            string resultText = "Properties with Discrepancies\n\n";
            foreach (long prop in sheet1.fileScan.Keys) {
              if (!sheet1.fileScan[prop]["finishBalance"].Equals (sheet2.fileScan[prop]["startBalance"] ))
                {
                    resultText += "\n "+ prop + " " + sheet1.fileScan[prop]["propName"] + " Close Balance " + sheet1.fileScan[prop]["finishBalance"] + " Opening Balance " +
                         sheet2.fileScan[prop]["startBalance"] + "\n";
                    long propp = prop;
                    Results.Text = resultText;
                }


            }
            MessageBox.Show("Finished");


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
    }


}
