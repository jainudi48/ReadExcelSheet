using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections;

namespace ReadExcelSheet
{
    public partial class Form1 : Form
    {
        Excel.Application xlApp;
        Excel.Workbook xlWorkbook;
        Excel._Worksheet xlWorksheet;
        Excel.Application xlAppInvalidOptions;
        Excel.Workbook xlWorkbookInvalidOptions;
        Excel._Worksheet xlWorksheetInvalidOptions;
        Excel.Range xlRange;
        int rowCount;
        int colCount;
        Array dummyPricesArray;
        string[] dummyPricesStrArray;

        public Form1()
        {
            InitializeComponent();
            InitializeObjects();
            //xlWorkbookInvalidOptions = createInvalidOptionsExcel();
            //readExcelDump();
            filterDummyPriceValues();
        }

        
        public void InitializeObjects()
        {
            xlApp = new Excel.Application();

            xlApp.Visible = true;
            try
            {
                xlWorkbook = xlApp.Workbooks.Open(@"C:\Users\U_jain\Documents\Visual Studio 2013\Projects\IT_27thJun_Q2Wk8.xlsx");
            }
            catch(Exception e)
            {
                MessageBox.Show("Error in opening excel file!");
            }

            try
            {
                xlWorksheet = xlWorkbook.Sheets[1];
            }
            catch(Exception e)
            {
                MessageBox.Show("Expected worksheet not found!");
            }

            xlRange = xlWorksheet.UsedRange;

            rowCount = xlWorksheet.UsedRange.Rows.Count;
            colCount = xlWorksheet.UsedRange.Columns.Count;

            dummyPricesStrArray = new string[]{ "999999.99", "9999999.99", "" };
            dummyPricesArray = dummyPricesStrArray;

        }

        /*public Excel.Workbook createInvalidOptionsExcel()
        {
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                Console.WriteLine("EXCEL could not be started. Check that your office installation and project references are correct.");
                return null;
            }
            xlApp.Visible = true;

            Excel.Workbook wb = xlApp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
            Excel._Worksheet ws = wb.Worksheets[1];

            if (ws == null)
            {
                Console.WriteLine("Worksheet could not be created. Check that your office installation and project references are correct.");
            }

            return wb;
        }


        public void readExcelDump()
        {
            String strCellValue = "";
            for(int i = 2; i <= rowCount; i++)
            {
                try
                {
                    strCellValue = xlRange.Cells[i, colCount].Value2.ToString();
                }
                catch(Microsoft.CSharp.RuntimeBinder.RuntimeBinderException e)
                {
                    // BLANK price values (CODE FOR INVALID OPTION)
                    writeInvalidOptions(i);
                }

                if (Convert.ToDouble(strCellValue) <= 0)
                {
                    // Negative price values (CODE FOR INVALID OPTION)
                    writeInvalidOptions(i);
                }
                else
                {
                    for (int j = 0; j < dummyPricesArray.Length; j++)
                    {
                        if (strCellValue == dummyPricesStrArray[j])
                        {
                            // Dummy price values (CODE FOR INVALID OPTION)
                            writeInvalidOptions(i);
                        }
                    }
                }
                
            }
            xlWorkbookInvalidOptions.SaveAs("C:\\Users\\U_jain\\Documents\\Visual Studio 2013\\Projects\\InvalidOptionsSheet.xlsx");
            Console.Write("Completed!");
        }

        public void writeInvalidOptions(int iRow)
        {
            for (int j = 1; j <= colCount; j++ )
            {
                try
                {
                    xlRange.Cells[iRow, j].Value2.ToString();
                    //Console.Write(xlRange.Cells[iRow, j].Value2.ToString());
                    xlWorkbookInvalidOptions.Worksheets[1].Cells[iXLWorksheetInvalidoptions, j] = xlRange.Cells[iRow, j].Value2;
                    arrInvalidOptionsValues.Add(xlRange.Cells[iRow, j].Value2);
                }
                catch(Microsoft.CSharp.RuntimeBinder.RuntimeBinderException e)
                {
                    xlWorkbookInvalidOptions.Worksheets[1].Cells[iXLWorksheetInvalidoptions, j] = null;
                    //Console.Write("");
                    arrInvalidOptionsValues.Add(null);
                }
            }
            iXLWorksheetInvalidoptions++;
            arrRowsInvalidOptionsvalues.Add(arrInvalidOptionsValues);
            //Console.Write("\n");
        }*/

        public void filterDummyPriceValues()
        {
            Excel.Range originalSheetRange = xlWorksheet.UsedRange;
            xlWorksheet.ListObjects.AddEx(Excel.XlListObjectSourceType.xlSrcRange, xlWorksheet.UsedRange, System.Type.Missing, Excel.XlYesNoGuess.xlYes).Name = "InvalidOptions";
            xlWorksheet.ListObjects["InvalidOptions"].Range.AutoFilter(21, dummyPricesStrArray, Excel.XlAutoFilterOperator.xlFilterValues);
            Excel.Range invalidOptionsRange = originalSheetRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible);
            xlWorkbook.SaveAs(@"C:\Users\U_jain\Documents\Visual Studio 2013\Projects\InvalidOptionsSheet.xlsx");
            xlWorksheetInvalidOptions = xlWorksheet;
            xlWorkbookInvalidOptions = xlWorkbook;
            xlAppInvalidOptions = xlApp;
            openOriginalWorkbook();
        }

        public void openOriginalWorkbook()
        {
            xlApp = new Excel.Application();

            xlApp.Visible = true;
            try
            {
                xlWorkbook = xlApp.Workbooks.Open(@"C:\Users\U_jain\Documents\Visual Studio 2013\Projects\IT_27thJun_Q2Wk8.xlsx");
            }
            catch (Exception e)
            {
                MessageBox.Show("Error in opening excel file!");
            }

            try
            {
                xlWorksheet = xlWorkbook.Sheets[1];
            }
            catch (Exception e)
            {
                MessageBox.Show("Expected worksheet not found!");
            }
        }
    }
    
}
