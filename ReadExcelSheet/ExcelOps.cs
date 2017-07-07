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
    public partial class ExcelOps : Form
    {
        Excel.Application xlApp;
        Excel.Workbook xlWorkbook;
        Excel._Worksheet xlWorksheet;
        Excel.Application xlAppInvalidOptions;
        Excel.Workbook xlWorkbookInvalidOptions;
        Excel._Worksheet xlWorksheetInvalidOptions;
        Excel.Range xlRange;
        Excel.Range xlRangeInvalidOptions;
        int rowCount;
        int colCount;
        Array dummyPricesArray;
        string[] dummyPricesStrArray;

        public ExcelOps()
        {
            InitializeComponent();
            InitializeObjects();
            //xlWorkbookInvalidOptions = createInvalidOptionsExcel();
            //readExcelDump();
            filterDefaultStatusFalse();
            
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

        public void filterDefaultStatusFalse()
        {
            Excel.Range originalSheetRange = xlWorksheet.UsedRange;
            xlWorksheet.ListObjects.AddEx(Excel.XlListObjectSourceType.xlSrcRange, xlWorksheet.UsedRange, System.Type.Missing, Excel.XlYesNoGuess.xlYes).Name = "DefaultStatus";
            xlWorksheet.ListObjects["DefaultStatus"].Range.AutoFilter(13, new string[]{"FALSE"}, Excel.XlAutoFilterOperator.xlFilterValues);

            xlRange = xlWorksheet.UsedRange;
            xlRange.Copy(Type.Missing);


            xlAppInvalidOptions = new Excel.Application();
            xlAppInvalidOptions.Visible = true;
            try
            {
                xlWorkbookInvalidOptions = xlAppInvalidOptions.Workbooks.Add(); //Open( @"C:\Users\U_jain\Documents\Visual Studio 2013\Projects\InvalidOptionsSheet.xlsx", false);
            }
            catch (Exception e)
            {
                MessageBox.Show("Error in opening excel file!");
            }

            try
            {
                xlWorksheetInvalidOptions = xlWorkbookInvalidOptions.Sheets[1];
                //xlWorkbookInvalidOptions.SaveAs(@"C:\Users\U_jain\Documents\Visual Studio 2013\Projects\InvalidOptionsSheet.xlsx");
            }
            catch (Exception e)
            {
                MessageBox.Show("Expected worksheet not found!");
            }

            xlRangeInvalidOptions = xlWorksheetInvalidOptions.Cells[1,1];

            xlRangeInvalidOptions.Select();
            xlWorksheetInvalidOptions.Paste(Type.Missing, Type.Missing);

            xlWorkbookInvalidOptions.SaveAs(@"C:\Users\U_jain\Documents\Visual Studio 2013\Projects\InvalidOptionsSheet.xlsx");
            xlWorkbookInvalidOptions.Close(0);
            xlAppInvalidOptions.Quit();
            xlWorkbook.Close(0);
            xlApp.Quit();
            filterSalesMediaCodeB();
        }

        public void filterSalesMediaCodeB()
        {
            xlAppInvalidOptions = new Excel.Application();
            xlAppInvalidOptions.Visible = true;

            xlApp = new Excel.Application();
            xlApp.Visible = true;
            try
            {
                xlWorkbook = xlApp.Workbooks.Open(@"C:\Users\U_jain\Documents\Visual Studio 2013\Projects\InvalidOptionsSheet.xlsx");
                xlWorkbookInvalidOptions = xlAppInvalidOptions.Workbooks.Add(); // Open(@"C:\Users\U_jain\Documents\Visual Studio 2013\Projects\InvalidOptionsSheet.xlsx");
            }
            catch (Exception e)
            {
                MessageBox.Show("Error in opening excel file!");
            }

            try
            {
                xlWorksheet = xlWorkbook.Sheets[1];
                xlWorksheetInvalidOptions = xlWorkbookInvalidOptions.Sheets[1];
            }
            catch (Exception e)
            {
                MessageBox.Show("Expected worksheet not found!");
            }

            xlRangeInvalidOptions = xlWorksheetInvalidOptions.UsedRange;
            xlRange = xlWorksheet.UsedRange;
            
            Excel.Range originalSheetRange = xlWorksheet.UsedRange;
            xlWorksheet.ListObjects.AddEx(Excel.XlListObjectSourceType.xlSrcRange, xlWorksheet.UsedRange, System.Type.Missing, Excel.XlYesNoGuess.xlYes).Name = "SalesMediaCode";
            xlWorksheet.ListObjects["SalesMediaCode"].Range.AutoFilter(19, new string[] { "B" }, Excel.XlAutoFilterOperator.xlFilterValues);
            Excel.Range salesMediaCodeBRange = originalSheetRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible);

            xlRange = xlWorksheet.UsedRange;
            xlRange.Copy(Type.Missing);
            xlRangeInvalidOptions = xlWorksheetInvalidOptions.Cells[1, 1];
            xlRangeInvalidOptions.Select();
            xlWorksheetInvalidOptions.Paste(Type.Missing, Type.Missing);

            xlWorksheetInvalidOptions.SaveAs(@"C:\Users\U_jain\Documents\Visual Studio 2013\Projects\InvalidOptionsSheetSalesMediaCodeB.xlsx");
            xlWorkbookInvalidOptions.Close(0);
            xlAppInvalidOptions.Quit();
            xlWorkbook.Close(0);
            xlApp.Quit();
            filterDummyPriceValues();
        }

        public void filterDummyPriceValues()
        {
            xlAppInvalidOptions = new Excel.Application();
            xlAppInvalidOptions.Visible = true;

            xlApp = new Excel.Application();
            xlApp.Visible = true;
            try
            {
                xlWorkbook = xlApp.Workbooks.Open(@"C:\Users\U_jain\Documents\Visual Studio 2013\Projects\InvalidOptionsSheet.xlsx");
                xlWorkbookInvalidOptions = xlAppInvalidOptions.Workbooks.Add(); // Open(@"C:\Users\U_jain\Documents\Visual Studio 2013\Projects\InvalidOptionsSheet.xlsx");
            }
            catch (Exception e)
            {
                MessageBox.Show("Error in opening excel file!");
            }

            try
            {
                xlWorksheet = xlWorkbook.Sheets[1];
                xlWorksheetInvalidOptions = xlWorkbookInvalidOptions.Sheets[1];
            }
            catch (Exception e)
            {
                MessageBox.Show("Expected worksheet not found!");
            }

            xlRangeInvalidOptions = xlWorksheetInvalidOptions.UsedRange;
            xlRange = xlWorksheet.UsedRange;

            
            Excel.Range originalSheetRange = xlWorksheet.UsedRange;
            xlWorksheet.ListObjects.AddEx(Excel.XlListObjectSourceType.xlSrcRange, xlWorksheet.UsedRange, System.Type.Missing, Excel.XlYesNoGuess.xlYes).Name = "OptionPrices";
            xlWorksheet.ListObjects["OptionPrices"].Range.AutoFilter(21, dummyPricesStrArray, Excel.XlAutoFilterOperator.xlFilterValues);
            Excel.Range invalidOptionsRange = originalSheetRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible);

            xlRange = xlWorksheet.UsedRange;
            xlRange.Copy(Type.Missing);
            xlRangeInvalidOptions = xlWorksheetInvalidOptions.Cells[1, 1];
            xlRangeInvalidOptions.Select();
            xlWorksheetInvalidOptions.Paste(Type.Missing, Type.Missing);

            xlWorksheetInvalidOptions.SaveAs(@"C:\Users\U_jain\Documents\Visual Studio 2013\Projects\InvalidOptionsDummyPriceValues.xlsx");
            xlWorkbookInvalidOptions.Close(0);
            xlAppInvalidOptions.Quit();
            xlWorkbook.Close(0);
            xlApp.Quit();
            
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
