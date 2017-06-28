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

namespace ReadExcelSheet
{
    public partial class Form1 : Form
    {
        Excel.Application xlApp;
        Excel.Workbook xlWorkbook;
        Excel._Worksheet xlWorksheet;
        Excel.Range xlRange;
        int rowCount;
        int colCount;
        Array dummyPricesArray;
        string[] dummyPricesStrArray;
        public Form1()
        {
            InitializeComponent();
            InitializeObjects();
            createInvalidOptionsExcel();
            //readExcelDump();
        }

        public void createInvalidOptionsExcel()
        {
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                Console.WriteLine("EXCEL could not be started. Check that your office installation and project references are correct.");
                return;
            }
            xlApp.Visible = true;

            Excel.Workbook wb = xlApp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
            Excel._Worksheet ws = wb.Worksheets[1];

            if (ws == null)
            {
                Console.WriteLine("Worksheet could not be created. Check that your office installation and project references are correct.");
            }

            // Select the Excel cells, in the range c1 to c7 in the worksheet.
            Excel.Range aRange = ws.get_Range("C1", "C7");

            if (aRange == null)
            {
                Console.WriteLine("Could not get a range. Check to be sure you have the correct versions of the office DLLs.");
            }

            // Fill the cells in the C1 to C7 range of the worksheet with the number 6.
            Object[] args = new Object[1];
            args[0] = 6;
            //aRange.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, aRange, args);
        }

        public void InitializeObjects()
        {
            xlApp = new Excel.Application();
            
            try
            {
                xlWorkbook = xlApp.Workbooks.Open(@"C:\Users\U_jain\MyData\LanguageTranslationToolAutomation\IT_27thJun_Q2Wk8.xlsx");
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

            dummyPricesStrArray = new string[]{ "999999.99", "9999999.99" };
            dummyPricesArray = dummyPricesStrArray;
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
                    //Console.Write("BLANK VALUE : " + strCellValue + "\n");
                }

                if (Convert.ToDouble(strCellValue) <= 0)
                {
                    // Negative price values (CODE FOR INVALID OPTION)
                    writeInvalidOptions(i);
                    //Console.Write("NEGATIVE VALUE : " + strCellValue + "\n");
                }
                else
                {
                    for (int j = 0; j < dummyPricesArray.Length; j++)
                    {
                        if (strCellValue == dummyPricesStrArray[j])
                        {
                            // Dummy price values (CODE FOR INVALID OPTION)
                            writeInvalidOptions(i);
                            //Console.Write("DUMMY VALUE : " + strCellValue + "\n");
                        }
                    }
                }
                
            }
        }

        public void writeInvalidOptions(int iRow)
        {
            for (int j = 1; j <= colCount; j++ )
            {
                try
                {
                    Console.Write(xlRange.Cells[iRow, j].Value2.ToString());
                }
                catch(Microsoft.CSharp.RuntimeBinder.RuntimeBinderException e)
                {
                    Console.Write("");
                }
                
            }
            Console.Write("\n");
        }
    }
}
