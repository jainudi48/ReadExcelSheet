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
            readExcelDump();
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

            dummyPricesStrArray = new string[]{ "999999.99", "9999999.99" };
            dummyPricesArray = dummyPricesStrArray;
        }
        public void readExcelDump()
        {
            
            rowCount = xlWorksheet.UsedRange.Rows.Count;
            colCount = xlWorksheet.UsedRange.Columns.Count;

            for (int i = 1; i <= rowCount; i++)
            {
                for (int j = 1; j <= colCount; j++)
                {
                    //new line
                    if (j == 1)
                        Console.Write("\r\n");

                    //write the value to the console
                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                        Console.Write(xlRange.Cells[i, j].Value2.ToString() + "\t");
                        
                    //add useful things here!   
                    
                }
            }

            /*for(int i = 2; i <= rowCount; i++)
            {
                for(int j = 0; j < dummyPricesArray.Length; j++)
                {
                    if((string)(xlWorksheet.Cells[i, colCount] as Excel.Range).Value == dummyPricesStrArray[j])
                    {
                        Console.Write(xlRange.Cells[i, colCount].Value2.ToString() + "\t");
                    }
                }
                
            }*/
        }
    }
}
