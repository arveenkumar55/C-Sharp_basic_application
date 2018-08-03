using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace Task
{
    public partial class Form1 : Form
    {
        public static List<string> data = new List<string>();
        public Form1()
        {
            InitializeComponent();
        }

        OpenFileDialog ofd = new OpenFileDialog();

        private void button1_Click(object sender, EventArgs e)
        {

            ofd.Filter = "CSV|*.csv";
           if(ofd.ShowDialog()==DialogResult.OK)
            {
                textBox1.Text = ofd.FileName;
                textBox2.Text = ofd.SafeFileName;




            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            System.IO.FileInfo fInfo = new System.IO.FileInfo(ofd.FileName);
            string strFilePath = fInfo.DirectoryName;
          //  textBox2.Text = fInfo.DirectoryName;

            var lines = File.ReadLines(ofd.FileName);

            foreach (string line in lines)
            {

                data.Add(line);
            }

            for (int i = 0; i < data.Count - 2; i++)
            {
                Console.WriteLine($"{i} = {data[i]}");
            }





            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                Console.WriteLine("Excel is not properly installed!!");
                return;
            }


            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Worksheet xlWorkSheet1;
            object misValue = System.Reflection.Missing.Value;
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            var xlSheets = xlWorkBook.Sheets as Excel.Sheets;
            var xlNewSheet = (Excel.Worksheet)xlSheets.Add(xlSheets[1], Type.Missing, Type.Missing, Type.Missing);
            var xlNewSheet2 = (Excel.Worksheet)xlSheets.Add(xlSheets[2], Type.Missing, Type.Missing, Type.Missing);
            xlWorkSheet = xlNewSheet;
            xlWorkSheet1 = xlNewSheet2;
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            xlWorkSheet1 = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(2);
            xlWorkSheet.Name = "Original file";
            xlWorkSheet1.Name = "what i want it to look like";

            xlWorkSheet1.Cells[1, 1] = "order_id";
            xlWorkSheet1.Cells[1, 2] = "date";
            xlWorkSheet1.Cells[1, 3] = "status";
            xlWorkSheet1.Cells[1, 4] = "shipping";
            xlWorkSheet1.Cells[1, 5] = "shopping_cost";
            xlWorkSheet1.Cells[1, 6] = "product_code 1";
            xlWorkSheet1.Cells[1, 7] = "product_name 1";
            xlWorkSheet1.Cells[1, 8] = "product_discount 1";
            xlWorkSheet1.Cells[1, 9] = "product_price 1";
            xlWorkSheet1.Cells[1, 10] = "product_quantity 1";
            xlWorkSheet1.Cells[1, 11] = "product_value 1";
            xlWorkSheet1.Cells[1, 12] = "sum";



      /*      xlWorkSheet1.Columns[1].ColumnWidth = 30;
            xlWorkSheet1.Columns[2].ColumnWidth = 30;
            xlWorkSheet1.Columns[3].ColumnWidth = 30;
            xlWorkSheet1.Columns[4].ColumnWidth = 30;
            xlWorkSheet1.Columns[5].ColumnWidth = 30;
            xlWorkSheet1.Columns[6].ColumnWidth = 30;
            xlWorkSheet1.Columns[7].ColumnWidth = 30;
            xlWorkSheet1.Columns[8].ColumnWidth = 30;
            xlWorkSheet1.Columns[9].ColumnWidth = 30;
            xlWorkSheet1.Columns[10].ColumnWidth = 30;
            xlWorkSheet1.Columns[11].ColumnWidth = 30;
            xlWorkSheet1.Columns[12].ColumnWidth = 30;
            */
             xlWorkSheet.Columns.AutoFit();
             xlWorkSheet1.Columns.AutoFit();


            /*      object misValue = System.Reflection.Missing.Value;

                  xlWorkBook = xlApp.Workbooks.Add(misValue);
                  xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(2);

                  xlWorkSheet.Name = "Original file";
              */

            for (int j = 0; j < data.Count; j++)
            {


                String[] splitData = data[j].Split(';');

                for (int i = 0; i < splitData.Length; i++)
                {
                    xlWorkSheet.Cells[j + 1, i + 1] = splitData[i];

                }

            }


            int row = 2;
            int col = 6;
            int counter = 0;
            for (int j = 1; j < data.Count; j++)
            {
                xlWorkSheet1.Cells[row, 1 ] = xlWorkSheet.Cells[j+1,1];
                xlWorkSheet1.Cells[row, 2] = xlWorkSheet.Cells[j + 1, 2];
                xlWorkSheet1.Cells[row, 3] = xlWorkSheet.Cells[j + 1, 3];
                xlWorkSheet1.Cells[row, 4] = xlWorkSheet.Cells[j + 1, 4];
                xlWorkSheet1.Cells[row, 5] = xlWorkSheet.Cells[j + 1, 5];
                xlWorkSheet1.Cells[row, 12] = xlWorkSheet.Cells[j + 1, 138];

                row = row + 1;
                String[] splitData = data[j].Split(';');

                for (int i = 5; i < splitData.Length-1; i++)
                {
                    

                    if(splitData[i].ToString().Equals(""))
                    {
                        break;
                    }

                    if (counter == 6)
                    {
                        row = row + 1;
                        col = 6;
                        counter = 0;
                    }

                    xlWorkSheet1.Cells[row, col] = splitData[i];
                    col = col + 1;
                    counter = counter + 1;
                }
                row = row + 1;
                col = 6;
                counter = 0;
               
            }





                xlWorkBook.SaveAs(strFilePath+"\\"+ ofd.SafeFileName+"_output.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();
            msg.Text = "Successfully generated";
        }
    }
}
