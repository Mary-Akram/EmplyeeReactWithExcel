using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;



namespace EmplyeeReactWithExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            try
            {

                //excel template
                string excelTemplateFilePath = @"C:\Users\Lenovo\Desktop\AddExcel\Employee.xlsx";
                //create Excel app instance
                Excel.Application excelApp = new Excel.Application();
                //Save Array of Byte in Hard
                byte[] myFileByte = Properties.Resources.Employee;

                System.IO.File.WriteAllBytes(excelTemplateFilePath, myFileByte);


                // open statistics template file
                excelApp.Workbooks.Open(excelTemplateFilePath);



                Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.ActiveSheet;
                EmployeeTaskEntities dbx = new EmployeeTaskEntities();
               
             var cellValues = dbx.GetAllEmployee().ToList();
                int CountOfRow = cellValues.Count();

                var row = 1;
                File.Delete(excelTemplateFilePath);


                foreach (var item in cellValues)
                {
                    row++;
                    workSheet.Cells[row, "A"].Value = item.EmpId;
                    workSheet.Cells[row, "B"].Value = item.EmpName;
                    workSheet.Cells[row, "C"].Value = item.Age;
                    workSheet.Cells[row, "D"] = item.Phone;
                    workSheet.Cells[row, "E"] = item.Department;
                }

                workSheet.SaveAs(excelTemplateFilePath);
                excelApp2.Quit();


                //seconed Try

                string excelTemplateFilePath2 = @"C:\Users\Lenovo\Desktop\AddExcel\Employee22.xlsx";
                // create Excel app instance
                Excel.Application excelApp2 = new Excel.Application();
                //Save Array of Byte in Hard
                byte[] myFileByte2 = Properties.Resources.Employee;

                System.IO.File.WriteAllBytes(excelTemplateFilePath2, myFileByte2);


                //// open statistics template file

                excelApp2.Workbooks.Open(excelTemplateFilePath2);


                Excel._Worksheet ws = (Excel.Worksheet)excelApp2.ActiveSheet;
                File.Delete(excelTemplateFilePath2);

                for (int i = 0; i < CountOfRow; i++)
                {
                  
                        //Excel.Range rng = ws.Cells[i, j] as Excel.Range;
                        // rng.Cells[i][1].Value = cellValues[i].EmpId;


                        ws.Range[ws.Cells[i + 1, 1]].Value = cellValues[i].EmpId;
                        ws.Range[ws.Cells[i + 1, 2]].Value = cellValues[i].EmpName;
                        ws.Range[ws.Cells[i + 1, 3]].Value = cellValues[i].Age;
                        ws.Range[ws.Cells[i + 1, 4]].Value = cellValues[i].Phone;
                        ws.Range[ws.Cells[i + 1, 5]].Value = cellValues[i].Department;




                }

                ws.SaveAs(excelTemplateFilePath2);

                // hide loading flag
                this.UseWaitCursor = false;

                // close excel
                excelApp2.Quit();

            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("because it is being used by another process"))
                {
                    MessageBox.Show("The file you are trying to save on is in use, please close it");
                }
                else
                {
                    throw ex;
                }
            }

        }
    }
}
