using Excel = Microsoft.Office.Interop.Excel;
using System.Data;

namespace MainApp
{
    static class ExcelExport
    {
        public static void GenerateExcel(DataTable dataTable, string path)
        {
            DataSet dataSet = new DataSet();
            dataSet.Tables.Add(dataTable);

            // create a excel app along side with workbook and worksheet and give a name to it
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWorkBook = excelApp.Workbooks.Add();
            Excel._Worksheet xlWorksheet = excelWorkBook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            foreach (DataTable table in dataSet.Tables)
            {
                //Add a new worksheet to workbook with the Datatable name
                Excel.Worksheet excelWorkSheet = excelWorkBook.Sheets.Add();
                excelWorkSheet.Name = table.TableName;

                // add all the columns
                for (int i = 1; i < table.Columns.Count + 1; i++)
                {
                    excelWorkSheet.Cells[1, i] = table.Columns[i - 1].ColumnName;
                }

                // add all the rows
                for (int j = 0; j < table.Rows.Count; j++)
                {
                    for (int k = 0; k < table.Columns.Count; k++)
                    {
                        excelWorkSheet.Cells[j + 2, k + 1] = table.Rows[j].ItemArray[k].ToString();
                    }
                }
            }
            excelWorkBook.SaveAs(path);
            excelWorkBook.Close();
            excelApp.Quit();
        }
    }
}
