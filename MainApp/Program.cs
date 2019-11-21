using MainApp.Model;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace MainApp
{
    class Program
    {
        static void Main(string[] args)
        {
            var userList = UserManager.SeedData();
            try
            {
                Console.WriteLine("Please select a operation to do ..");
                Console.WriteLine("1. Export Data as EXCEL");
                int command = Convert.ToInt32(Console.ReadLine());
                switch (command)
                {
                    case 1:
                        string fileName = "UserManager.xlsx";
                        Console.WriteLine("Please give a location to save :");
                        string location = Console.ReadLine();
                        string customExcelSavingPath = location + "\\" + fileName;
                        ExcelExport.GenerateExcel(ConvertToDataTable(userList), customExcelSavingPath);
                        break;
                    default:
                        break;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


        // T : Generic Class
        static DataTable ConvertToDataTable<T>(List<T> models)
        {
            DataTable dataTable = new DataTable(typeof(T).Name);

            //Get all the properties
            PropertyInfo[] Props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);

            // Loop through all the properties            
            // Adding Column to our datatable
            foreach (PropertyInfo prop in Props)
            {
                //Setting column names as Property names  
                dataTable.Columns.Add(prop.Name);
            }
            // Adding Row
            foreach (T item in models)
            {
                var values = new object[Props.Length];
                for (int i = 0; i < Props.Length; i++)
                {
                    //inserting property values to datatable rows  
                    values[i] = Props[i].GetValue(item, null);
                }
                // Finally add value to datatable  
                dataTable.Rows.Add(values);
            }
            return dataTable;
        }
    }
}
