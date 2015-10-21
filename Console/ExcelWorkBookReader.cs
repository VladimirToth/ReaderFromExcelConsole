using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Console
{
    class ExcelWorkBookReader
    {

       public void ReadHeader()
        {
            List<string> nameOfColumns = new List<string>();

            string filename = @"C:\Users\toth\Desktop\Kópia.xlsx";

            string con = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filename + ";Extended Properties='Excel 12.0;HDR=NO;'";
            using (OleDbConnection connection = new OleDbConnection(con))
            {
                connection.Open();

                OleDbCommand command = new OleDbCommand("select * from [e-Order$]", connection);
                using (OleDbDataReader dr = command.ExecuteReader())
                {
                    

                    for (int i = 0; i < 1; i++)
                    {
                        while (dr.Read())
                        {
                            if (dr[0].ToString() != "")
                            {
                                nameOfColumns.Add(dr[8].ToString());
                            }
                        }
                    }
                    connection.Close();
                }

            }

            foreach (var item in nameOfColumns)
            {
                System.Console.WriteLine(item);
            }

        }

        public List<Data> ReadExcelWorkBook()
        {
            string filename = @"C:\Users\toth\Desktop\Kópia.xlsx";

            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook excelBook = xlApp.Workbooks.Open(filename);

            string[] excelSheets = new string[excelBook.Worksheets.Count];
            int i = 0;
            foreach (Microsoft.Office.Interop.Excel.Worksheet wSheet in excelBook.Worksheets)
            {
                excelSheets[i] = wSheet.Name;
                i++;
            }

            string sheet = excelSheets[0] + "$";

            excelBook.Close();

            ExcelWorkBook workBook = new ExcelWorkBook();

            List<Data> list = new List<Data>();

    

            string con = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filename + ";Extended Properties='Excel 12.0;HDR=YES;'";
            using (OleDbConnection connection = new OleDbConnection(con))
            {
                connection.Open();

                OleDbCommand command = new OleDbCommand("select * from [" + sheet+ "]", connection);
                using (OleDbDataReader dr = command.ExecuteReader())
                    while (dr.Read())
                    {
                        if (dr[0].ToString() != "")
                        {
                            list.Add(new Data
                            {
                                oldCode = dr[0].ToString(),
                                newCode = Convert.ToInt32(dr[1]),
                                good = dr[2].ToString(),
                                mj = dr[3].ToString(),
                                quantity = Convert.ToInt32(dr[4]),
                                oldPrice = Convert.ToDouble(dr[5]),
                                newPrice = Convert.ToDouble(dr[7]),
                                EAN = Convert.ToInt64(dr[8])
                            }
                                 );
                        }
                    }

                connection.Close();
            }
            return list;

        }

        public void WriterData(IEnumerable<Data> data)
        {
            foreach (var item in data)
            {
                System.Console.WriteLine(item);
            }

        }
    }
}
