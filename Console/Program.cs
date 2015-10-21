using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Console
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelWorkBookReader reader = new ExcelWorkBookReader();

            reader.ReadHeader();
            //List<Data> data = reader.ReadExcelWorkBook();

            //reader.WriterData(data);

            System.Console.ReadLine();

        }
    }
}
