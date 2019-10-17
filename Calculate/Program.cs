using System;
using System.Text;
using System.Threading.Tasks;

namespace Calculate
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelFunction excelFunction = new ExcelFunction();

            //Console.WriteLine("请输入文科录取人数：");
            //excelFunction.wenKePersons = (int)double.Parse(Console.ReadLine());

            //Console.WriteLine("请输入理科录取人数：");
            //excelFunction.liKePersons = (int)double.Parse(Console.ReadLine());

            if (excelFunction.wenKePersons > excelFunction.liKePersons)
            {
                Console.WriteLine("文科人数大于理科人数，请检查");
                Console.Read();
                return;
            }

            excelFunction.ReadExcel();
        }
    }
}
