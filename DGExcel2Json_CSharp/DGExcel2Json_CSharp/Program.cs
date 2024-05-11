using System;

namespace DGExcel2Json_CSharp
{
    internal class Program
    {
        public static int Main(string[] args)
        {
            if (args == null)
            {
                Console.WriteLine("Please enter the 3 arguments.");
                return (int)EDGExcel2JsonResult.EXECUTE_ARGUMENT_REQUIRED;
            }

            if (args.Length < 3)
            {
                Console.WriteLine("Argument count is not enough. ExcelPath, JsonPath, ScriptPath");
                return (int)EDGExcel2JsonResult.EXECUTE_ARGUMENT_LENGTH;
            }

            Excel2Json excel2Json = new Excel2Json();
            return (int)excel2Json.CreateJson(args[0], args[1], args[2]);
        }
    }
}
