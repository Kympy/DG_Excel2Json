using System;
using System.Collections.Generic;
using System.IO;

namespace DGExcel2Json_CSharp
{
    internal class Program
    {
        public static int Main(string[] args)
        {
            if (args == null || args.Length == 0)
            {
                Console.WriteLine();
                Console.WriteLine("------ Process Start ------");
                Console.WriteLine();

                bool useLastArgument = false;

                var lastArguments = LoadLastArguments();
                if (lastArguments != null)
                {
                    Console.WriteLine("Last Argument : ");
                    for (int i = 0; i < lastArguments.Length; i++)
                    {
                        switch (i)
                        {
                            case 0:
                                Console.WriteLine($"\t[1] Excel Path : {lastArguments[i]}");
                                break;
                            case 1:
                                Console.WriteLine($"\t[2] Json Path : {lastArguments[i]}");
                                break;
                            case 2:
                                Console.WriteLine($"\t[3] Script Path : {lastArguments[i]}");
                                break;
                            case 3:
                                Console.WriteLine($"\t[4] Table Loader Path : {lastArguments[i]}");
                                break;
                        }
                    }

                    Console.WriteLine();
                    Console.WriteLine("- Continue with last arguments?");
                    Console.WriteLine("\tEnter: YES / Other: NO");
                    while (true)
                    {
                        var key = Console.ReadKey(true);
                        if (key.Key == ConsoleKey.Enter)
                        {
                            args = lastArguments;
                            useLastArgument = true;
                            break;
                        }
                        else
                        {
                            break;
                        }
                    }
                }

                if (useLastArgument == false)
                {
                    Console.WriteLine();
                    Console.WriteLine("- Enter the arguments with space.");
                    Console.WriteLine("\t[1] Excel Path / [2] Json Path / [3] Script Path / [4] Table Loader Path");
                    Console.Write(">>> ");

                    var read = Console.ReadLine();
                    args = read.Split(' ');
                }
            }
            
            Console.WriteLine();
            Console.WriteLine(" ----> Processing...");
            Console.WriteLine();

            if (args == null)
            {
                Console.WriteLine("Please enter the at least 2 arguments.");
                return (int)EDGExcel2JsonResult.EXECUTE_ARGUMENT_REQUIRED;
            }

            EDGExcel2JsonResult result = EDGExcel2JsonResult.FAILED;

            switch (args.Length)
            {
                case 1:
                {
                    result = EDGExcel2JsonResult.EXECUTE_ARGUMENT_REQUIRED;
                    break;
                }
                case 2:
                {
                    Excel2Json excel2Json = new Excel2Json();
                    result = excel2Json.CreateJsonOnly(args[0], args[1]);
                    break;
                }
                case 3:
                {
                    result = EDGExcel2JsonResult.ARGUMENT_COUNT_ERROR_3;
                    break;
                }
                case 4:
                {
                    Excel2Json excel2Json = new Excel2Json();
                    result = excel2Json.CreateAll(args[0], args[1], args[2], args[3]);
                    break;
                }
                case 0:
                {
                    result = EDGExcel2JsonResult.EXECUTE_ARGUMENT_REQUIRED;
                    break;
                }
                default:
                {
                    result = EDGExcel2JsonResult.ARGUMENT_COUNT_ERROR_MORE;
                    break;
                }
            }

            if (result == EDGExcel2JsonResult.SUCCESS) SaveLastArguments(args);
            Console.WriteLine("Program Finished.");
            Console.WriteLine("\tResult: " + result.ToString());
            return (int)result;
        }

        private static string argSaveFileName = "LastData.txt";

        private static void SaveLastArguments(string[] args)
        {
            var currentDir = Directory.GetCurrentDirectory();
            var fileDir = Path.Combine(currentDir, argSaveFileName);
            using (StreamWriter sw = new StreamWriter(fileDir))
            {
                foreach (var arg in args)
                {
                    sw.WriteLine(arg);
                }

                sw.Close();
            }
        }

        private static string[] LoadLastArguments()
        {
            var currentDir = Directory.GetCurrentDirectory();
            var fileDir = Path.Combine(currentDir, argSaveFileName);

            if (File.Exists(fileDir) == false) return null;

            List<string> argList = new List<string>();
            using (StreamReader sr = new StreamReader(fileDir))
            {
                while (sr.EndOfStream == false) argList.Add(sr.ReadLine());

                sr.Close();
            }

            return argList.ToArray();
        }
    }
}
