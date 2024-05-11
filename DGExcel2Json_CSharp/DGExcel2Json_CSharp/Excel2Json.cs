using System;
using System.IO;
using System.Runtime.InteropServices;
using Excel2Json_CSharp;
using Microsoft.Office.Interop.Excel;

namespace DGExcel2Json_CSharp
{
    public class Excel2Json
    {
        private Application currentApp;
        private Workbook currentWorkbook;
        private Worksheet currentSheet;

        public EDGExcel2JsonResult CreateJson(string inExcelPath, string outJsonPath, string outScriptPath)
        {
            if (string.IsNullOrEmpty(inExcelPath))
            {
                Console.WriteLine("Excel path is not valid.");
                return EDGExcel2JsonResult.EXCEL_PATH_WRONG;
            }
            
            bool pathCheck = CreateDirectoryIfNotExist(outJsonPath) && CreateDirectoryIfNotExist(outScriptPath);
            if (pathCheck == false) return EDGExcel2JsonResult.SAVE_PATH_WRONG;

            string[] excelList = GetAllExcel(inExcelPath);
            if (excelList == null) return EDGExcel2JsonResult.EXCEL_NOT_EXIST;

            foreach (var file in excelList)
            {
                currentApp = new Microsoft.Office.Interop.Excel.Application();
                currentWorkbook = currentApp.Workbooks.Open(file);
                currentSheet = currentWorkbook.Worksheets.Item[1] as Worksheet;
                EDGExcel2JsonResult result = MakeJsonFile(file, outJsonPath, outScriptPath);
                currentWorkbook.Close();
                currentApp.Quit();
                if (currentSheet != null)
                {
                    Marshal.ReleaseComObject(currentSheet);
                }

                if (currentWorkbook != null)
                {
                    Marshal.ReleaseComObject(currentWorkbook);
                }

                if (currentApp != null)
                {
                    Marshal.ReleaseComObject(currentApp);
                }

                if (result != EDGExcel2JsonResult.SUCCESS) return result;
            }

            GC.Collect();

            return EDGExcel2JsonResult.SUCCESS;
        }

        private string searchPattern = "*.xlsx";

        private string[] GetAllExcel(string inExcelPath)
        {
            if (Directory.Exists(inExcelPath) == false)
            {
                Console.WriteLine($"Cannot find the excel directory : {inExcelPath}");
                return null;
            }

            // 정해진 경로안의 파일들을 모두 가져옴
            string[] excelList = Directory.GetFiles(inExcelPath, searchPattern);
            // 파일들의 갯수가 0이면 종료
            if (excelList.Length == 0)
            {
                Console.WriteLine("There's no excel files in this directory.");
                return null;
            }

            return excelList;
        }

        private EDGExcel2JsonResult MakeJsonFile(string inExcelFile, string outJsonPath, string outScriptPath)
        {
            int startRow = -1;
            int startCol = 1;

            // A1 셀 부터 시작하여 Id 가 있는 셀을 찾는다.
            Range lastCell = (currentSheet.Cells[1, "A"] as Range).Cells;
            for (int i = 0; i < 10; i++)
            {
                // 빈 셀
                if (lastCell.Value2 == null)
                {
                    lastCell = lastCell.get_End(XlDirection.xlDown).Cells;
                    continue;
                }

                // Id 가 아닌 셀
                if (string.Compare(lastCell.Value2.ToString().ToLower(), "id") != 0)
                {
                    lastCell = (currentSheet.Cells[lastCell.Cells.Row + 1, lastCell.Cells.Column] as Range).Cells;
                    continue;
                }

                startRow = lastCell.Row;
            }

            if (startRow == -1)
            {
                Console.WriteLine($"Cannot find Id cell. File : {inExcelFile}");
                return EDGExcel2JsonResult.NO_ID_COLUMN;
            }

            int endRow = (currentSheet.Cells[startRow, "A"] as Range).get_End(XlDirection.xlDown).Row;
            int endCol = (currentSheet.Cells[startRow, "A"] as Range).get_End(XlDirection.xlToRight).Column;

            int rowCount = endRow - startRow + 1;
            int colCount = endCol;

            string[] names = new string[colCount];
            string[] types = new string[colCount];
            string[,] datas = new string[rowCount - 2, colCount];

            // 데이터 이름과 데이터 타입을 가져옴
            for (int currentColumn = startCol; currentColumn <= endCol; currentColumn++)
            {
                object columnName = (currentSheet.Cells[startRow, currentColumn] as Range).Value2;
                if (columnName == null || string.IsNullOrWhiteSpace(columnName.ToString()))
                {
                    Console.WriteLine($"Column Name Error : Cell[{startRow},{currentColumn}]");
                    return EDGExcel2JsonResult.COLUMN_NAME_ERROR;
                }

                names[currentColumn - startCol] = columnName.ToString();

                object valueType = (currentSheet.Cells[startRow + 1, currentColumn] as Range).Value2;
                if (valueType == null || string.IsNullOrWhiteSpace(valueType.ToString()))
                {
                    Console.WriteLine($"Value type Error : Cell[{startRow},{currentColumn}]");
                    return EDGExcel2JsonResult.TYPE_NAME_ERROR;
                }

                types[currentColumn - startCol] = valueType.ToString();
            }

            // 데이터 수집
            for (int curCol = startCol; curCol <= endCol; curCol++)
            {
                for (int curRow = startRow + 2; curRow <= endRow; curRow++)
                {
                    var read = (currentSheet.Cells[curRow, curCol] as Range).Value2;
                    if (read == null)
                    {
                        Console.WriteLine($"Data read error. Cell[{curRow}, {curCol}]");
                        return EDGExcel2JsonResult.DATA_READ_ERROR;
                    }

                    datas[curRow - (startRow + 2), curCol - startCol] = read?.ToString();
                }
            }

            // 파일이름 추출 : ex) 'testDocs' + '.xlsx'
            string[] fileNames = currentWorkbook.Name.Split('.');
            string fileName = fileNames[0];
            string fileFullName = Path.Combine(outJsonPath, $"{fileName}.json");
            EDGExcel2JsonResult jsonResult = WriteJson(names, types, datas, fileFullName);
            if (jsonResult != EDGExcel2JsonResult.SUCCESS) return jsonResult;

            string classFullName = Path.Combine(outScriptPath, $"{fileName}.cs");
            WriteCSharpClass(names, types, classFullName, fileName);
            return EDGExcel2JsonResult.SUCCESS;
        }

        private EDGExcel2JsonResult WriteJson(string[] columnNames, string[] valueTypes, string[,] datas, string fileFullName)
        {
            StreamWriter writer = File.CreateText(fileFullName);
            writer.WriteLine("[");
            // 행
            for (int j = 0; j < datas.GetLength(0); j++)
            {
                writer.WriteLine("\t{");
                for (int i = 0; i < columnNames.Length; i++)
                {
                    writer.Write($"\t\t\"{columnNames[i]}\": ");
                    EDataType dataType = GetDataType(valueTypes[i]);
                    switch (dataType)
                    {
                        case EDataType.Int:
                        case EDataType.Bool:
                        case EDataType.Float:
                            writer.Write(datas[j, i]);
                            break;
                        case EDataType.String:
                            writer.Write($"\"{datas[j, i]}\"");
                            break;
                        case EDataType.Vector3:
                            writer.Write(datas[j, i]);
                            break;
                        case EDataType.IntArray:
                        case EDataType.FloatArray:
                            writer.Write("[");
                            writer.Write(datas[j, i]);
                            writer.Write("]");
                            break;
                        case EDataType.NOT_DEFINED:
                            writer.Close();
                            Console.WriteLine($"NOT DEFINED data type. {valueTypes[i]}");
                            return EDGExcel2JsonResult.DATA_TYPE_NOT_DEFINED;
                    }

                    // 마지막 열이거나 배열이면 , 생략
                    if (i == columnNames.Length - 1 || IsArrayType(dataType))
                    {
                        writer.WriteLine();
                    }
                    else // 마지막 열이 아니고 배열이 아닐 때
                    {
                        writer.WriteLine(",");
                    }
                }

                // 마지막 행인지?
                if (j == datas.GetLength(0) - 1)
                {
                    writer.WriteLine("\t}");
                }
                else
                {
                    writer.WriteLine("\t},");
                }
            }

            writer.Write("]");
            writer.Close();
            return EDGExcel2JsonResult.SUCCESS;
        }

        private EDataType GetDataType(string data)
        {
            switch (data.ToLower())
            {
                case "int": return EDataType.Int;
                case "bool": return EDataType.Bool;
                case "float": return EDataType.Float;
                case "string": return EDataType.String;
                case "vector3": return EDataType.Vector3;
                case "int[]": return EDataType.IntArray;
                case "float[]": return EDataType.FloatArray;
                default: return EDataType.NOT_DEFINED;
            }
        }

        private bool IsArrayType(EDataType dataType) { return dataType == EDataType.IntArray || dataType == EDataType.FloatArray || dataType == EDataType.BoolArray || dataType == EDataType.StringArray; }

        private void WriteCSharpClass(string[] columnNames, string[] valueTypes, string outScriptPath, string className)
        {
            StreamWriter writer = File.CreateText(outScriptPath);
            writer.WriteLine("// Auto Created by DG Excel2Json.");
            writer.WriteLine();
            writer.WriteLine($"public class {className}");
            writer.WriteLine("{");
            for (int i = 0; i < columnNames.Length; i++)
            {
                writer.WriteLine($"\tpublic {valueTypes[i]} {columnNames[i]};");
            }

            writer.Write("}");
            writer.Close();
        }

        private bool CreateDirectoryIfNotExist(string path)
        {
            if (string.IsNullOrEmpty(path))
            {
                Console.WriteLine($"Path is null or empty : {path}");
                return false;
            }
            try
            {
                Directory.CreateDirectory(path);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                return false;
            }
            return true;
        }
    }
}
