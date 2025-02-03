using System;
using System.Collections.Generic;
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

        public EDGExcel2JsonResult CreateAll(string inExcelPath, string outJsonPath, string outScriptPath, string loaderPath)
        {
            if (string.IsNullOrEmpty(inExcelPath))
            {
                Console.WriteLine("Excel path is not valid.");
                return EDGExcel2JsonResult.EXCEL_PATH_WRONG;
            }
            
            bool pathCheck = CreateDirectoryIfNotExist(outJsonPath) && CreateDirectoryIfNotExist(outScriptPath) && CreateDirectoryIfNotExist(loaderPath);
            if (pathCheck == false) return EDGExcel2JsonResult.SAVE_PATH_WRONG;

            string[] excelList = GetAllExcel(inExcelPath);
            if (excelList == null) return EDGExcel2JsonResult.EXCEL_NOT_EXIST;

            List<string> fileNames = new List<string>();
            fileNames.Clear();

            Application activeExcel = null;
            try
            {
                activeExcel = (Application)Marshal.GetActiveObject("Excel.Application");
            }
            catch (Exception e)
            {
                activeExcel = null;
            }
            if (activeExcel != null)
            {
                return EDGExcel2JsonResult.EXCEL_IS_RUNNING;
            }
            
            foreach (var file in excelList)
            {
                currentApp = new Microsoft.Office.Interop.Excel.Application();
                currentWorkbook = currentApp.Workbooks.Open(file);
                var sheets = currentWorkbook.Worksheets;
                currentSheet = sheets.Item[1] as Worksheet;
                EDGExcel2JsonResult result = MakeJsonFile(file, outJsonPath, outScriptPath);
                currentWorkbook.Close();
                currentApp.Quit();
                if (currentSheet != null)
                {
                    Marshal.FinalReleaseComObject(currentSheet);
                }
                if (sheets != null)
                {
                    Marshal.FinalReleaseComObject(sheets);
                }

                if (currentWorkbook != null)
                {
                    Marshal.FinalReleaseComObject(currentWorkbook);
                }

                if (currentApp != null)
                {
                    Marshal.FinalReleaseComObject(currentApp);
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();

                if (result != EDGExcel2JsonResult.SUCCESS)
                {
                    return result;
                }

                fileNames.Add(Path.GetFileNameWithoutExtension(file));
            }

            EDGExcel2JsonResult loaderResult = WriteTableLoader(fileNames.ToArray(), loaderPath);
            if (loaderResult != EDGExcel2JsonResult.SUCCESS)
            {
                GC.Collect();
                return loaderResult;
            }

            GC.Collect();
            return EDGExcel2JsonResult.SUCCESS;
        }

        public EDGExcel2JsonResult CreateJsonOnly(string inExcelPath, string outJsonPath)
        {
            if (string.IsNullOrEmpty(inExcelPath))
            {
                Console.WriteLine("Excel path is not valid.");
                return EDGExcel2JsonResult.EXCEL_PATH_WRONG;
            }
            
            bool pathCheck = CreateDirectoryIfNotExist(outJsonPath);
            if (pathCheck == false) return EDGExcel2JsonResult.SAVE_PATH_WRONG;

            string[] excelList = GetAllExcel(inExcelPath);
            if (excelList == null) return EDGExcel2JsonResult.EXCEL_NOT_EXIST;

            List<string> fileNames = new List<string>();
            fileNames.Clear();

            Application activeExcel = null;
            try
            {
                activeExcel = (Application)Marshal.GetActiveObject("Excel.Application");
            }
            catch (Exception e)
            {
                activeExcel = null;
            }
            if (activeExcel != null)
            {
                return EDGExcel2JsonResult.EXCEL_IS_RUNNING;
            }
            
            foreach (var file in excelList)
            {
                currentApp = new Microsoft.Office.Interop.Excel.Application();
                currentWorkbook = currentApp.Workbooks.Open(file);
                var sheets = currentWorkbook.Worksheets;
                currentSheet = sheets.Item[1] as Worksheet;
                EDGExcel2JsonResult result = MakeJsonFile(file, outJsonPath, null);

                currentWorkbook.Close(false);
                currentApp.Quit();
                if (currentSheet != null)
                {
                    Marshal.FinalReleaseComObject(currentSheet);
                }
                if (sheets != null)
                {
                    Marshal.FinalReleaseComObject(sheets);
                }

                if (currentWorkbook != null)
                {
                    Marshal.FinalReleaseComObject(currentWorkbook);
                }

                if (currentApp != null)
                {
                    Marshal.FinalReleaseComObject(currentApp);
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();

                if (result != EDGExcel2JsonResult.SUCCESS)
                {
                    return result;
                }
                
                fileNames.Add(Path.GetFileNameWithoutExtension(file));
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
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

        private const string color = "Color";
        private const string colorArr = "Color[]";
        private const string colorLower = "color";
        private const string colorArrLower = "color[]";
        private EDGExcel2JsonResult MakeJsonFile(string inExcelFile, string outJsonPath, string outScriptPath = null)
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

                names[currentColumn - startCol] = MakeUpperFirstCharacter(columnName.ToString());

                object valueType = (currentSheet.Cells[startRow + 1, currentColumn] as Range).Value2;
                if (valueType == null || string.IsNullOrWhiteSpace(valueType.ToString()))
                {
                    Console.WriteLine($"Value type Error : Cell[{startRow},{currentColumn}]");
                    return EDGExcel2JsonResult.TYPE_NAME_ERROR;
                }

                string typeString = valueType.ToString();
                switch (typeString.ToLower())
                {
                    default:
                    {
                        types[currentColumn - startCol] = valueType.ToString();
                        break;
                    }
                }
            }

            // 데이터 수집
            for (int curCol = startCol; curCol <= endCol; curCol++)
            {
                if (names[curCol - startCol].Contains(IgnoreColumn)) continue;
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

            if (string.IsNullOrEmpty(outScriptPath) == false)
            {
                string classFullName = Path.Combine(outScriptPath, $"{fileName}Row.cs");
                WriteCSharpClass(names, types, classFullName, fileName);
            }
            return EDGExcel2JsonResult.SUCCESS;
        }

        private static string IgnoreColumn = "#";
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
                    if (columnNames[i].Contains(IgnoreColumn)) continue;
                    writer.Write($"\t\t\"{columnNames[i]}\": ");
                    EDataType dataType = GetDataType(valueTypes[i]);
                    switch (dataType)
                    {
                        case EDataType.Bool:
                            {
                                var lower = datas[j, i].ToString().ToLower();
                                writer.Write(lower);
                                break;
                            }
                        case EDataType.Int:
                        case EDataType.Float:
                        case EDataType.Long:
                            {
                                writer.Write(datas[j, i]);
                                break;
                            }
                        case EDataType.String:
                            {
                                writer.Write($"\"{datas[j, i]}\"");
                                break;
                            }
                        case EDataType.Vector3:
                        case EDataType.Color:
                        case EDataType.IntArray:
                        case EDataType.FloatArray:
                        case EDataType.ColorArray:
                        case EDataType.StringArray:
                            {
                                writer.Write("[");
                                writer.Write(datas[j, i]);
                                writer.Write("]");
                                break;
                            }
                        case EDataType.NOT_DEFINED:
                            {
                                writer.Close();
                                Console.WriteLine($"NOT DEFINED data type. {valueTypes[i]}");
                                return EDGExcel2JsonResult.DATA_TYPE_NOT_DEFINED;
                            }
                    }

                    // 마지막 열임
                    if (i == columnNames.Length - 1)
                    {
                        writer.WriteLine();
                    }
                    else // 마지막 열이 아님
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
                case "long": return EDataType.Long;
                case "string": return EDataType.String;
                case "vector3": return EDataType.Vector3;
                case "int[]": return EDataType.IntArray;
                case "float[]": return EDataType.FloatArray;
                case "color" : return EDataType.Color;
                case "color[]" : return EDataType.ColorArray;
                case "string[]": return EDataType.StringArray;
                
                default: return EDataType.NOT_DEFINED;
            }
        }

        private bool IsArrayType(EDataType dataType) { return dataType == EDataType.IntArray || dataType == EDataType.FloatArray || dataType == EDataType.BoolArray || dataType == EDataType.StringArray; }

        private void WriteCSharpClass(string[] columnNames, string[] valueTypes, string outScriptPath, string className)
        {
            string rowName = $"{className}Row";
            StreamWriter writer = File.CreateText(outScriptPath);
            writer.WriteLine("// Auto Created by DG Excel2Json.");
            writer.WriteLine();
            writer.WriteLine("using UnityEngine;");
            writer.WriteLine();
            writer.WriteLine("[System.Serializable]");
            writer.WriteLine($"public class {rowName} : DGTableData");
            writer.WriteLine("{");
            for (int i = 0; i < columnNames.Length; i++)
            {
                if (i == 0) continue; // Id 스킵
                if (columnNames[i].Contains(IgnoreColumn)) continue;

                bool isColor = string.Compare(valueTypes[i].ToLower(), colorLower) == 0;
                bool isColorArr = string.Compare(valueTypes[i].ToLower(), colorArrLower) == 0; 
                if (isColor)
                {   
                    writer.WriteLine($"\tpublic float[] {columnNames[i]};");
                    writer.WriteLine($"\tpublic {color} Get{columnNames[i]} {{ get {{ return new Color({columnNames[i]}[0], {columnNames[i]}[1], {columnNames[i]}[2]); }} }}");
                }
                else if (isColorArr)
                {
                    writer.WriteLine($"\tpublic float[][] {columnNames[i]};");
                    writer.WriteLine($"\tpublic {colorArr} Get{columnNames[i]}(int index)\n");
                    writer.WriteLine($"\t{{");
                    writer.WriteLine($"\t\treturn new Color({columnNames[i]}[0][index], {columnNames[i]}[1][index], {columnNames[i]}[2][index]);");
                    writer.WriteLine($"\t}}");
                }
                else
                {
                    writer.WriteLine($"\tpublic {valueTypes[i]} {columnNames[i]};");
                }
            }
            writer.WriteLine();
            writer.WriteLine($"\tpublic static readonly {className}Table Table = new {className}Table();");
            writer.Write("}");
            
            writer.WriteLine();
            writer.WriteLine($"public class {className}Table : DGTable<{rowName}> {{ }}");
            
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

        private static EDGExcel2JsonResult WriteTableLoader(string[] tableNames, string outLoaderPath)
        {
            var loaderName = Path.Combine(outLoaderPath, "DGTableLoader.cs");
            StreamWriter writer = null;
            try
            {
                writer = File.CreateText(loaderName);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                return EDGExcel2JsonResult.FILE_WRITE_ACCESS_DENIED;
            }
            writer.WriteLine("// Auto Created by DG Excel2Json.");
            writer.WriteLine();
            writer.WriteLine($"public class DGTableLoader");
            writer.WriteLine("{");
            writer.WriteLine("\tpublic string JsonLoadPath = \"Assets/Json\";");
            writer.WriteLine("\tpublic void LoadAll()");
            writer.WriteLine("\t{");
            for (int i = 0; i < tableNames.Length; i++)
            {
                writer.WriteLine($"\t\t{tableNames[i]}Row.Table.Load(JsonLoadPath);");
            }
            writer.WriteLine("\t}");
            writer.WriteLine("}");
            writer.Close();

            return EDGExcel2JsonResult.SUCCESS;
        }

        private string MakeUpperFirstCharacter(string text)
        {
            if (string.IsNullOrEmpty(text)) return null;
            if (text.Length == 1)
            {
                return "{char.ToUpper(text[0])}";
            }
            return $"{char.ToUpper(text[0])}{text.Substring(1)}";
        }
    }
}
