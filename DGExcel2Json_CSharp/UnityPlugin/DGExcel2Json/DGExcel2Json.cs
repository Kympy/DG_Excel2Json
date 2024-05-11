using System;
using System.Diagnostics;
using System.IO;
using UnityEditor;
using UnityEngine;
using Debug = UnityEngine.Debug;

namespace DGExcel2Json
{
    public class DGExcel2Json : MonoBehaviour
    {
        private static string ROOTFOLDER = "DGExcel2Json";
        private static string FILE = "DGExcel2Json_CSharp.exe";

        [MenuItem("Tools/DGExcel2Json/Generate And ReCompile", priority = 1)]
        public static void GenerateRecompile()
        {
            Generate(true);
        }

        [MenuItem("Tools/DGExcel2Json/Generate Only", priority = 2)]
        public static void GenerateNoCompile()
        {
            Generate(false);
        }

        public static void Generate(bool bRecompile = false)
        {
            string projectPath = Path.GetDirectoryName(Application.dataPath);
            string rootDir = Path.Combine(projectPath, ROOTFOLDER);
            if (Directory.Exists(rootDir) == false)
            {
                Directory.CreateDirectory(rootDir);
            }

            string fullPath = Path.Combine(rootDir, FILE);
            if (File.Exists(fullPath) == false)
            {
                Debug.LogError($"EXE file is not exist in {projectPath}. {FILE}");
                return;
            }

            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.FileName = fullPath;
            startInfo.UseShellExecute = false;
            startInfo.CreateNoWindow = true;
            startInfo.Arguments = $"{CreateExcelFolder()} {CreateJsonFolder()} {CreateScriptFolder()}";

            try
            {
                using (Process process = Process.Start(startInfo))
                {
                    process.WaitForExit();
                    int exitCode = process.ExitCode;
                    if (exitCode != 0)
                    {
                        Debug.LogError($"DG Excel2Json finished : Exit Code -> {exitCode}:{(EDGExcel2JsonResult)exitCode}");
                    }
                    else
                    {
                        Debug.Log($"DG Excel2Json finished : Exit Code -> {exitCode}:{(EDGExcel2JsonResult)exitCode}");
                    }
                    if (bRecompile)
                        ReCompile();
                }
            }
            catch (Exception e)
            {
                Debug.LogError(e.ToString());
                throw;
            }
        }

        [MenuItem("Tools/DGExcel2Json/Create excel folder")]
        public static string CreateExcelFolder()
        {
            string projectPath = Path.GetDirectoryName(Application.dataPath);
            string excelPath = Path.Combine(projectPath, "Excel");
            if (Directory.Exists(excelPath) == false)
            {
                Directory.CreateDirectory(excelPath);
            }

            return excelPath;
        }

        [MenuItem("Tools/DGExcel2Json/Create json folder")]
        public static string CreateJsonFolder()
        {
            string jsonPath = Path.Combine(Application.dataPath, "Json");
            if (Directory.Exists(jsonPath) == false)
            {
                Directory.CreateDirectory(jsonPath);
            }

            return jsonPath;
        }

        [MenuItem("Tools/DGExcel2Json/Create class script folder")]
        public static string CreateScriptFolder()
        {
            string classPath = Path.Combine(Application.dataPath, "Scripts/DataClass");
            if (Directory.Exists(classPath) == false)
            {
                Directory.CreateDirectory(classPath);
            }

            return classPath;
        }

        [MenuItem("Tools/DGExcel2Json/Recompile", priority = 3)]
        public static void ReCompile()
        {
            UnityEditor.Compilation.CompilationPipeline.RequestScriptCompilation();
        }
    }
}
