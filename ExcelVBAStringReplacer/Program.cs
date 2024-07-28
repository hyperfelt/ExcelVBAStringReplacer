using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Microsoft.Vbe.Interop;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace ExcelVBAStringReplacer
{
    internal class Program
    {
        private const string DIRECTORY_ARGUMENT_FULL = "--directory";
        private const string DIRECTORY_ARGUMENT_SHORT = "-d";

        private const string STRING_TO_SEARCH_ARGUMENT_FULL = "--string-to-search";
        private const string STRING_TO_SEARCH_ARGUMENT_SHORT = "-s";

        private const string STRING_TO_REPLACE_WITH_ARGUMENT_FULL = "--replace-with";
        private const string STRING_TO_REPLACE_WITH_ARGUMENT_SHORT = "-r";

        private const string FILES_TO_PROCESS_LIST_FILENAME = "excelVbaStringReplacer_filesToProcess.txt";
        private const string FAILED_FILES_LIST_FILENAME = "excelVbaStringReplacer_failedFiles.txt";

        static void Main(string[] args)
        {
            // Initialize statistic variables
            int nbFilesToEdit = 0;
            int editedFilesCount = 0;

            // Start a new stopwatch to measure execution time
            Stopwatch sw = Stopwatch.StartNew();
            Console.WriteLine("Starting application...");

            try
            {
                // Check if all three CLI arguments are provided and valid
                if (args.Length != 3)
                    throw new ArgumentException("Please provide the following arguments to start the application:" +
                                                $"\n{DIRECTORY_ARGUMENT_SHORT}|{DIRECTORY_ARGUMENT_FULL}:\"C:\\path\\to\\the\\folder\"" +
                                                $"\n{STRING_TO_SEARCH_ARGUMENT_SHORT}|{STRING_TO_SEARCH_ARGUMENT_FULL}:\"string to replace\"" +
                                                $"\n{STRING_TO_REPLACE_WITH_ARGUMENT_SHORT}|{STRING_TO_REPLACE_WITH_ARGUMENT_FULL}:\"new string\"",
                        nameof(args));

                string directory = args.FirstOrDefault(x =>
                    x.StartsWith(DIRECTORY_ARGUMENT_SHORT) || x.StartsWith(DIRECTORY_ARGUMENT_FULL));

                string stringToSearch = args.FirstOrDefault(x =>
                    x.StartsWith(STRING_TO_SEARCH_ARGUMENT_SHORT) || x.StartsWith(STRING_TO_SEARCH_ARGUMENT_FULL));

                string stringToReplaceWith = args.FirstOrDefault(x =>
                    x.StartsWith(STRING_TO_REPLACE_WITH_ARGUMENT_SHORT) ||
                    x.StartsWith(STRING_TO_REPLACE_WITH_ARGUMENT_FULL));

                if (string.IsNullOrWhiteSpace(directory))
                    throw new ArgumentException(
                        $"Please check if the argument {DIRECTORY_ARGUMENT_SHORT}|{DIRECTORY_ARGUMENT_FULL} is provided with a value in your command." +
                        $"\nExample : {DIRECTORY_ARGUMENT_FULL}:\"C:\\path\\to\\the\\folder\".");

                directory = directory.Replace(DIRECTORY_ARGUMENT_SHORT + ":", string.Empty);
                directory = directory.Replace(DIRECTORY_ARGUMENT_FULL + ":", string.Empty);
                directory = directory.Replace("\"", string.Empty);

                if (string.IsNullOrWhiteSpace(stringToSearch))
                    throw new ArgumentException(
                        $"Please check if the argument {STRING_TO_SEARCH_ARGUMENT_SHORT}|{STRING_TO_SEARCH_ARGUMENT_FULL} is provided with a value in your command." +
                        $"\nExample : {STRING_TO_SEARCH_ARGUMENT_FULL}:\"string to replace\".");

                stringToSearch = stringToSearch.Replace(STRING_TO_SEARCH_ARGUMENT_SHORT + ":", string.Empty);
                stringToSearch = stringToSearch.Replace(STRING_TO_SEARCH_ARGUMENT_FULL + ":", string.Empty);
                stringToSearch = stringToSearch.Replace("\"", string.Empty);

                if (string.IsNullOrWhiteSpace(stringToReplaceWith))
                    throw new ArgumentException(
                        $"Please check if the argument {STRING_TO_REPLACE_WITH_ARGUMENT_SHORT}|{STRING_TO_REPLACE_WITH_ARGUMENT_FULL} is provided with a value in your command." +
                        $"\nExample : {STRING_TO_REPLACE_WITH_ARGUMENT_FULL}:\"new string\".");

                stringToReplaceWith =
                    stringToReplaceWith.Replace(STRING_TO_REPLACE_WITH_ARGUMENT_SHORT + ":", string.Empty);

                stringToReplaceWith =
                    stringToReplaceWith.Replace(STRING_TO_REPLACE_WITH_ARGUMENT_FULL + ":", string.Empty);

                stringToReplaceWith = stringToReplaceWith.Replace("\"", string.Empty);

                Console.WriteLine("Application started.");

                // Load the file containing a remaining list of Excel files to process if there is one
                List<string> xlsFiles = new List<string>();

                if (File.Exists(FILES_TO_PROCESS_LIST_FILENAME))
                {
                    Console.WriteLine("A file containing a remaining list of Excel files to process has been found. Resuming process of those files...");

                    using (StreamReader reader = File.OpenText(FILES_TO_PROCESS_LIST_FILENAME))
                    {
                        string line;

                        while (!string.IsNullOrWhiteSpace(line = reader.ReadLine()))
                        {
                            xlsFiles.Add(line);
                        }
                    }

                    nbFilesToEdit = xlsFiles.Count;
                }
                else
                {
                    // List *.xls* file paths that are in the provided folder path and in the subfolders
                    Console.WriteLine("Listing files with the *.xls* extension in the provided folder path and subfolders...");

                    xlsFiles = Directory.GetFileSystemEntries(directory, "*.xls",
                        searchOption: SearchOption.AllDirectories).ToList();

                    nbFilesToEdit = xlsFiles.Count;

                    // Save list of files to process to be able to resume if there is any issue (application crash, ...)
                    using (StreamWriter writer = File.CreateText(FILES_TO_PROCESS_LIST_FILENAME))
                    {
                        foreach (string xlsFile in xlsFiles)
                        {
                            writer.WriteLine(xlsFile);
                        }
                    }
                }

                Console.WriteLine($"Number of files to process: {nbFilesToEdit}.");

                // Open an Excel instance
                Console.WriteLine("Opening Excel instance...");

                Application xlApp = new Application();
                xlApp.DisplayAlerts = false;
                xlApp.AutomationSecurity =
                    Microsoft.Office.Core.MsoAutomationSecurity
                        .msoAutomationSecurityForceDisable; // So it doesn't display a VBA error message waiting for the user to close it if there is any

                Console.WriteLine("Excel instance opened. The process of replacing VBA strings in all the Excel files will begin.");

                // Open and edit VBA script in every Excel file
                List<string> xlsPathsInFilesToProcessTxtFile = new List<string>(xlsFiles);

                foreach (string xlsFile in xlsFiles)
                {
                    try
                    {
                        // Load Excel file
                        Workbook workbook = xlApp.Workbooks.Open(xlsFile);
                        workbook.CheckCompatibility = false;
                        workbook.DoNotPromptForConvert = true;

                        // Check if there is a VBA project inside the Excel file
                        if (workbook.HasVBProject)
                        {
                            // Search and replace the old string by the new one in the VBA code
                            VBProject vbProject = workbook.VBProject;

                            foreach (VBComponent vbComponent in vbProject.VBComponents)
                            {
                                CodeModule codeModule = vbComponent.CodeModule;

                                string[] lines = null;

                                if (codeModule.CountOfLines > 0)
                                {
                                    lines = codeModule.get_Lines(1, codeModule.CountOfLines)
                                        .Split(new string[] { "\r\n" }, StringSplitOptions.None);
                                }

                                if (lines != null)
                                {
                                    for (int i = 0; i < lines.Length; i++)
                                    {
                                        if (lines[i].Contains(stringToSearch))
                                        {
                                            lines[i] = lines[i].Replace(stringToSearch, stringToReplaceWith);
                                            codeModule.ReplaceLine(i + 1, lines[i]);
                                        }
                                    }
                                }
                            }

                            // Save edited Excel file
                            workbook.Save();

                            editedFilesCount++;
                            Console.WriteLine(
                                $"Processed file: {Path.GetFileName(xlsFile)} ({editedFilesCount}/{nbFilesToEdit})");
                        }

                        // Close Excel file
                        workbook.Close();
                    }
                    catch (COMException)
                    {
                        Console.WriteLine($"ERROR WHEN PROCESSING: {xlsFile}. The file will be ignored.");

                        if (!File.Exists(FAILED_FILES_LIST_FILENAME))
                        {
                            using (_ = File.CreateText(FAILED_FILES_LIST_FILENAME)) { }
                        }

                        using (StreamWriter writer = File.AppendText(FAILED_FILES_LIST_FILENAME))
                        {
                            writer.WriteLine(xlsFile);
                        }
                    }
                    finally
                    {
                        // Delete file from the list of files to process
                        File.Delete(FILES_TO_PROCESS_LIST_FILENAME);

                        using (StreamWriter writer = File.CreateText(FILES_TO_PROCESS_LIST_FILENAME))
                        {
                            xlsPathsInFilesToProcessTxtFile.Remove(xlsFile);

                            foreach (string xlsPathInFilesToProcessTxtFile in xlsPathsInFilesToProcessTxtFile)
                            {
                                writer.WriteLine(xlsPathInFilesToProcessTxtFile);
                            }
                        }
                    }
                }

                // Close the Excel instance
                xlApp.Quit();

                // Write statistics to the console
                Console.WriteLine($"Number of processed files: {editedFilesCount}/{nbFilesToEdit}.");

                int nbUneditedFiles = nbFilesToEdit - editedFilesCount;

                if (nbUneditedFiles > 0)
                    Console.WriteLine($"{nbUneditedFiles} files have not been edited either because they didn't contain any VBA project or an error occured when loading them.");
                
                // Delete the file containing the list of Excel files to process
                File.Delete(FILES_TO_PROCESS_LIST_FILENAME);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occured during the execution of the application:");
                Console.WriteLine(ex.Message);
            }
            finally
            {
                // Stop the stopwatch and display the execution time
                sw.Stop();

                Console.WriteLine($"Execution time: {sw.Elapsed}");

                Console.WriteLine("Press Enter to close the application...");
                Console.ReadLine();
            }
        }
    }
}