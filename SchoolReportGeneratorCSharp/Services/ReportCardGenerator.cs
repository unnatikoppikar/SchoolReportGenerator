using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

namespace SchoolReportGenerator.Services;

/// <summary>
/// Manages the generation of report cards.
/// Uses Word Interop for template filling AND PDF conversion.
/// This matches the original Python approach and handles placeholders correctly.
/// </summary>
public class ReportCardGenerator
{
    /// <summary>
    /// Generate report cards for a given class.
    /// </summary>
    public void GenerateReportCards(string excelPath, string templatePath, string mappingPath, string className, Action<int, int, string>? progressCallback = null)
    {
        Console.WriteLine($"excel_path: {excelPath}");
        Console.WriteLine($"template_path: {templatePath}");
        Console.WriteLine($"mapping_path: {mappingPath}");

        // Setup directories
        var reportCardsDir = Path.GetFullPath($"{className} report_cards");
        EnsureDirectoryExists(reportCardsDir);

        // Process data
        using var dataProcessor = new DataProcessor(excelPath, mappingPath);

        // Get all student rows first to know total count
        var studentRows = dataProcessor.GetStudentRows().ToList();
        var total = studentRows.Count;

        if (total == 0)
        {
            Console.WriteLine("No student data found!");
            return;
        }

        // Check if Word is available
        var wordType = Type.GetTypeFromProgID("Word.Application");
        if (wordType == null)
        {
            Console.WriteLine("ERROR: Microsoft Word is not installed. Cannot generate report cards.");
            return;
        }

        dynamic? wordApp = null;
        
        try
        {
            // Create Word application ONCE for all documents
            wordApp = Activator.CreateInstance(wordType);
            wordApp.Visible = false;
            wordApp.DisplayAlerts = 0; // wdAlertsNone

            var current = 0;

            foreach (var row in studentRows)
            {
                var studentData = dataProcessor.ProcessStudentData(row, className);
                
                if (studentData == null || !studentData.ContainsKey("name"))
                {
                    Console.WriteLine("Skipping row - no name found");
                    continue;
                }

                current++;
                var studentName = studentData["name"];
                var safeFileName = SanitizeFileName(studentName);
                
                // Report progress
                progressCallback?.Invoke(current, total, studentName);
                Console.WriteLine($"Generating report card for {studentName} ({current}/{total})");

                var pdfPath = Path.Combine(reportCardsDir, $"{safeFileName}.pdf");

                // Open template, fill placeholders, save as PDF
                GenerateSingleReport(wordApp, templatePath, studentData, pdfPath);
            }

            Console.WriteLine($"\nCompleted! Generated {current} report cards in: {reportCardsDir}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"ERROR: {ex.Message}");
            Console.WriteLine(ex.StackTrace);
        }
        finally
        {
            if (wordApp != null)
            {
                try
                {
                    wordApp.Quit(0); // wdDoNotSaveChanges
                    Marshal.ReleaseComObject(wordApp);
                }
                catch { }
            }
            
            // Force garbage collection to release COM objects
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }

    /// <summary>
    /// Generate a single report card using Word.
    /// Opens template, replaces ALL placeholders, saves as PDF.
    /// </summary>
    private void GenerateSingleReport(dynamic wordApp, string templatePath, Dictionary<string, string> studentData, string pdfPath)
    {
        dynamic? doc = null;
        
        try
        {
            // Open template
            doc = wordApp.Documents.Open(Path.GetFullPath(templatePath), ReadOnly: false);

            // Replace ALL placeholders using Word's Find & Replace
            // This handles text split across multiple runs correctly!
            foreach (var kvp in studentData)
            {
                var placeholder = "{{" + kvp.Key + "}}";
                var value = kvp.Value ?? "---";
                
                // Use Word's Find & Replace - handles split runs correctly
                var find = doc.Content.Find;
                find.ClearFormatting();
                find.Replacement.ClearFormatting();
                
                // Find and replace
                find.Execute(
                    FindText: placeholder,
                    MatchCase: false,
                    MatchWholeWord: false,
                    MatchWildcards: false,
                    MatchSoundsLike: false,
                    MatchAllWordForms: false,
                    Forward: true,
                    Wrap: 1, // wdFindContinue
                    Format: false,
                    ReplaceWith: value,
                    Replace: 2 // wdReplaceAll
                );
            }

            // Save as PDF (WdSaveFormat.wdFormatPDF = 17)
            doc.SaveAs2(Path.GetFullPath(pdfPath), 17);
            
            Console.WriteLine($"  -> Saved: {Path.GetFileName(pdfPath)}");
        }
        finally
        {
            if (doc != null)
            {
                try
                {
                    doc.Close(0); // wdDoNotSaveChanges
                    Marshal.ReleaseComObject(doc);
                }
                catch { }
            }
        }
    }

    private void EnsureDirectoryExists(string directory)
    {
        if (!Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }
    }

    /// <summary>
    /// Sanitize filename by removing invalid characters.
    /// Windows doesn't allow: \ / : * ? " < > |
    /// </summary>
    private string SanitizeFileName(string fileName)
    {
        var invalidChars = Path.GetInvalidFileNameChars();
        var sanitized = string.Join("_", fileName.Split(invalidChars, StringSplitOptions.RemoveEmptyEntries));
        return sanitized.Trim();
    }
}
