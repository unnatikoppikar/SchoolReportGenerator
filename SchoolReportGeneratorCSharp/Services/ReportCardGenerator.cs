using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace SchoolReportGenerator.Services;

/// <summary>
/// Manages the generation of report cards.
/// Uses Open XML SDK to fill templates + Word Interop for PDF conversion.
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
        var reportCardsDir = Path.Combine(Directory.GetCurrentDirectory(), $"{className} report_cards");
        var tempDir = Path.Combine(Directory.GetCurrentDirectory(), "temp_docx");
        EnsureDirectoryExists(reportCardsDir);
        EnsureDirectoryExists(tempDir);

        // Process data
        using var dataProcessor = new DataProcessor(excelPath, mappingPath);

        // Get all student rows first to know total count
        var studentRows = dataProcessor.GetStudentRows().ToList();
        var total = studentRows.Count;
        var current = 0;

        // Collect all docx files to convert
        var docxFiles = new List<(string docxPath, string pdfPath)>();

        // Step 1: Fill templates (fast, no Word needed)
        foreach (var row in studentRows)
        {
            var studentData = dataProcessor.ProcessStudentData(row, className);
            
            if (studentData == null || !studentData.ContainsKey("name"))
            {
                continue;
            }

            current++;
            var studentName = studentData["name"];
            
            // Report progress
            progressCallback?.Invoke(current, total, $"{studentName} (filling template)");
            Console.WriteLine($"Filling template for {studentName} ({current}/{total})");

            var docxPath = Path.Combine(tempDir, $"{studentName}.docx");
            var pdfPath = Path.Combine(reportCardsDir, $"{studentName}.pdf");
            
            FillWordTemplate(templatePath, studentData, docxPath);
            docxFiles.Add((docxPath, pdfPath));
        }

        // Step 2: Convert all to PDF using Word (batch for efficiency)
        Console.WriteLine("Converting to PDF...");
        ConvertDocxToPdfBatch(docxFiles, progressCallback, total);

        // Cleanup temp files
        try { Directory.Delete(tempDir, true); } catch { }
    }

    /// <summary>
    /// Fill Word template using Open XML SDK (no Word needed).
    /// </summary>
    private void FillWordTemplate(string templatePath, Dictionary<string, string> studentData, string outputPath)
    {
        // Copy template to output
        File.Copy(templatePath, outputPath, true);

        // Open and modify
        using var doc = WordprocessingDocument.Open(outputPath, true);
        var body = doc.MainDocumentPart?.Document?.Body;
        
        if (body == null) return;

        // Replace all placeholders
        foreach (var text in body.Descendants<Text>())
        {
            foreach (var kvp in studentData)
            {
                var placeholder = "{{" + kvp.Key + "}}";
                if (text.Text.Contains(placeholder))
                {
                    text.Text = text.Text.Replace(placeholder, kvp.Value);
                }
            }
        }

        doc.MainDocumentPart?.Document?.Save();
    }

    /// <summary>
    /// Convert docx files to PDF using Word Interop (batch processing).
    /// </summary>
    private void ConvertDocxToPdfBatch(List<(string docxPath, string pdfPath)> files, Action<int, int, string>? progressCallback, int total)
    {
        dynamic? wordApp = null;
        
        try
        {
            // Create Word application
            var wordType = Type.GetTypeFromProgID("Word.Application");
            if (wordType == null)
            {
                Console.WriteLine("Microsoft Word is not installed. Keeping .docx files only.");
                // Copy docx to output folder
                foreach (var (docxPath, pdfPath) in files)
                {
                    var docxOutput = Path.ChangeExtension(pdfPath, ".docx");
                    File.Copy(docxPath, docxOutput, true);
                }
                return;
            }

            wordApp = Activator.CreateInstance(wordType);
            wordApp.Visible = false;
            wordApp.DisplayAlerts = 0; // wdAlertsNone

            var current = 0;
            foreach (var (docxPath, pdfPath) in files)
            {
                current++;
                var studentName = Path.GetFileNameWithoutExtension(docxPath);
                progressCallback?.Invoke(current, total, $"{studentName} (converting to PDF)");
                Console.WriteLine($"Converting {studentName} to PDF ({current}/{files.Count})");

                dynamic doc = wordApp.Documents.Open(Path.GetFullPath(docxPath));
                
                // WdSaveFormat.wdFormatPDF = 17
                doc.SaveAs2(Path.GetFullPath(pdfPath), 17);
                doc.Close(0); // wdDoNotSaveChanges
                
                Marshal.ReleaseComObject(doc);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"PDF conversion error: {ex.Message}");
            Console.WriteLine("Keeping .docx files only.");
            
            // Copy docx to output folder as fallback
            foreach (var (docxPath, pdfPath) in files)
            {
                var docxOutput = Path.ChangeExtension(pdfPath, ".docx");
                if (File.Exists(docxPath))
                    File.Copy(docxPath, docxOutput, true);
            }
        }
        finally
        {
            if (wordApp != null)
            {
                try
                {
                    wordApp.Quit(0);
                    Marshal.ReleaseComObject(wordApp);
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
}
