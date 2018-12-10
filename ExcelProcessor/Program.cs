using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;

public class Program
{
    private static List<string> warnings = new List<string>();
    private static string processedFilePath = string.Empty;

    private static void Main (string[] args)
    {
        var exePath = Assembly.GetEntryAssembly().Location;
        var dirPath = Path.GetDirectoryName(exePath);
        var dirInfo = new DirectoryInfo(dirPath);

        var sourceFiles = dirInfo.GetFiles("*.xlsx", SearchOption.TopDirectoryOnly);
        if (sourceFiles.Length == 0) return;

        var wrongFiles = dirInfo.GetFiles("*.xls", SearchOption.TopDirectoryOnly);
        foreach (var wrongFile in wrongFiles)
            if (wrongFile.Name.EndsWith(".xls", StringComparison.InvariantCultureIgnoreCase))
                warnings.Add($"File `{wrongFile.FullName}` has a legacy Excel format and won't be processed.");

        var summaryBook = new XSSFWorkbook();
        foreach (var sourceFile in sourceFiles)
            ProcessXlsFile(sourceFile, summaryBook);

        var summaryDirPath = Path.Combine(dirPath, "Summary");
        Directory.CreateDirectory(summaryDirPath);
        var summaryFilePath = Path.Combine(summaryDirPath, "Summary.xlsx");
        using (var stream = new FileStream(summaryFilePath, FileMode.Create, FileAccess.ReadWrite))
            summaryBook.Write(stream);

        Console.ForegroundColor = ConsoleColor.Yellow;
        foreach (var warning in warnings)
            Console.WriteLine($"WARNING: {warning}");
        Console.ForegroundColor = ConsoleColor.White;

        Console.WriteLine($"Summary created at `{summaryFilePath}`.");
        Console.WriteLine("Press any key to exit...");
        Console.ReadKey();
    }

    private static void ProcessXlsFile (FileInfo sourceFile, XSSFWorkbook summaryBook)
    {
        processedFilePath = sourceFile.FullName;

        var stream = new FileStream(sourceFile.FullName, FileMode.Open);
        var sourceBook = new XSSFWorkbook(stream);

        for (int sheetIndex = 0; sheetIndex < sourceBook.NumberOfSheets; sheetIndex++)
        {
            if (summaryBook.NumberOfSheets < (sheetIndex + 1))
                summaryBook.CreateSheet($"Sheet{sheetIndex}");

            var summarySheet = summaryBook.GetSheetAt(sheetIndex);
            var sourceSheet = sourceBook.GetSheetAt(sheetIndex);

            Console.Clear();
            Console.Write($"\rProcessing `{sourceFile.FullName}`...");
            ProcessSheet(sourceSheet, summarySheet);
        }

        Console.Clear();
        stream.Dispose();
    }

    private static void ProcessSheet (ISheet sourceSheet, ISheet summarySheet)
    {
        foreach (IRow sourceRow in sourceSheet)
        {
            var summaryRow = summarySheet.GetRow(sourceRow.RowNum) ?? summarySheet.CreateRow(sourceRow.RowNum);
            ProcessRow(sourceRow, summaryRow);
        }
    }

    private static void ProcessRow (IRow sourceRow, IRow summaryRow)
    {
        summaryRow.Height = sourceRow.Height;

        if (sourceRow.RowStyle != null)
        {
            summaryRow.RowStyle = summaryRow.Sheet.Workbook.CreateCellStyle();
            summaryRow.RowStyle.CloneStyleFrom(sourceRow.RowStyle);
        }

        foreach (var sourceCell in sourceRow)
        {
            var policy = MissingCellPolicy.RETURN_BLANK_AS_NULL;
            var summaryCell = summaryRow.GetCell(sourceCell.ColumnIndex, policy) ?? summaryRow.CreateCell(sourceCell.ColumnIndex, sourceCell.CellType);
            ProcessCell(sourceCell, summaryCell);
        }

        for (int columnIndex = 0; columnIndex < summaryRow.LastCellNum; columnIndex++)
        {
            try { summaryRow.Sheet.AutoSizeColumn(columnIndex); }
            catch { /*warnings.Add($"Failed to auto-size column with index `{columnIndex}` of `{processedFilePath}`.");*/ }
        }
    }

    private static void ProcessCell (ICell sourceCell, ICell summaryCell)
    {
        if (sourceCell.CellStyle != null)
        {
            summaryCell.CellStyle = summaryCell.Sheet.Workbook.CreateCellStyle();
            summaryCell.CellStyle.CloneStyleFrom(sourceCell.CellStyle);
        }

        if (sourceCell.CellType == CellType.Blank) return;

        if (sourceCell.CellType == CellType.Numeric)
        {
            var sourceValue = sourceCell.NumericCellValue;
            switch (summaryCell.CellType)
            {
                case CellType.Numeric:
                    summaryCell.SetCellValue(summaryCell.NumericCellValue + sourceValue);
                    return;
                case CellType.Blank:
                    summaryCell.SetCellValue(sourceValue);
                    return;
                case CellType.String:
                    if (string.IsNullOrWhiteSpace(summaryCell.StringCellValue))
                        summaryCell.SetCellValue("0");
                    var parsed = double.TryParse(summaryCell.StringCellValue, out var summaryValue);
                    if (parsed) summaryCell.SetCellValue(summaryValue + sourceValue);
                    else goto default;
                    return;
                default:
                    AddPCEWarning(1, summaryCell);
                    return;
            }
        }

        if (sourceCell.CellType == CellType.String)
        {
            switch (summaryCell.CellType)
            {
                case CellType.Numeric:
                    if (string.IsNullOrWhiteSpace(sourceCell.StringCellValue))
                        sourceCell.SetCellValue("0");
                    var parsed = double.TryParse(sourceCell.StringCellValue, out var sourceValue);
                    if (parsed) summaryCell.SetCellValue(summaryCell.NumericCellValue + sourceValue);
                    else goto default;
                    return;
                case CellType.Blank:
                case CellType.String:
                    summaryCell.SetCellValue(sourceCell.StringCellValue);
                    return;
                default:
                    AddPCEWarning(2, summaryCell);
                    return;
            }
        }
    }

    private static void AddPCEWarning (int code, ICell cell)
    {
        warnings.Add($"PCE#{code} in book `{processedFilePath}` sheet `{cell.Sheet.SheetName}` row `{cell.RowIndex + 1}` column `{ColumnNumberToName(cell.ColumnIndex + 1)}`.");
    }

    private static string ColumnNumberToName (int columnNumber)
    {
        var dividend = columnNumber;
        var columnName = string.Empty;
        int modulo;

        while (dividend > 0)
        {
            modulo = (dividend - 1) % 26;
            columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
            dividend = (dividend - modulo) / 26;
        }

        return columnName;
    }
}
