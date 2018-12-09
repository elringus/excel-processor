using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.IO;
using System.Reflection;

public class Program
{
    private static void Main (string[] args)
    {
        var exePath = Assembly.GetEntryAssembly().Location;
        var dirPath = Path.GetDirectoryName(exePath);
        var dirInfo = new DirectoryInfo(dirPath);

        var sourceFiles = dirInfo.GetFiles("*.xlsx", SearchOption.TopDirectoryOnly);
        if (sourceFiles.Length == 0) return;

        var summaryBook = new XSSFWorkbook();
        foreach (var sourceFile in sourceFiles)
            ProcessXlsFile(sourceFile, summaryBook);

        var summaryDirPath = Path.Combine(dirPath, "Summary");
        Directory.CreateDirectory(summaryDirPath);
        var summaryFilePath = Path.Combine(summaryDirPath, "Summary.xlsx");
        using (var stream = new FileStream(summaryFilePath, FileMode.Create, FileAccess.ReadWrite))
            summaryBook.Write(stream);
    }

    private static void ProcessXlsFile (FileInfo sourceFile, XSSFWorkbook summaryBook)
    {
        var stream = new FileStream(sourceFile.FullName, FileMode.Open);
        var sourceBook = new XSSFWorkbook(stream);

        for (int sheetIndex = 0; sheetIndex < sourceBook.NumberOfSheets; sheetIndex++)
        {
            if (summaryBook.NumberOfSheets < (sheetIndex + 1))
                summaryBook.CreateSheet($"Sheet{sheetIndex}");

            var summarySheet = summaryBook.GetSheetAt(sheetIndex);
            var sourceSheet = sourceBook.GetSheetAt(sheetIndex);

            ProcessSheet(sourceSheet, summarySheet);
        }

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
            catch { }
        }
    }

    private static void ProcessCell (ICell sourceCell, ICell summaryCell)
    {
        if (sourceCell.CellStyle != null)
        {
            summaryCell.CellStyle = summaryCell.Sheet.Workbook.CreateCellStyle();
            summaryCell.CellStyle.CloneStyleFrom(sourceCell.CellStyle);
        }

        if (sourceCell.CellType != CellType.Numeric)
        {
            if (summaryCell.CellType == CellType.Numeric) return;
            summaryCell.SetCellValue(sourceCell.StringCellValue);
            return;
        }

        if (summaryCell.CellType != CellType.Numeric)
            summaryCell.SetCellType(CellType.Numeric);

        var summaryValue = summaryCell.NumericCellValue + Math.Abs(sourceCell.NumericCellValue);
        summaryCell.SetCellValue(summaryValue);
    }
}
