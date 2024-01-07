/////////////////////////////////////////////////////////////////////////////
// <copyright file="ExcelWrapper.cs" company="James John McGuire">
// Copyright © 2006 - 2024 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

using Common.Logging;
using System;
using System.Data.OleDb;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

using Excel = Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace DigitalZenWorks.Common.OfficeHelper
{
	/// <summary>
	/// Represents a Excel object.
	/// </summary>
	public class ExcelWrapper
	{
#pragma warning disable CA1823 // Avoid unused private fields
		private static readonly ILog Log = LogManager.GetLogger(
#pragma warning restore CA1823 // Avoid unused private fields
			System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

		private int columnCount;

		private Microsoft.Office.Interop.Excel.Application excelApplication;

		private string filename = string.Empty;
		private bool hasHeaderRow;

		private Workbook workBook;
		private Worksheet workSheet;
		private Sheets workSheets;

		/// <summary>
		/// Initializes a new instance of the <see cref="ExcelWrapper"/> class.
		/// </summary>
		public ExcelWrapper()
		{
			excelApplication = new Microsoft.Office.Interop.Excel.Application();

			excelApplication.DisplayAlerts = false;
		}

		/// <summary>
		/// Initializes a new instance of the <see cref="ExcelWrapper"/> class.
		/// </summary>
		/// <param name="fileName">The file name to use.</param>
		public ExcelWrapper(string fileName)
			: this()
		{
			this.filename = fileName;
		}

		/// <summary>
		/// Gets or sets the column count.
		/// </summary>
		/// <value>The column count.</value>
		public int ColumnCount
		{
			get { return columnCount; }
			set { columnCount = value; }
		}

		/// <summary>
		/// Gets or sets the file name.
		/// </summary>
		/// <value>The file name.</value>
		public string FileName
		{
			get { return filename; }
			set { filename = value; }
		}

		/// <summary>
		/// Gets or sets a value indicating whether has header row.
		/// </summary>
		/// <value>A value indicating whether has header row.</value>
		public bool HasHeaderRow
		{
			get { return hasHeaderRow; }
			set { hasHeaderRow = value; }
		}

		/// <summary>
		/// Gets the Header of the sheet.
		/// </summary>
		/// <value>The Header of the sheet.</value>
		public Range Header
		{
			get
			{
				// normally, we compensate for the header, here we don't
				string columnName = GetExcelColumnName(LastColumnUsed);
				string rangeQuery = "A1:" + columnName + "1";

				Range range = workSheet.get_Range(rangeQuery, Type.Missing);

				return range;
			}
		}

		/// <summary>
		/// Gets the last column used.
		/// </summary>
		/// <value>The last column used.</value>
		public int LastColumnUsed
		{
			get
			{
				int lastUsedColumn = -1;

				if (null != workSheet)
				{
					Range last = workSheet.Cells.SpecialCells(
						XlCellType.xlCellTypeLastCell, Type.Missing);

					lastUsedColumn = last.Column;
				}

				return lastUsedColumn;
			}
		}

		/// <summary>
		/// Gets the last row used.
		/// </summary>
		/// <value>The last row used.</value>
		public int LastRowUsed
		{
			get
			{
				int lastUsedRow = -1;

				if (null != workSheet)
				{
					Range last = workSheet.Cells.SpecialCells(
						XlCellType.xlCellTypeLastCell, Type.Missing);

					lastUsedRow = last.Row;

					// Excel uses a 1 based index
					lastUsedRow--;
				}

				return lastUsedRow;
			}
		}

		/// <summary>
		/// Retrieve data from the Excel spreadsheet.
		/// </summary>
		/// <param name="fileName">file name.</param>
		/// <param name="sheetName">Worksheet name.</param>
		/// <returns>DataTable Data.</returns>
		[System.Diagnostics.CodeAnalysis.SuppressMessage(
			"Microsoft.Security",
			"CA2100:Review SQL queries for security vulnerabilities",
			Justification = "For internal use")]
		public static System.Data.DataTable GetEntireSheet(
			string fileName,
			string sheetName)
		{
			System.Data.DataTable excelTable = null;

			try
			{
				string connectionString = GetConnectionString(fileName);
				excelTable = new System.Data.DataTable();
				excelTable.Locale = CultureInfo.InvariantCulture;

				using OleDbConnection connection = new (connectionString);

				string query = string.Format(
					CultureInfo.InvariantCulture,
					"SELECT * FROM [{0}$]",
					sheetName);

				using OleDbDataAdapter adaptor = new (query, connection);

				adaptor.Fill(excelTable);
			}
			catch
			{
				excelTable.Dispose();
				throw;
			}

			return excelTable;
		}

		/// <summary>
		/// Get excel column name.
		/// </summary>
		/// <param name="columnNumber">The column number.</param>
		/// <returns>The column name.</returns>
		public static string GetExcelColumnName(int columnNumber)
		{
			int dividend = columnNumber;
			string columnName = string.Empty;
			int modulo;

			while (dividend > 0)
			{
				modulo = (dividend - 1) % 26;
				columnName =
					Convert.ToChar(65 + modulo).ToString() + columnName;
				dividend = (int)((dividend - modulo) / 26);
			}

			return columnName;
		}

		/// <summary>
		/// Close method.
		/// </summary>
		public void Close()
		{
			CloseFile();

			if (excelApplication != null)
			{
				excelApplication.Quit();
				Marshal.ReleaseComObject(excelApplication);
				excelApplication = null;
			}

			GC.Collect();
			GC.WaitForPendingFinalizers();
		}

		/// <summary>
		/// Close file method.
		/// </summary>
		public void CloseFile()
		{
			if (workSheet != null)
			{
				Marshal.ReleaseComObject(workSheet);
				workSheet = null;
			}

			if (workSheets != null)
			{
				Marshal.ReleaseComObject(workSheets);
				workSheets = null;
			}

			if (workBook != null)
			{
				workBook.Close(false, null, false);
				Marshal.ReleaseComObject(workBook);
				workBook = null;
			}
		}

		/// <summary>
		/// Create method.
		/// </summary>
		/// <param name="sheetName">The sheet name.</param>
		/// <returns>The created workbook.</returns>
		public Workbook Create(string sheetName)
		{
			workBook = excelApplication.Workbooks.Add();
			workSheets = workBook.Worksheets;

			workSheet = (Worksheet)workSheets[1];
			workSheet.Name = sheetName;

			return workBook;
		}

		/// <summary>
		/// Delete method.
		/// </summary>
		/// <param name="row">The starting row.</param>
		/// <param name="column">The starting column.</param>
		/// <param name="direction">The direction of range expansion.</param>
		public void Delete(
			int row, int column, XlDeleteShiftDirection direction)
		{
			Range range = GetCell(row, column);

			range.Delete(direction);
			Marshal.ReleaseComObject(range);
		}

		/// <summary>
		/// Delete row method.
		/// </summary>
		/// <param name="row">The row to delete.</param>
		public void DeleteRow(int row)
		{
			Range range = GetRange(row, row, 0, LastColumnUsed);

			range.Delete(XlDeleteShiftDirection.xlShiftUp);
			Marshal.ReleaseComObject(range);
		}

		/// <summary>
		/// Find excel worksheet method.
		/// </summary>
		/// <param name="worksheetName">The worksheet to find.</param>
		/// <returns>A value indicating whether the worksheet was found.</returns>
		public bool FindExcelWorksheet(string worksheetName)
		{
			bool sheetFound = false;

			if (workSheets != null)
			{
				// Step through the worksheet collection and see if the sheet
				// is available. If found return true;
				for (int index = 1; index <= workSheets.Count; index++)
				{
					Worksheet testSheet =
						(Worksheet)workSheets.get_Item((object)index);
					if (testSheet.Name.Equals(
						worksheetName, StringComparison.OrdinalIgnoreCase))
					{
						// Get method interface
						_Worksheet sheet = (_Worksheet)testSheet;
						sheet.Activate();
						workSheet = testSheet;
						sheetFound = true;
						break;
					}
				}
			}

			return sheetFound;
		}

		/// <summary>
		/// The get cell method.
		/// </summary>
		/// <param name="row">The row to use.</param>
		/// <param name="column">The column to  use.</param>
		/// <returns>The cell range.</returns>
		public Range GetCell(int row, int column)
		{
			row = AdjustRow(row);
			if (column < int.MaxValue)
			{
				// excel is 1 based
				column++;
			}

			Range range = (Range)workSheet.Cells[row, column];

			return range;
		}

		/// <summary>
		/// Get cell background color.
		/// </summary>
		/// <param name="row">The row to use.</param>
		/// <param name="column">The column to use.</param>
		/// <returns>The background color code.</returns>
		public double GetCellBackgroundColor(int row, int column)
		{
			Range range = GetCell(row, column);

			double color = (double)range.Interior.Color;

			return color;
		}

		/// <summary>
		/// Get cell background color index.
		/// </summary>
		/// <param name="row">The row to use.</param>
		/// <param name="column">The column to use.</param>
		/// <returns>The background color index.</returns>
		public int GetCellBackgroundColorIndex(int row, int column)
		{
			Range range = GetCell(row, column);

			int color = (int)range.Interior.ColorIndex;

			return color;
		}

		/// <summary>
		/// Get cell font color.
		/// </summary>
		/// <param name="row">The row to use.</param>
		/// <param name="column">The column to use.</param>
		/// <returns>The cell font color code.</returns>
		public double GetCellFontColor(int row, int column)
		{
			Range range = GetCell(row, column);

			double color = (double)range.Font.Color;

			return color;
		}

		/// <summary>
		/// Get cell font color index.
		/// </summary>
		/// <param name="row">The row to use.</param>
		/// <param name="column">The column to use.</param>
		/// <returns>The celll font color index.</returns>
		public int GetCellFontColorIndex(int row, int column)
		{
			Range range = GetCell(row, column);

			int color = (int)range.Font.ColorIndex;

			return color;
		}

		/// <summary>
		/// Get cell value.
		/// </summary>
		/// <param name="row">The row to use.</param>
		/// <param name="column">The column to use.</param>
		/// <returns>The cell value.</returns>
		public string GetCellValue(int row, int column)
		{
			string cellValue = null;
			Range cell = GetCell(row, column);

			if (null != cell.Value2)
			{
				cellValue = cell.Value2.ToString();
			}

			return cellValue;
		}

		/// <summary>
		/// Get column range.
		/// </summary>
		/// <param name="columnNumber">The column number to use.</param>
		/// <returns>The column range.</returns>
		public Range GetColumnRange(int columnNumber)
		{
			Range range = GetRange(0, LastRowUsed, columnNumber, columnNumber);

			range = range.EntireColumn;

			return range;
		}

		/// <summary>
		/// Get count of non-empty cells.
		/// </summary>
		/// <param name="range">The range to check.</param>
		/// <returns>The count of non-empty cells.</returns>
		public int GetCountNonemptyCells(Range range)
		{
			double result = excelApplication.WorksheetFunction.CountA(range);

			return Convert.ToInt32(result);
		}

		/// <summary>
		/// Get entire sheet method.
		/// </summary>
		/// <returns>The entire sheet.</returns>
		public string[][] GetEntireSheet()
		{
			string[][] values = null;

			if (null != workSheet)
			{
				values = GetRangeValues(0, LastRowUsed, 0, LastColumnUsed);
			}

			return values;
		}

		/// <summary>
		/// The get range method.
		/// </summary>
		/// <param name="rowBegin">The row to begin with.</param>
		/// <param name="rowEnd">The row to end with.</param>
		/// <param name="columnBegin">The column to begin with.</param>
		/// <param name="columnEnd">The column to end with.</param>
		/// <returns>The range.</returns>
		public Range GetRange(
			int rowBegin, int rowEnd, int columnBegin, int columnEnd)
		{
			// excel is 1 based
			rowBegin = AdjustRow(rowBegin);
			rowEnd = AdjustRow(rowEnd);
			if (columnBegin < int.MaxValue)
			{
				columnBegin++;
			}

			string columnBeginName = GetExcelColumnName(columnBegin);
			string columnEndName = GetExcelColumnName(columnEnd);

			string rangeQuery =
				columnBeginName + rowBegin + ":" + columnEndName + rowEnd;

			Range range = workSheet.get_Range(rangeQuery, Type.Missing);

			return range;
		}

		/// <summary>
		/// Get range values.
		/// </summary>
		/// <param name="rowBegin">The row to begin with.</param>
		/// <param name="rowEnd">The row to end with.</param>
		/// <param name="columnBegin">The column to begin with.</param>
		/// <param name="columnEnd">The column to end with.</param>
		/// <returns>The range values.</returns>
		public string[][] GetRangeValues(
			int rowBegin, int rowEnd, int columnBegin, int columnEnd)
		{
			Range range = GetRange(rowBegin, rowEnd, columnBegin, columnEnd);

			string[][] stringArray = GetStringArray(range.Cells.Value2);

			Marshal.ReleaseComObject(range);

			return stringArray;
		}

		/// <summary>
		/// Get row values.
		/// </summary>
		/// <param name="rowId">The row id to use.</param>
		/// <returns>The array of row values.</returns>
		public string[] GetRowValues(int rowId)
		{
			string[][] rows = GetRangeValues(rowId, rowId, 0, LastColumnUsed);
			string[] row = rows[0];

			return row;
		}

		/// <summary>
		/// Get row method.
		/// </summary>
		/// <param name="rowId">The row id to use.</param>
		/// <returns>The range of the row.</returns>
		public Range GetRow(int rowId)
		{
			int lastUsedColumn = LastColumnUsed;

			Range range = GetRange(rowId, rowId, 0, lastUsedColumn);

			return range;
		}

		/// <summary>
		/// Is cell empty method.
		/// </summary>
		/// <param name="row">The row to use.</param>
		/// <param name="column">The column to use.</param>
		/// <returns>A value indicating whether the cell is empty or
		/// not.</returns>
		public bool IsCellEmpty(int row, int column)
		{
			bool empty = false;

			string contents = GetCellValue(row, column);

			if (string.IsNullOrWhiteSpace(contents))
			{
				empty = true;
			}

			return empty;
		}

		/// <summary>
		/// The open file method.
		/// </summary>
		/// <returns>A value indicating success or not.</returns>
		public bool OpenFile()
		{
			return OpenFile(filename);
		}

		/// <summary>
		/// The open file method.
		/// </summary>
		/// <param name="fileName">The file name to open.</param>
		/// <param name="readOnly">Indicates whether the file should be
		/// opened in read only mode or not.</param>
		/// <returns>A value indicating success or not.</returns>
		public bool OpenFile(string fileName, bool readOnly = false)
		{
			bool result = false;

			if ((!string.IsNullOrEmpty(fileName)) &&
				File.Exists(fileName))
			{
				filename = fileName;
				workBook = excelApplication.Workbooks.Open(
					fileName,
					0,
					readOnly,
					1,
					true,
					System.Reflection.Missing.Value,
					System.Reflection.Missing.Value,
					true,
					System.Reflection.Missing.Value,
					true,
					System.Reflection.Missing.Value,
					false,
					System.Reflection.Missing.Value,
					false,
					false);

				if (workBook != null)
				{
					workSheets = workBook.Worksheets;
					workSheet = (Worksheet)workSheets[1];
				}

				result = true;
			}

			return result;
		}

		/// <summary>
		/// Save method.
		/// </summary>
		public void Save()
		{
			workBook.SaveAs(
				filename,
				XlFileFormat.xlWorkbookDefault,
				null,
				null,
				false,
				false,
				XlSaveAsAccessMode.xlExclusive,
				XlSaveAsAccessMode.xlExclusive,
				System.Reflection.Missing.Value,
				System.Reflection.Missing.Value,
				System.Reflection.Missing.Value,
				System.Reflection.Missing.Value);
		}

		/// <summary>
		/// Save as CSV method.
		/// </summary>
		/// <param name="filePath">The file path to save to.</param>
		public void SaveAsCsv(string filePath)
		{
			workBook.SaveAs(
				filePath,
				XlFileFormat.xlCSVWindows,
				Type.Missing,
				Type.Missing,
				Type.Missing,
				Type.Missing,
				XlSaveAsAccessMode.xlNoChange,
				XlSaveConflictResolution.xlLocalSessionChanges,
				false,
				Type.Missing,
				Type.Missing,
				true);
		}

		/// <summary>
		/// Set background color method.
		/// </summary>
		/// <param name="row">The row to use.</param>
		/// <param name="column">The column to use.</param>
		/// <param name="color">The color code to set to.</param>
		public void SetBackgroundColor(int row, int column, double color)
		{
			Range range = GetCell(row, column);
			range.Interior.Color = color;

			Marshal.ReleaseComObject(range);
		}

		/// <summary>
		/// Set cell method.
		/// </summary>
		/// <param name="row">The row to use.</param>
		/// <param name="column">The column to use.</param>
		/// <param name="value">The value to set to.</param>
		public void SetCell(int row, int column, string value)
		{
			Range cell = GetCell(row, column);

			cell.Value = value;
		}

		/// <summary>
		/// Set column format method.
		/// </summary>
		/// <param name="column">The column to use.</param>
		/// <param name="format">The format to set to.</param>
		public void SetColumnFormat(int column, Format format)
		{
			Range columnRange = GetColumnRange(column);

			switch (format)
			{
				case Format.Date:
					columnRange.NumberFormat = "yyyy-MM-dd";
					break;
				case Format.Text:
					columnRange.NumberFormat = "@";
					break;
				default:
					break;
			}

			Marshal.ReleaseComObject(columnRange);
		}

		/// <summary>
		/// Set font color method.
		/// </summary>
		/// <param name="row">The row to use.</param>
		/// <param name="column">The column to use.</param>
		/// <param name="color">The color to set to.</param>
		public void SetFontColor(int row, int column, Color color)
		{
			Range range = GetCell(row, column);
			range.Font.Color = System.Drawing.ColorTranslator.ToOle(color);

			Marshal.ReleaseComObject(range);
		}

		/// <summary>
		/// Set font color method.
		/// </summary>
		/// <param name="row">The row to use.</param>
		/// <param name="column">The column to use.</param>
		/// <param name="color">The color code to set to.</param>
		public void SetFontColor(int row, int column, double color)
		{
			Range range = GetCell(row, column);
			range.Font.Color = color;

			Marshal.ReleaseComObject(range);
		}

		/// <summary>
		/// Set row method.
		/// </summary>
		/// <param name="row">The row to use.</param>
		/// <param name="data">The data array to set to.</param>
		public void SetRow(int row, string[] data)
		{
			if ((null != data) && (data.Length > 0))
			{
				Range range = GetRange(row, row, 0, LastColumnUsed);

				range.get_Resize(1, data.Length).Value2 = data;

				Marshal.ReleaseComObject(range);
			}
		}

		/// <summary>
		/// Set text format method.
		/// </summary>
		/// <param name="row">The row to use.</param>
		/// <param name="column">The column to use.</param>
		public void SetTextFormat(int row, int column)
		{
			Range range = GetCell(row, column);
			range.NumberFormat = "@";

			Marshal.ReleaseComObject(range);
		}

		/// <summary>
		/// Set worksheet name method.
		/// </summary>
		/// <param name="sheetName">The sheet name.</param>
		public void SetWorksheetName(string sheetName)
		{
			workSheet.Name = sheetName;
		}

		private static string GetConnectionString(string fileName)
		{
			string connectionString =
				"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
				fileName + ";Extended Properties='Excel 12.0 Xml;HDR=Yes'";

			return connectionString;
		}

		private static string[][] GetStringArray(object rangeValues)
		{
			string[][] stringArray = null;

			if (rangeValues is Array array)
			{
				int rank = array.Rank;
				if (rank > 1)
				{
					int rowCount = array.GetLength(0);
					int columnCount = array.GetUpperBound(1);

					stringArray = new string[rowCount][];

					for (int index = 0; index < rowCount; index++)
					{
						stringArray[index] = new string[columnCount];

						for (int index2 = 0; index2 < columnCount; index2++)
						{
							object obj = array.GetValue(index + 1, index2 + 1);

							if (null != obj)
							{
								string value = obj.ToString();

								stringArray[index][index2] = value;
							}
							else
							{
								// design choice - empty seems to reflect more
								// of the excel model
								stringArray[index][index2] = string.Empty;
							}
						}
					}
				}
			}

			return stringArray;
		}

		/// <summary>
		/// Excel uses a 1 based index. Programs using this, use a 0 based
		/// index, so need to adjust.  Also, compensate, if there is a header.
		/// </summary>
		/// <param name="row">The row to adjust.</param>
		/// <returns>The adjusted row index.</returns>
		private int AdjustRow(int row)
		{
			row++;

			if (true == hasHeaderRow)
			{
				row++;
			}

			return row;
		}
	}
}