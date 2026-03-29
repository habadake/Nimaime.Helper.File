using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Nimaime.Helper.File
{
	public static class ExcelHelper
	{
		/// <summary>
		/// 读取Excel表
		/// </summary>
		/// <param name="fileName">Excel文件名</param>
		/// <returns>与Excel文件相同结构的DataSet</returns>
		public static DataSet? Excel2DataSet(string fileName)
		{
			if (!System.IO.File.Exists(fileName))
				return null;

			using FileStream fs = new(fileName, FileMode.Open, FileAccess.Read);
			IWorkbook workbook = WorkbookFactory.Create(fs);

			DataSet dataSet = new();

			for (int i = 0; i < workbook.NumberOfSheets; i++)
			{
				ISheet sheet = workbook.GetSheetAt(i);
				if (sheet.PhysicalNumberOfRows == 0)
					continue;

				DataTable dt = new(sheet.SheetName);

				IRow headerRow = sheet.GetRow(sheet.FirstRowNum);
				int colCount = headerRow.LastCellNum;

				// ====== 推断列类型 ======
				Type[] columnTypes = new Type[colCount];

				for (int col = 0; col < colCount; col++)
				{
					columnTypes[col] = typeof(string); // 默认 string

					for (int row = sheet.FirstRowNum + 1; row <= sheet.LastRowNum; row++)
					{
						IRow dataRow = sheet.GetRow(row);
						if (dataRow == null) continue;

						ICell cell = dataRow.GetCell(col);
						if (cell == null) continue;

						var type = GetCellType(cell);

						if (type != typeof(string))
						{
							columnTypes[col] = type;
							break;
						}
					}
				}

				// ====== 2️⃣ 创建列 ======
				for (int col = 0; col < colCount; col++)
				{
					string colName = headerRow.GetCell(col)?.ToString() ?? $"Column{col}";
					dt.Columns.Add(colName, columnTypes[col]);
				}

				// ====== 3️⃣ 填充数据 ======
				for (int row = sheet.FirstRowNum + 1; row <= sheet.LastRowNum; row++)
				{
					IRow sheetRow = sheet.GetRow(row);
					if (sheetRow == null) continue;

					DataRow dr = dt.NewRow();

					for (int col = 0; col < colCount; col++)
					{
						ICell cell = sheetRow.GetCell(col);

						if (cell == null)
						{
							dr[col] = DBNull.Value;
							continue;
						}

						dr[col] = GetCellValue(cell, columnTypes[col]);
					}

					dt.Rows.Add(dr);
				}

				dataSet.Tables.Add(dt);
			}

			return dataSet;
		}

		private static Type GetCellType(ICell cell)
		{
			return cell.CellType switch
			{
				CellType.Boolean => typeof(bool),

				CellType.Numeric => DateUtil.IsCellDateFormatted(cell)
					? typeof(DateTime)
					: typeof(double),

				CellType.Formula => cell.CachedFormulaResultType switch
				{
					CellType.Boolean => typeof(bool),
					CellType.Numeric => DateUtil.IsCellDateFormatted(cell)
						? typeof(DateTime)
						: typeof(double),
					_ => typeof(string)
				},

				_ => typeof(string)
			};
		}

		private static object GetCellValue(ICell cell, Type targetType)
		{
			try
			{
				if (cell.CellType == CellType.Formula)
				{
					return GetFormulaValue(cell, targetType);
				}

				return targetType switch
				{
					Type t when t == typeof(bool) => cell.BooleanCellValue,

					Type t when t == typeof(DateTime) =>
						DateUtil.IsCellDateFormatted(cell)
							? cell.DateCellValue
							: DateTime.TryParse(cell.ToString(), out var dt) ? dt : DBNull.Value,

					Type t when t == typeof(double) => cell.NumericCellValue,

					_ => cell.ToString()
				};
			}
			catch
			{
				return DBNull.Value;
			}
		}

		private static object GetFormulaValue(ICell cell, Type targetType)
		{
			return cell.CachedFormulaResultType switch
			{
				CellType.Boolean => cell.BooleanCellValue,

				CellType.Numeric => targetType == typeof(DateTime)
					? cell.DateCellValue
					: cell.NumericCellValue,

				_ => cell.ToString()
			};
		}

		/// <summary>
		/// 将DataSet写入Excel文件
		/// </summary>
		/// <param name="dataSet">数据集</param>
		/// <param name="fileName">保存Excel路径</param>
		public static void SaveDataSet2Excel(DataSet dataSet, string fileName)
		{
			if (!Directory.Exists(Path.GetDirectoryName(fileName)))
			{
				Directory.CreateDirectory(Path.GetDirectoryName(fileName));
			}

			IWorkbook workbook = new XSSFWorkbook();

			foreach(DataTable dataTable in dataSet.Tables)
			{
				string tableName;
				if (string.IsNullOrWhiteSpace(dataTable.TableName))
				{
					tableName = "Sheet" + (workbook.NumberOfSheets + 1).ToString();
				}
				else
				{
					tableName = dataTable.TableName;
				}
				ISheet sheet = workbook.CreateSheet(tableName);

				// Write column headers
				IRow headerRow = sheet.CreateRow(0);
				for (int colIndex = 0; colIndex < dataTable.Columns.Count; colIndex++)
				{
					ICell cell = headerRow.CreateCell(colIndex);
					cell.SetCellValue(dataTable.Columns[colIndex].ColumnName);
				}

				// Write data rows
				for (int rowIndex = 0; rowIndex < dataTable.Rows.Count; rowIndex++)
				{
					IRow dataRow = sheet.CreateRow(rowIndex + 1);
					for (int colIndex = 0; colIndex < dataTable.Columns.Count; colIndex++)
					{
						ICell cell = dataRow.CreateCell(colIndex);
						cell.SetCellValue(dataTable.Rows[rowIndex][colIndex].ToString());
					}
				}
			}

			using FileStream fs = new(fileName, FileMode.Create, FileAccess.Write);
			workbook.Write(fs);
		}

		/// <summary>
		/// 将DataTable写入Excel文件，如果DataTable属于DataSet则写入DataSet对应的Sheet，否则写入默认Sheet
		/// </summary>
		/// <param name="dataTable">数据表</param>
		/// <param name="fileName">保存Excel路径</param>
		public static void SaveDataTable2Excel(DataTable dataTable, string fileName)
		{
			if (dataTable.DataSet != null)
			{
				SaveDataSet2Excel(dataTable.DataSet, fileName);
				return;
			}
			DataSet dataSet = new();
			dataSet.Tables.Add(dataTable);
			SaveDataSet2Excel(dataSet, fileName);
		}

		/// <summary>
		/// CSV文件转XLSX
		/// 输出在CSV同目录同文件名
		/// </summary>
		/// <param name="fileName">CSV文件路径</param>
		public static void CSV2Excel(string fileName)
		{
			if (!System.IO.File.Exists(fileName))
			{
				return;
			}
			SaveDataTable2Excel(CSVHelper.CSV2DataTable(fileName), Path.Combine(Path.GetDirectoryName(fileName), Path.GetFileNameWithoutExtension(fileName) + ".xlsx"));
		}
	}
}
