using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.OpenXml4Net;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.XSSF.UserModel;

namespace ImportacaoCEP
{
	public class ImportarCEP
	{
		private ISheet GetFileStream(string fullFilePath)
		{
			var fileExtension = Path.GetExtension(fullFilePath);
			string sheetName;
			ISheet sheet = null;
			switch (fileExtension)
			{
				case ".xlsx":
					using (var fs = new FileStream(fullFilePath, FileMode.Open, FileAccess.Read))
					{
						var wb = new XSSFWorkbook(fs);
						sheetName = wb.GetSheetAt(0).SheetName;
						sheet = (XSSFSheet)wb.GetSheet(sheetName);
					}
					break;
				case ".xls":
					using (var fs = new FileStream(fullFilePath, FileMode.Open, FileAccess.Read))
					{
						var wb = new HSSFWorkbook(fs);
						sheetName = wb.GetSheetAt(0).SheetName;
						sheet = (HSSFSheet)wb.GetSheet(sheetName);
					}
					break;
			}
			return sheet;
		}

		public DataTable Importar(string fullFilePath)
		{
			try
			{
				var sh = GetFileStream(fullFilePath);
				var dtExcelTable = new DataTable();
				dtExcelTable.Rows.Clear();
				dtExcelTable.Columns.Clear();
				var headerRow = sh.GetRow(0);
				int colCount = headerRow.LastCellNum;
				for (var c = 0; c < colCount; c++)
					dtExcelTable.Columns.Add(headerRow.GetCell(c).ToString());
				var i = 1;
				var currentRow = sh.GetRow(i);
				while (currentRow != null)
				{
					var dr = dtExcelTable.NewRow();
					for (var j = 0; j < currentRow.Cells.Count; j++)
					{
						var cell = currentRow.GetCell(j);

						if (cell != null)
							switch (cell.CellType)
							{
								case CellType.Numeric:
									dr[j] = DateUtil.IsCellDateFormatted(cell)
										? cell.DateCellValue.ToString(CultureInfo.InvariantCulture)
										: cell.NumericCellValue.ToString(CultureInfo.InvariantCulture);
									break;
								case CellType.String:
									dr[j] = cell.StringCellValue;
									break;
								case CellType.Blank:
									dr[j] = string.Empty;
									break;
							}
					}
					dtExcelTable.Rows.Add(dr);
					i++;
					currentRow = sh.GetRow(i);
				}
				return dtExcelTable;
			}
			catch (Exception e)
			{
				throw;
			}
		}


	}
}
