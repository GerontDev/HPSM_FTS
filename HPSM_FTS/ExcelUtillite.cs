using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace HPSM_FTS
{
	public class ExcelUtillite
	{
		private readonly NLog.ILogger Log;

		public ExcelUtillite(NLog.ILogger logger)
		{
			Log = logger;
		}

		public void SaveExcel_Mo(DataResult data, string NameFileExcel)
		{
			Excel.Application excel = null;
			Excel.Workbook wb = null;
			try
			{
				excel = new Excel.Application();
				//excel.Visible = true;
				wb = (Excel.Workbook)(excel.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet));
				System.Threading.Thread.Sleep(1000);
				Excel.Worksheet worksheet_last = null;

				{
					Excel.Worksheet worksheet_report1 = wb.Worksheets.Add(Type.Missing, (worksheet_last == null) ? Type.Missing : worksheet_last, Type.Missing, Type.Missing);
					worksheet_last = worksheet_report1;
					worksheet_report1.Name = string.Format("открытозакрыто");
					string[] columns_report1 = new string[]
					{
					"Дата",
					"Открыто заявок",
					"Завершено заявок"
					};
					int iColumn = 1;
					foreach (var column_name in columns_report1)
					{
						Excel.Range r = worksheet_report1.Cells[1, iColumn];
						r.Value = column_name;
						iColumn++;
					}
					int iRow = 2;

					DateTime fest = data.Report1.Keys.Select(q => DateTime.Parse(q)).Min();
					DateTime last = data.Report1.Keys.Select(q => DateTime.Parse(q)).Max();
					DateTime festDay = new DateTime(fest.Year, fest.Month, 1);
					DateTime lastDay = new DateTime(last.Year, last.Month, DateTime.DaysInMonth(last.Year, last.Month));

					for (DateTime i = festDay; i <= lastDay; i = i.AddDays(1))
					{
						string currentdate = i.ToShortDateString();
						Excel.Range range_data_row = worksheet_report1.Range[worksheet_report1.Cells[iRow, 1], worksheet_report1.Cells[iRow, columns_report1.Length]];

						if (data.Report1.ContainsKey(currentdate))
						{
							range_data_row.Value = new object[]
								{
								currentdate,
								data.Report1[currentdate].OpenedCount,
								data.Report1[currentdate].ClosedCount,
								};
						}
						else
						{
							range_data_row.Value = new object[] { currentdate, null, null, };
						}
						if (i.DayOfWeek == DayOfWeek.Sunday || i.DayOfWeek == DayOfWeek.Saturday)
							range_data_row.Interior.Color = Excel.XlRgbColor.rgbLightBlue;
						iRow++;
					}

					//foreach (var rect in data.Report1)
					//{
					//	Excel.Range range_data_row = worksheet_report1.Range[worksheet_report1.Cells[iRow, 1], worksheet_report1.Cells[iRow, columns_report1.Length]];
					//	range_data_row.Value = new object[]
					//		{
					//			rect.Key,
					//			rect.Value.OpenedCount,
					//			rect.Value.ClosedCount,
					//		};
					//	iRow++;
					//}
					Excel.Range range = worksheet_report1.Range[worksheet_report1.Cells[1, 1], worksheet_report1.Cells[iRow - 1, iColumn - 1]];
					range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
					range.Borders.Weight = Excel.XlBorderWeight.xlThin;
					worksheet_report1.Columns.AutoFit();
				}

				{
					Excel.Worksheet worksheet_report2 = wb.Worksheets.Add(Type.Missing, (worksheet_last == null) ? Type.Missing : worksheet_last, Type.Missing, Type.Missing);
					worksheet_last = worksheet_report2;
					worksheet_report2.Name = string.Format("Акты выполненных работ");
					string[] columns_report2 = new string[]
					{
					"№ акт.",
					"Нормер заявки",
					"Этап",
					"ФИО пользователя",
					"Контакные данные",
					"IP адрес/сетевое имя:",
					"Наименование и серийный номер средства вычислительной техники",
					"Время начала работ",
					"Время окончания работ",
					"Приоритет",
					"Выполненная процедура",
					"Суть проблемы",
					"Проведенные работы",
					"Представитель Заказчика"
					};
					int iColumn = 1;
					foreach (var column_name in columns_report2)
					{
						Excel.Range r = worksheet_report2.Cells[1, iColumn];
						r.Value = column_name;
						iColumn++;
					}

					List<Phase> PhaseList = new List<Phase>();
					PhaseList.Add(new Phase(1, new DateTime(2018, 12, 28), new DateTime(2019, 2, 28)));
					PhaseList.Add(new Phase(2, new DateTime(2019, 3, 01), new DateTime(2019, 5, 31)));
					PhaseList.Add(new Phase(3, new DateTime(2019, 6, 01), new DateTime(2019, 8, 31)));
					PhaseList.Add(new Phase(4, new DateTime(2019, 9, 01), new DateTime(2019, 12, 10)));
					PhaseList.Add(new Phase(5, new DateTime(2019, 12, 11), new DateTime(2020, 2, 29)));
					PhaseList.Add(new Phase(6, new DateTime(2020, 3, 01), new DateTime(2020, 5, 31)));
					PhaseList.Add(new Phase(7, new DateTime(2020, 6, 01), new DateTime(2020, 8, 31)));
					PhaseList.Add(new Phase(8, new DateTime(2020, 9, 01), new DateTime(2020, 12, 10)));

					int iRow = 2;
					foreach (Incendent item in data.IncendentList)
					{
						Excel.Range range_data_row = worksheet_report2.Range[worksheet_report2.Cells[iRow, 1], worksheet_report2.Cells[iRow, columns_report2.Length]];
						Phase P = PhaseList.FirstOrDefault(q => q.Begin <= item.Opened && item.Opened <= q.End);
						range_data_row.Value = new object[]
							{
								iRow -1,
								item.ENC,
								P == null ? "-" : P.ToString(),
								item.Applicant,
								"",
								"",
								"",
								item.Opened,
								item.Closed,
								item.Priority,
								item.ВидРаботы,
								item.Описание,
								item.Решение,
								item.Subsystem == "АСВДТО" ? "Сутягин А.Н." : "Карпунина Т.Н" 
							};						
						iRow++;
					}

					//foreach (var rect in data.Report1)
					//{
					//	Excel.Range range_data_row = worksheet_report1.Range[worksheet_report1.Cells[iRow, 1], worksheet_report1.Cells[iRow, columns_report1.Length]];
					//	range_data_row.Value = new object[]
					//		{
					//			rect.Key,
					//			rect.Value.OpenedCount,
					//			rect.Value.ClosedCount,
					//		};
					//	iRow++;
					//}
					Excel.Range range = worksheet_report2.Range[worksheet_report2.Cells[1, 1], worksheet_report2.Cells[iRow - 1, iColumn - 1]];
					range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
					range.Borders.Weight = Excel.XlBorderWeight.xlThin;
					worksheet_report2.Columns.AutoFit();
				}
				this.Log.Trace(string.Format("Данные сохранены  в файл {0}", NameFileExcel));
			}
			catch (Exception ex)
			{
				this.Log.Warn(string.Format("Ошибка: {0}\r\nФайла:{1}", ex.Message, NameFileExcel));
				throw;
			}
			finally
			{
				try
				{
					if (wb != null)
					{
						wb.Close(true, NameFileExcel, Type.Missing);
						wb = null;
					}
					if (excel != null)
					{
						excel.Quit();
						excel = null;
					}
					GC.Collect();
					GC.WaitForPendingFinalizers();
				}
				catch (Exception ex1)
				{
					this.Log.Warn(string.Format("Ошибка при закрытии : {0}\r\nФайла:{1}", ex1.Message, NameFileExcel));
				}
			}
		}

		public bool IsRowEmptyAnd(Excel.Worksheet xlWorkSheet, int row_idx, int[] column_indexs_check)
		{
			foreach (int idx_row in column_indexs_check)
			{
				Excel.Range rStart = xlWorkSheet.Cells[row_idx, idx_row];
				if (rStart.Value == null)
					return true;
			}
			return false;
		}

		public bool IsRowEmptyAnd(Object[,] value, int row_idx, int[] column_indexs_check)
		{
			foreach (int idx_row in column_indexs_check)
			{				
				if (value.GetLength(0) < row_idx || value[row_idx, idx_row] == null)
					return true;
			}
			return false;
		}

		public bool IsRowEmptyOr(Excel.Worksheet xlWorkSheet, int row_idx, int[] column_indexs_check)
		{
			foreach (int idx_row in column_indexs_check)
			{
				Excel.Range rStart = xlWorkSheet.Cells[row_idx, idx_row];
				if (rStart.Value != null)
					return false;
			}
			return true;
		}
		/// <summary>
		/// 
		/// </summary>
		/// <param name="PathExcel"></param>
		/// <param name="row_indexs_check"></param>
		/// <param name="countemptyrow"></param>
		/// <param name="WorksheetNames"></param>
		/// <returns></returns>
		public Dictionary<string, TableIner> LoadExcelAllTable(string PathExcel, int[] column_indexs_check_row, int countemptyrow, int? max_count_column, HashSet<string> WorksheetNames)
		{
			var list = new Dictionary<string, TableIner>();
			Excel.Application excel = null;
			Excel.Workbook wb = null;

			try
			{
				excel = new Excel.Application();
				wb = excel.Workbooks.Open(PathExcel);
				foreach (Excel.Worksheet xlWorkSheet in wb.Worksheets)
				{
					string WorkSheetName = xlWorkSheet.Name;

					if (WorksheetNames != null)
					{
						if (!WorksheetNames.Contains(WorkSheetName))
							continue;
					}
					
					this.Log.Trace(string.Format("Загрзука Excel Worksheet \"{0}\"", WorkSheetName));
					var watch_count = System.Diagnostics.Stopwatch.StartNew();

					object[,] row_value_all = (object[,])xlWorkSheet.UsedRange.Value;
					int WorkSheetColumnsCount = row_value_all.GetLength(1);
					int WorkSheetRowsCount = row_value_all.GetLength(0);

					List<string> ColumntList = new List<string>();
					for (int col_idx = 1; col_idx <= WorkSheetColumnsCount && (max_count_column == null || col_idx < max_count_column); col_idx++)
					{
						object value = row_value_all[1, col_idx];
						if (value == null)
							ColumntList.Add("");
						else
							ColumntList.Add(value.ToString());
					}
					TableIner table = new TableIner(ColumntList.ToArray());
					int countemptyrow_begin = countemptyrow;
					for (int row_idx = 2; row_idx < WorkSheetRowsCount; row_idx++)
					{
						if (IsRowEmptyAnd(row_value_all, row_idx, column_indexs_check_row))
						{
							countemptyrow_begin--;
							if (countemptyrow_begin > 0)
								continue;
							break;
						}
						countemptyrow_begin = countemptyrow;
						object[] row = new object[table.Column.Length];

						for (int col_idx = 1; col_idx <= ColumntList.Count; col_idx++)
							row[col_idx - 1] = row_value_all[row_idx, col_idx];

						table.AddRow(row);
					}
					list.Add(WorkSheetName, table);
					watch_count.Stop();
					Log.Trace(string.Format("Загрузен Excel Worksheet \"{0}\" за время {1}. Размерность {2}x{3} результируюшая {4}х{5}", WorkSheetName, watch_count.Elapsed, row_value_all.GetLength(0), row_value_all.GetLength(1), table.Row.Count, table.Column.Length));
				}
				return list;
			}
			catch (Exception ex)
			{
				this.Log.Trace(string.Format("Ошибка: {0}\r\nФайла:{1}", ex.Message, PathExcel));
				throw;
			}
			finally
			{
				try
				{
					if (wb != null)
					{
						wb.Close(false, Type.Missing, Type.Missing);
						wb = null;
					}
					if (excel != null)
					{
						excel.Quit();
						excel = null;
					}
					GC.Collect();
					GC.WaitForPendingFinalizers();
					//System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
				}
				catch (Exception ex1)
				{
					this.Log.Trace(string.Format("Ошибка при закрытии : {0}\r\nФайла:{1}", ex1.Message, PathExcel));
				}
			}
		}
		/// <summary>
		/// 
		/// </summary>
		/// <param name="PathExcel"></param>
		/// <param name="row_indexs_check"></param>
		/// <param name="countemptyrow"></param>
		/// <param name="WorksheetNames"></param>
		/// <returns></returns>
		public Dictionary<string, TableIner> LoadExcelAllTable_only_Data(string PathExcel, int[] column_indexs_check_row, int countemptyrow, int count_column, HashSet<string> WorksheetNames)
		{
			TableIner table = null;

			var list = new Dictionary<string, TableIner>();

			Excel.Application excel = null;
			Excel.Workbook wb = null;

			try
			{
				excel = new Excel.Application();
				wb = excel.Workbooks.Open(PathExcel);
				foreach (Excel.Worksheet xlWorkSheet in wb.Worksheets)
				{
					if (WorksheetNames != null)
					{
						if (!WorksheetNames.Contains(xlWorkSheet.Name))
							continue;
					}
					this.Log.Trace(string.Format("Загрзука Excel Worksheet \"{0}\"", xlWorkSheet.Name));
					List<string> ColumntList = new List<string>();
					for (int col_idx = 1; col_idx <= count_column; col_idx++)
					{
						Excel.Range r = xlWorkSheet.Cells[1, col_idx];
						if (r.Value == null)
							ColumntList.Add("");
						else
							ColumntList.Add(r.Text);
					}

					table = new TableIner(ColumntList.ToArray());
					int countemptyrow_begin = countemptyrow;
					for (int row_idx = 2; row_idx < xlWorkSheet.Rows.Count; row_idx++)
					{
						if (IsRowEmptyOr(xlWorkSheet, row_idx, column_indexs_check_row))
						{
							countemptyrow_begin--;
							if (countemptyrow_begin > 0)
								continue;
							break;
						}

						countemptyrow_begin = countemptyrow;

						object[] row = new object[table.Column.Length];

						for (int col_idx = 1; col_idx <= ColumntList.Count; col_idx++)
						{
							Excel.Range r = xlWorkSheet.Cells[row_idx, col_idx];
							if (r.Value == null)
								continue;
							row[col_idx - 1] = r.Value;
						}
						table.AddRow(row);
					}
					list.Add(xlWorkSheet.Name, table);
				}
				return list;
			}
			catch (Exception ex)
			{
				this.Log.Trace(string.Format("Ошибка: {0}\r\nФайла:{1}", ex.Message, PathExcel));
				throw;
			}
			finally
			{
				try
				{
					if (wb != null)
					{
						wb.Close(false, Type.Missing, Type.Missing);
						wb = null;
					}
					if (excel != null)
					{
						excel.Quit();
						excel = null;
					}
					GC.Collect();
					GC.WaitForPendingFinalizers();
				}
				catch (Exception ex1)
				{
					this.Log.Trace(string.Format("Ошибка при закрытии : {0}\r\nФайла:{1}", ex1.Message, PathExcel));
				}
			}
		}

		public IEnumerable<string> GetExcelWorksheetName(string PathExcel)
		{
			Excel.Application excel = null;
			Excel.Workbook wb = null;

			try
			{
				List<string> s = new List<string>();
				excel = new Excel.Application();
				wb = excel.Workbooks.Open(PathExcel);
				foreach (Excel.Worksheet xlWorkSheet in wb.Worksheets)
					s.Add(xlWorkSheet.Name);
				return s;
			}
			catch (Exception ex)
			{
				this.Log.Trace(string.Format("Ошибка: {0}\r\nФайла:{1}", ex.Message, PathExcel));
				throw;
			}
			finally
			{
				try
				{
					if (wb != null)
					{
						wb.Close(false, Type.Missing, Type.Missing);
						wb = null;
					}
					if (excel != null)
					{
						excel.Quit();
						excel = null;
					}
					GC.Collect();
					GC.WaitForPendingFinalizers();
				}
				catch (Exception ex1)
				{
					this.Log.Trace(string.Format("Ошибка при закрытии : {0}\r\nФайла:{1}", ex1.Message, PathExcel));
				}
			}
		}

		public TableIner LoadExcel(string PathExcel, object idxOrNameSheet)
		{
			TableIner table = null;

			Excel.Application excel = null;
			Excel.Workbook wb = null;

			try
			{
				excel = new Excel.Application();
				wb = excel.Workbooks.Open(PathExcel);
				Excel.Worksheet xlWorkSheet = (Excel.Worksheet)wb.Worksheets.get_Item(idxOrNameSheet);
				List<string> ColumntList = new List<string>();
				for (int col_idx = 1; col_idx < xlWorkSheet.Columns.Count; col_idx++)
				{
					Excel.Range r = xlWorkSheet.Cells[1, col_idx];
					if (r.Value == null)
						break;
					ColumntList.Add(r.Text);
				}

				table = new TableIner(ColumntList.ToArray());

				for (int row_idx = 3; row_idx < xlWorkSheet.Rows.Count; row_idx++)
				{
					Excel.Range rStart = xlWorkSheet.Cells[row_idx, 2];
					if (rStart.Value == null)
						break;
					object[] row = new object[table.Column.Length];

					for (int col_idx = 1; col_idx <= ColumntList.Count; col_idx++)
					{
						Excel.Range r = xlWorkSheet.Cells[row_idx, col_idx];
						if (r.Value == null)
							continue;
						row[col_idx - 1] = r.Value;
					}
					table.AddRow(row);
				}
				return table;
			}
			catch (Exception ex)
			{
				this.Log.Trace(string.Format("Ошибка: {0}", ex.Message));
				throw;
			}
			finally
			{
				if (wb != null)
					wb.Close(false, Type.Missing, Type.Missing);
				if (excel != null)
					excel.Quit();
			}
		}

		public class TableIner
		{
			public List<object[]> Row { get; private set; }
			public string[] Column { get; private set; }

			public TableIner(string[] C)
			{
				Row = new List<object[]>();
				Column = C;
			}

			public void AddRow(object[] r)
			{
				Row.Add(r);
			}
		}
	}
}
