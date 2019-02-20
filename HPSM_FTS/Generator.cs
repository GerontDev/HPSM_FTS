using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NLog.Layouts;

namespace HPSM_FTS
{
	public class Generator
	{
		private readonly NLog.ILogger Log;

		public Generator(NLog.ILogger logger)
		{
			this.Log = logger;
		}

		public List<Incendent> LoadData(string filename)
		{
			this.Log.Trace("Загрузка данных");
			var task_IP_table = Task.Run(
					() =>
					{
						var excel = new ExcelUtillite(this.Log);
						Dictionary<string, ExcelUtillite.TableIner> table_list
							= excel.LoadExcelAllTable_only_Data(
										PathExcel: filename,
										column_indexs_check_row: new int[] { 1, 2 },
										countemptyrow: 50,
										count_column: 14,
										WorksheetNames: new HashSet<string> { "Лист1" });

						List<Incendent> list = new List<Incendent>();

						foreach (var table_rasp_ip in table_list)
						{
							foreach (var item in table_rasp_ip.Value.Row)
							{
								Incendent i = new Incendent();
								i.ENC = item[0].ToString();
								i.Opened = (DateTime)item[1];
								i.WorkGroup = item[4].ToString();
								i.Applicant = item[8].ToString();
								i.Closed = item[9] == null? (DateTime?) null : ((DateTime)item[9]);
								
								string sPriority = item[11] == null? null : item[11].ToString();

								if (sPriority == "Обычный")
								{
									i.Priority = EPriority.Обычный;
								}
								else if (sPriority == "Важный")
								{
									i.Priority = EPriority.Важный;
								}
								else if (sPriority == "Высокий")
								{
									i.Priority = EPriority.Высокий;
								}
								else if (sPriority == "Требуется вмешательство более квалифицированного специалиста")
								{
									i.Priority = EPriority.Требуется_вмешательство_более_квалифицированного_специалиста;
								}
								else
									throw new Exception(string.Format("Не извесный тип исключения \"{0}\"", sPriority));

								i.ВидРаботы =  item[13] == null? null : item[13].ToString();
								i.Описание = item[7].ToString();
								i.Решение = item[10] == null ? null :  item[10].ToString();
								i.Subsystem = item[4].ToString();
								list.Add(i);
							}
						}
						return list;
					}
				);
			
			Task.WaitAll(task_IP_table);
			return task_IP_table.GetAwaiter().GetResult();
		}		

		public class Setting
		{
			public bool GeneratorDistributionIPAddresses { get; set; }
		}

		private static Random rnd = new Random();

		public static int BeginWorkHours = 9;
		public static int EndWorkHours = 18;

		public DateTime GeneratorTimeEnd(DateTime begin, int minutes_min, int minutes_max)
		{
			var endday = new DateTime(begin.Year, begin.Month, begin.Day, EndWorkHours, 0, 0);
			var toend = endday.Subtract(begin);
			int minute_toend = (int)toend.TotalMinutes;
			int addminutes = (int)(rnd.Next((toend.TotalMinutes < minutes_min )? (int)toend.TotalMinutes : minutes_min, minute_toend > minutes_max ? minutes_max : minute_toend));
			return begin.AddMinutes(addminutes);
		}

		public DateTime GeneratorTimeBegin(DateTime end, int minutes_min, int minutes_max)
		{
			var beginday = new DateTime(end.Year, end.Month, end.Day, BeginWorkHours, 0, 0);
			var tobegin = end.Subtract(beginday);
			int minute_tobegin = (int)tobegin.TotalMinutes;
			int addminutes = (int)(rnd.Next((minute_tobegin < minutes_min) ? (int)minute_tobegin : minutes_min, minute_tobegin > minutes_max ? minutes_max : minute_tobegin));
			return end.AddMinutes(-addminutes);
		}

		public DateTime GeneratorDateTimeEnd(DateTime begin, int minutes_min, int minutes_max)
		{
			DateTime MaxDateTime = begin.AddMinutes(minutes_max);
			DateTime MinDateTime = begin.AddMinutes(minutes_min);

			if (MaxDateTime.Hour < BeginWorkHours) // Первод на предедушкий день
			{
				DateTime DateTime_ = MaxDateTime.AddDays(-1);
				DateTime_ = new DateTime(MaxDateTime.Year, MaxDateTime.Month, MaxDateTime.Day, EndWorkHours, 0, 0);
				var gap = DateTime_.Subtract(MaxDateTime);
				minutes_max -= (int)gap.TotalMinutes;
				MaxDateTime = DateTime_;
			}

			if (MaxDateTime.Hour > EndWorkHours) // Первод конец дня
			{
				DateTime DateTime_ = new DateTime(MaxDateTime.Year, MaxDateTime.Month, MaxDateTime.Day, EndWorkHours, 0, 0);
				var gap = MaxDateTime.Subtract(DateTime_);
				minutes_max -= (int)gap.TotalMinutes;
				MaxDateTime = DateTime_;
			}

			if (MinDateTime.Hour > EndWorkHours)
			{
				DateTime DateTime_ = MaxDateTime.AddDays(1);
				DateTime_ = new DateTime(MaxDateTime.Year, MaxDateTime.Month, MaxDateTime.Day, EndWorkHours, 0, 0);

				var gap = MaxDateTime.Subtract(MaxDateTime);
				minutes_min -= (int)gap.TotalMinutes;
				MinDateTime = DateTime_;
			}

			DateTime dtRnd = begin;

			do
			{
				int minute = rnd.Next(minutes_min, minutes_max);
				dtRnd = begin.AddMinutes(minute);
			}
			while (dtRnd.Hour < BeginWorkHours || dtRnd.Hour > EndWorkHours);
			
			//если поподаеть выходные все переносим на понедельник
			if (dtRnd.DayOfWeek >= DayOfWeek.Saturday)
				dtRnd = dtRnd.AddDays(2);

			return dtRnd;
		}

		public DateTime GeneratorClosed(DateTime Opened, EPriority priority)
		{
			if (priority == EPriority.Обычный)
			{
				return GeneratorDateTimeEnd(Opened, 1 * 60, 24 * 60);
			}
			else if (priority == EPriority.Важный)
			{
				return GeneratorTimeEnd(Opened, 15, (int)(2.5 * 60));
			}
			else if (priority == EPriority.Высокий)
			{
				return GeneratorTimeEnd(Opened, 15, 60);
			}
			else if (priority == EPriority.Требуется_вмешательство_более_квалифицированного_специалиста)
			{
				return GeneratorDateTimeEnd(Opened, 24 * 60, 48 * 60);
			}
			else throw new Exception("Приорететь не указан");
		}

		public DateTime GeneratorOpened(DateTime Closed, EPriority priority)
		{
			if (priority == EPriority.Обычный)
			{
				return GeneratorDateTimeBegin(Closed, 1 * 60, 24 * 60);
			}
			else if (priority == EPriority.Важный)
			{
				return GeneratorTimeBegin(Closed, 15, (int)(2.5 * 60));
			}
			else if (priority == EPriority.Высокий)
			{
				return GeneratorTimeBegin(Closed, 15, 60);
			}
			else if (priority == EPriority.Требуется_вмешательство_более_квалифицированного_специалиста)
			{
				return GeneratorDateTimeEnd(Closed, 24 * 60, 48 * 60);
			}
			else throw new Exception("Приорететь не указан");
		}

		public bool IsValideClosedOpened(DateTime Opened, DateTime Closed, EPriority priority)
		{
			if (priority == EPriority.Обычный)
			{
				int max = (int)(24 * 60);
				if (Opened.DayOfWeek == DayOfWeek.Friday)
					max += 24 * 2 * 60;
				return Closed.Subtract(Opened).TotalMinutes < max;
			}
			else if (priority == EPriority.Важный)
			{				
				int max = (int)(2.5 * 60);
				return Closed.Subtract(Opened).TotalMinutes < max;
			}
			else if (priority == EPriority.Высокий)
			{
				int max = 60;
				return Closed.Subtract(Opened).TotalMinutes < max;
			}
			else if (priority == EPriority.Требуется_вмешательство_более_квалифицированного_специалиста)
			{
				int max = (int)(48 * 60);
				if (Opened.DayOfWeek == DayOfWeek.Friday)
					max += 24 * 2 * 60;
				return Closed.Subtract(Opened).TotalMinutes < max;
			}
			return false;
		}

		public List<Report2> Report2(List<Incendent> datalist)
		{
			List<Report2> res = new List<Report2>();

			List<Phase> PhaseList = new List<Phase>();
			PhaseList.Add(new Phase(1, new DateTime(2018, 12, 28), new DateTime(2019, 2, 28)));
			PhaseList.Add(new Phase(2, new DateTime(2019, 3, 01), new DateTime(2019, 5, 31)));
			PhaseList.Add(new Phase(3, new DateTime(2019, 6, 01), new DateTime(2019, 8, 31)));
			PhaseList.Add(new Phase(4, new DateTime(2019, 9, 01), new DateTime(2019, 12, 10)));
			PhaseList.Add(new Phase(5, new DateTime(2019, 12, 11), new DateTime(2020, 2, 29)));
			PhaseList.Add(new Phase(6, new DateTime(2020, 3, 01), new DateTime(2020, 5, 31)));
			PhaseList.Add(new Phase(7, new DateTime(2020, 6, 01), new DateTime(2020, 8, 31)));
			PhaseList.Add(new Phase(8, new DateTime(2020, 9, 01), new DateTime(2020, 12, 10)));

			int i = 1;
			foreach (Incendent item in datalist)
			{
				Report2 report2 = new Report2()
				{
					Applicant = item.Applicant,
					Opened = item.Opened,
					ENC = item.ENC,
					Number =i,
					ВидРаботы = item.ВидРаботы,
					Описание = item.Описание,
					Решение = item.Решение
				};
				i++;
				report2.Phase = PhaseList.FirstOrDefault(q => q.Begin <= item.Opened && item.Opened <= q.End);

				if (item.Closed == null)
				{
					item.Closed = GeneratorClosed(item.Opened, item.Priority);
				}
				else if (!IsValideClosedOpened(item.Opened, item.Closed.Value, item.Priority))
				{
					item.Opened = GeneratorOpened(item.Closed.Value, item.Priority);
				}

				report2.CustomerRepresentative = item.Subsystem == "АСВДТО" ? "Сутягин А.Н." : "Карпунина Т.Н";
				res.Add(report2);
			}
			return res;
		}

		public DataResult Run(List<Incendent> datalist, Setting setting)
		{
			this.Log.Trace("Процесс обработки");

			try
			{
				DataResult ret = new DataResult();
				ret.Report1 = new Report1Result();

				foreach (var item in datalist.GroupBy(q => q.OpenedDateString))
				{
					string Date = item.Key;
					int C = item.Count();
					if (ret.Report1.ContainsKey(Date))
						ret.Report1[Date].OpenedCount += C;
					else
						ret.Report1[Date] = new Report1Data() { OpenedCount = C };
				}

				foreach (var item in datalist.GroupBy(q => q.ClosedDateString))
				{
					if (string.IsNullOrEmpty(item.Key))
						continue;
					string Date = item.Key;
					int C = item.Count();

					if (ret.Report1.ContainsKey(Date))
						ret.Report1[Date].ClosedCount += C;
					else
						ret.Report1[Date] = new Report1Data() { ClosedCount = C };
				}
				ret.Report2 = Report2(datalist);
				this.Log.Trace("Процесс заверщен");
				return ret;
			}
			catch (Exception ex)
			{
				this.Log.Trace(string.Format("Ошибка: {0}", ex.Message));
				throw;
			}
		}
	}
}
