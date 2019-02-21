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

		public DataMain LoadData(string filename, bool filterGroup)
		{
			this.Log.Trace("Загрузка данных");
			var Incendent_table = Task.Run(
					() =>
					{
						var excel = new ExcelUtillite(this.Log);
						Dictionary<string, ExcelUtillite.TableIner> table_list
							= excel.LoadExcelAllTable(
										PathExcel: filename,
										column_indexs_check_row: new int[] { 1, 2 },
										countemptyrow: 50,
										max_count_column: 15,
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
								i.Closed = item[9] == null ? (DateTime?)null : ((DateTime)item[9]);

								string sPriority = item[11] == null ? null : item[11].ToString();

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
								else if (string.IsNullOrEmpty(sPriority))
								{
									i.Priority = EPriority.Обычный;
								}
								else
									throw new Exception(string.Format("Неизвестный тип приоритета = \"{0}\",  ИНЦ = {1}", sPriority, i.ENC));

								i.ВидРаботы = item.Length <=13 || item[13] == null ? null : item[13].ToString();
								i.Описание = item[7].ToString();
								i.Решение = item[10] == null ? null : item[10].ToString();
								i.Subsystem = item[4].ToString();
								if (filterGroup && !(i.Subsystem == "ОДСИПЕАИСТОГПДС" ||
								   i.Subsystem == "АСВДТО"))
									continue;
								list.Add(i);
							}
						}
						this.Log.Trace(string.Format("Загружено инцидентов {0}", list.Count));
						return list;
					}
				);

			var DicNetName_table = Task.Run(
				() =>
				{
					var list = new Dictionary<string, NetName>(); 
					string FullName = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), Properties.Settings.Default.FileNameDicNetName);
					foreach (var itemRow in System.IO.File.ReadAllLines(FullName))
					{
						if (string.IsNullOrEmpty(itemRow))
							continue;
						string[] columnvalue = itemRow.Split(';');

						if (columnvalue.Length < 6)
							continue;
						string name = columnvalue[3].Trim();

						if (list.ContainsKey(name.ToLower()))
							continue;

						NetName i = new NetName();					
						i.PCName = columnvalue[0].Trim();
						i.IP = columnvalue[1].Trim();
						i.Text1 = columnvalue[2];
						i.NameUser = name;
						i.Date1 = columnvalue[4];
						i.Text1 = columnvalue[5];
						list.Add(i.NameUser.ToLower(), i);
					}
					this.Log.Trace(string.Format("Загружено сетевых устройств {0}", list.Count));
					return list;
				}
			);

			var DicContact_table = Task.Run(
				() =>
				{
					var excel = new ExcelUtillite(this.Log);
					Dictionary<string, ExcelUtillite.TableIner> table_list = excel.LoadExcelAllTable(
						PathExcel: System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), Properties.Settings.Default.FilaNameDicContact),
						column_indexs_check_row: new int[] { 1 },
						countemptyrow: 1,
						max_count_column: 8,
						WorksheetNames: new HashSet<string> { "Выгрузка контактов" });

					var list = new Dictionary<string, Contact>();

					foreach (var table_rasp_ip in table_list)
					{
						foreach (var item in table_rasp_ip.Value.Row)
						{
							Contact i = new Contact();
							i.Name = item[0].ToString().Trim();
							if (!item[1].ToString().Contains("NULL"))
								i.Phone = item[1].ToString().Trim();
							i.Email = item[2].ToString().Trim();
							i.Tite = item[3].ToString().Trim();
							if (!item[4].ToString().Contains("NULL"))
								i.Address1 = item[4].ToString().Trim();
							else
								i.Address1 = string.Empty;
							if (!item[5].ToString().Contains("NULL"))
								i.Address2 = item[5].ToString().Trim();
							else
								i.Address2 = string.Empty;
							i.Dept_Name = item[6].ToString().Trim();
							list.Add(i.Name.ToLower(), i);
						}
					}
					this.Log.Trace(string.Format("Загружено контактов {0}", list.Count));
					return list;
				});

			Task.WaitAll(Incendent_table, DicContact_table, DicNetName_table);
			return new DataMain()
			{
				Incendent = Incendent_table.GetAwaiter().GetResult(),
				Contactlist = DicContact_table.GetAwaiter().GetResult(),
				NetNameList = DicNetName_table.GetAwaiter().GetResult()
			};
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

		public DateTime GeneratorDateTimeBegin(DateTime Closed, int minutes_min, int minutes_max)
		{
            return Closed.AddMinutes(-27);

			//DateTime MaxDateTime = Closed.AddMinutes(minutes_max);
			//DateTime MinDateTime = Closed.AddMinutes(minutes_min);

			//if (MaxDateTime.Hour < BeginWorkHours) // Первод на предедушкий день
			//{
			//	DateTime DateTime_ = MaxDateTime.AddDays(-1);
			//	DateTime_ = new DateTime(MaxDateTime.Year, MaxDateTime.Month, MaxDateTime.Day, EndWorkHours, 0, 0);
			//	var gap = DateTime_.Subtract(MaxDateTime);
			//	minutes_max -= (int)gap.TotalMinutes;
			//	MaxDateTime = DateTime_;
			//}

			//if (MaxDateTime.Hour > EndWorkHours) // Первод конец дня
			//{
			//	DateTime DateTime_ = new DateTime(MaxDateTime.Year, MaxDateTime.Month, MaxDateTime.Day, EndWorkHours, 0, 0);
			//	var gap = MaxDateTime.Subtract(DateTime_);
			//	minutes_max -= (int)gap.TotalMinutes;
			//	MaxDateTime = DateTime_;
			//}

			//if (MinDateTime.Hour > EndWorkHours)
			//{
			//	DateTime DateTime_ = MaxDateTime.AddDays(1);
			//	DateTime_ = new DateTime(MaxDateTime.Year, MaxDateTime.Month, MaxDateTime.Day, EndWorkHours, 0, 0);

			//	var gap = MaxDateTime.Subtract(MaxDateTime);
			//	minutes_min -= (int)gap.TotalMinutes;
			//	MinDateTime = DateTime_;
			//}

			//DateTime dtRnd = begin;

			//do
			//{
			//	int minute = rnd.Next(minutes_min, minutes_max);
			//	dtRnd = begin.AddMinutes(minute);
			//}
			//while (dtRnd.Hour < BeginWorkHours || dtRnd.Hour > EndWorkHours);

			////если поподаеть выходные все переносим на понедельник
			//if (dtRnd.DayOfWeek >= DayOfWeek.Saturday)
			//	dtRnd = dtRnd.AddDays(2);

			//return Closed;
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
			else throw new Exception("Приоритет не указан");
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
			else throw new Exception("Приоритет не указан");
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

		public List<Report2> Report2(DataMain datalist)
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
			foreach (Incendent item in datalist.Incendent)
			{
				Report2 report2 = new Report2()
				{
					Applicant = item.Applicant,
					Opened = item.Opened,
					ENC = item.ENC,
					Number =i,
					ВидРаботы = item.ВидРаботы,
					Описание = item.Описание,
					Решение = item.Решение, 
					Priority = item.Priority

				};
				i++;
				report2.Phase = PhaseList.FirstOrDefault(q => q.Begin <= item.Opened && item.Opened <= q.End);


				if (!string.IsNullOrEmpty(item.Applicant))
				{
					if (datalist.Contactlist.ContainsKey(report2.Applicant.ToLower()))
					{
						var c = datalist.Contactlist[report2.Applicant.ToLower()];
						report2.Контакт = string.Format("{0}, {1}", c.Address2, c.Phone);
						if (report2.Контакт.StartsWith(","))
							report2.Контакт = report2.Контакт.Substring(1);
					}
					else
						Log.Warn(string.Format("Контакт для пользователя {0} не найден", report2.Applicant));

					string shortName = item.NameLoginOnly.ToLower();
					if (!string.IsNullOrEmpty(shortName) && datalist.NetNameList.ContainsKey(shortName))
					{
						var c = datalist.NetNameList[shortName];
						report2.IPAndName = string.Format("{0}/{1}", c.IP, c.PCName);
					}
					else
						Log.Warn(string.Format("Не найдено сетевое устройство для пользователя {0}", report2.Applicant));
				}


				if (item.Closed == null)
				{
					report2.Closed = GeneratorClosed(item.Opened, item.Priority);
					Log.Trace(string.Format("Инцидент {0} не закрыт", report2.ENC));
				}
				else if (!IsValideClosedOpened(item.Opened, item.Closed.Value, item.Priority))
				{
					Log.Trace(string.Format("Инцидент {0}: превышен норматив времени закрытия", report2.ENC));
					report2.Opened = GeneratorOpened(item.Closed.Value, item.Priority);
				}

				if (report2.Closed == DateTime.MinValue)
					report2.Closed = item.Closed.Value;


				report2.CustomerRepresentative = item.Subsystem == "АСВДТО" ? "Сутягин А.Н." : "Карпунина Т.Н";
				res.Add(report2);
			}
			return res;
		}
		
		public List<Report3> Report3(DataMain datalist)
		{
			List<Report3> list = new List<Report3>();
			
			foreach (var item in datalist.Incendent.GroupBy(q => q.ВидРаботы))
			{
				list.Add(new Report3()
				{
					ViewWork = item.Key,
					IncendentCount = item.Count(),
					Prochent = ((double)item.Count()) / datalist.Incendent.Count
				});
			}

			return list;
		}

		public Report1Result Report1(List<Report2> Incendents)
		{
			var list = new Report1Result();

			foreach (var item in Incendents.GroupBy(q => q.OpenedDateString))
			{
				string Date = item.Key;
				int C = item.Count();
				if (list.ContainsKey(Date))
					list[Date].OpenedCount += C;
				else
					list[Date] = new Report1Data() { OpenedCount = C };
			}

			foreach (var item in Incendents.GroupBy(q => q.ClosedDateString))
			{
				if (string.IsNullOrEmpty(item.Key))
					continue;
				string Date = item.Key;
				int C = item.Count();

				if (list.ContainsKey(Date))
					list[Date].ClosedCount += C;
				else
					list[Date] = new Report1Data() { ClosedCount = C };
			}
			return list;
		}

		public DataResult Run(DataMain datalist, Setting setting)
		{
			this.Log.Trace("Процесс обработки");

			try
			{
				DataResult ret = new DataResult();				
				ret.Report2 = Report2(datalist);				
				ret.Report1 = Report1(ret.Report2);
				ret.Report3 = Report3(datalist);
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
