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
								i.Closed = item.Length <= 9 || item[9] == null? (DateTime?) null : ((DateTime)item[9]);
								i.Priority = item.Length <= 11 || item[11] == null? null : item[11].ToString();
								i.ВидРаботы = item.Length <= 13 || item[13] == null? null : item[13].ToString();
								i.Описание = item[7].ToString();
								i.Решение = item.Length <= 10 || item[10] == null ? null :  item[10].ToString();
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
				ret.IncendentList = datalist;
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
