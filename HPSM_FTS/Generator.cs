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
										column_indexs_check_row: new int[] { 2, 5 },
										countemptyrow: 50,
										count_column: 9,
										WorksheetNames: new HashSet<string> { "Лист1" });

						List<Incendent> list = new List<Incendent>();

						foreach (var table_rasp_ip in table_list)
						{
							foreach (var item in table_rasp_ip.Value.Row)
							{
								Incendent i = new Incendent();
								i.Number = int.Parse(item[0].ToString());
								i.ENC = item[1].ToString();
								i.Opened = (DateTime)item[2];
								i.Closed = (DateTime)item[3];
								i.Subsystem = item[4].ToString();
								i.WorkProccess = item[5].ToString();
								i.Priority = item[6].ToString();
								i.CategoryWork = item[7].ToString();
								i.NameWork = item[8].ToString();
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
					string Date = item.Key;
					int C = item.Count();

					if (ret.Report1.ContainsKey(Date))
						ret.Report1[Date].ClosedCount += C;
					else
						ret.Report1[Date] = new Report1Data() { ClosedCount = C };
				}
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
