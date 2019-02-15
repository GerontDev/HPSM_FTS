using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HPSM_FTS
{
	public class Incendent
	{
		/// <summary>
		/// №п/п
		/// </summary>
		public int Number { get; set; }
		/// <summary>
		/// ИНЦ
		/// </summary>
		public string ENC { get; set; }
		/// <summary>
		/// открыто
		/// </summary>
		public DateTime Opened { get; set; }

		public String OpenedDateString
		{
			get
			{
				return Opened.ToShortDateString();
			}
		}
		/// <summary>
		/// закрыто
		/// </summary>
		public DateTime Closed { get; set; }

		public String ClosedDateString
		{
			get
			{
				return Closed.ToShortDateString();
			}
		}

		/// <summary>
		/// ДС/АСВД/СКАД/ВШ/СКАД-Контроль
		/// </summary>
		public string Subsystem { get; set; }
		/// <summary>
		/// Выполненная процедура
		/// </summary>
		public string WorkProccess { get; set; }
		/// <summary>
		/// Приоритет
		/// </summary>
		public string Priority { get; set; }
		/// <summary>
		/// Категория работ
		/// </summary>
		public string CategoryWork { get; set; }
		/// <summary>
		/// Название работ
		/// </summary>
		public string NameWork { get; set; }


	}

	public class Report1Data
	{
		public int OpenedCount { get; set; }
		public int ClosedCount { get; set; }
	}

	public class Report1Result : Dictionary<string, Report1Data>
	{

	}

	public class DataResult
	{
		public Report1Result Report1 { get; set; }
	}
}
