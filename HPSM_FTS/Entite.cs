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
		/// <summary>
		/// Рабочая  группа
		/// </summary>
		public string WorkGroup { get; set; }
		/// <summary>
		/// Заявитель
		/// </summary>
		public string Applicant { get; set; }
		/// <summary>
		/// 
		/// </summary>
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
		public DateTime? Closed { get; set; }

		public String ClosedDateString
		{
			get
			{
				return Closed == null ? "" : Closed.Value.ToShortDateString();
			}
		}

		/// <summary>
		/// ДС/АСВД/СКАД/ВШ/СКАД-Контроль
		/// </summary>
		public string Subsystem { get; set; }
		/// <summary>
		/// Виде работы
		/// </summary>
		public string ВидРаботы { get; set; }
		/// <summary>
		/// Описание (супть проблемы)
		/// </summary>
		public string Описание { get; set; }
		/// <summary>
		/// Проведение рабоы 
		/// </summary>
		public string Решение { get; set; }
		/// <summary>
		/// Приоритет
		/// </summary>
		public string Priority { get; set; }

		///// <summary>
		///// Категория работ
		///// </summary>
		//public string CategoryWork { get; set; }
		///// <summary>
		///// Название работ
		///// </summary>
		//public string NameWork { get; set; }
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
		public List<Incendent> IncendentList { get; set; }
	}


	public class Phase 
	{
		public Phase(int Number, DateTime DateBegin, DateTime DateEnd)
		{
			this.Number = Number;
			this.Begin = DateBegin;
			this.End = DateEnd;
		}		
		public int Number { get;}
		public DateTime Begin { get; }
		public DateTime End { get;  }

		public override string ToString()
		{
			return string.Format("Этапа {0} ({1}-{2})", Number, Begin.ToShortDateString(), End.ToShortDateString());
		}
	}
}
