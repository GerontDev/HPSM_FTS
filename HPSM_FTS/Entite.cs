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

		public string Login
		{
			get
			{
				var index = Applicant.IndexOf('(');
				if (index <= 0)
					return string.Empty;

				string shortname = (Applicant.Contains('(') ? Applicant.Substring(
					startIndex: index + 1,
					length: Applicant.Count() - index - 2 ) : "");
				return shortname;
			}
		}

		public string NameLoginOnly
		{
			get
			{
				string l = Login;

				if (l.IndexOf('@') < 0)
					return String.Empty;
				return l.Substring(0, l.IndexOf('@'));
			}
		}
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
		/// Приорететь
		/// </summary>
		public EPriority Priority { get; set; }
	}

	public enum EPriority
	{
		Обычный,
		Важный,
		Высокий,
		Требуется_вмешательство_более_квалифицированного_специалиста
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
		public List<Report2> Report2 { get; set; }
		public List<Report3> Report3 { get; set; }
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


	public class Report2
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
		/// Фаза
		/// </summary>
		public Phase Phase { get; set; }
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

        public string FIO
        {
            get
            {
                var index = Applicant.IndexOf('(');
                if (index <= 0)
                    return string.Empty;

                string fio = (Applicant.Contains('(') ?
                    Applicant.Substring(
                    startIndex: 0, length: index) : "");
                return fio;
            }
        }

        public string FIOandNameLoginOnly
        {
            get
            {
                var index = Applicant.IndexOf('(');
                if (index <= 0)
                    return string.Empty;

                return string.Format(@"{0}({1})", FIO, NameLoginOnly);
            }
        }
        public string Login
        {
            get
            {
                var index = Applicant.IndexOf('(');
                if (index <= 0)
                    return string.Empty;

                string shortname = (Applicant.Contains('(') ? Applicant.Substring(
                    startIndex: index + 1,
                    length: Applicant.Count() - index - 2) : "");
                return shortname;
            }
        }
        public string NameLoginOnly
        {
            get
            {
                string l = Login;

                if (l.IndexOf('@') < 0)
                    return String.Empty;
                return l.Substring(0, l.IndexOf('@'));
            }
        }
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
				return Closed == null ? "" : Closed.ToShortDateString();
			}
		}
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
		public EPriority Priority { get; set; }

		public string TextPriority
		{
			get
			{
				if (Priority ==  EPriority.Обычный)
				{
					return "Обычный (В течение 24 часов (в течение одних суток))";
				}
				else if (Priority == EPriority.Важный)
				{
					return "Важный (Не более 2,5 часов (в течение текущего рабочего дня))";
				}
				else if (Priority == EPriority.Высокий)
				{
					return "Высокий (Не более 1 часа (в течение текущего рабочего дня))";
				}
				else if (Priority == EPriority.Требуется_вмешательство_более_квалифицированного_специалиста)
				{
					return "Требуется вмешательство более квалифицированного специалиста (В течение 48 часов (в течение двух рабочих дней с момента установления приоритета))";
				}
				throw new Exception("Не указан");
			}
		}
		/// <summary>
		/// Представитель Заказчика
		/// </summary>
		public string CustomerRepresentative { get; set; }

		public string Контакт { get; set; }

		public string IPAndName { get; set; }
	}

	public class Contact
	{
		public string Name { get; set; }
		public string Phone { get; set; }
		public string Email { get; set; }
		public string Tite { get; set; }
		public string Address1 { get; set; }
		public string Address2 { get; set; }
		public string Dept_Name { get; set; }
	}

	public class NetName
	{
		public string PCName { get; set; }
		public string IP { get; set; }
		public string Text1 { get; set; }
		public string NameUser { get; set; }
		public string Date1 { get; set; }
		public string Time1 { get; set; }
	}

	public class DataMain
	{
		public Dictionary<string, Contact> Contactlist { get; set; }
		public List<Incendent> Incendent { get; set; }
		public Dictionary<string, NetName> NetNameList { get; set; }	
	}


	public class Report3
	{
		public string ViewWork { get; set; }
		public int IncendentCount { get; set; }
		public double Prochent { get; set; }
	}
}
