using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NLog;
using NLog.Targets;
using NLog.Common;
using NLog.Config;

namespace HPSM_FTS
{
	public class NlogMemoryTarget : Target
	{
		public Action<LogEventInfo> Log = delegate { };

		public NlogMemoryTarget(string name, LogLevel levelmin, LogLevel levelmax)
		{
			LogManager.Configuration.AddTarget(name, this);
			LogManager.Configuration.LoggingRules.Add(new LoggingRule("*", levelmin, levelmax, this));//This will ensure that exsiting rules are not overwritten
			LogManager.Configuration.Reload(); //This is important statement to reload all applied settings

			//SimpleConfigurator.ConfigureForTargetLogging (this, level); //use this if you are intending to use only NlogMemoryTarget  rule
		}

		protected override void Write(IList<AsyncLogEventInfo> logEvents)
		{
			foreach (var logEvent in logEvents)
			{
				Write(logEvent);
			}
		}

		protected override void Write(AsyncLogEventInfo logEvent)
		{
			Write(logEvent.LogEvent);
		}

		protected override void Write(LogEventInfo logEvent)
		{
			Log(logEvent);
		}
	}

}
