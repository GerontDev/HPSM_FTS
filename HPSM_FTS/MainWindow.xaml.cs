using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Reflection;
using System.IO;
using Microsoft.Win32;

namespace HPSM_FTS
{
	/// <summary>
	/// Логика взаимодействия для MainWindow.xaml
	/// </summary>
	public partial class MainWindow : Window
	{
		private NlogMemoryTarget _Target;
		private NLog.ILogger _logger;
		public List<CheckedMO> MoList = new List<CheckedMO>();

		public MainWindow()
		{
			InitializeComponent();

		}

		private void Window_Loaded(object sender, RoutedEventArgs e)
		{
			try
			{
				NLog.LogManager.LoadConfiguration("Nlog.Config");
				_Target = new NlogMemoryTarget("WindowLog", NLog.LogLevel.Trace, NLog.LogLevel.Error);
				_Target.Log += MessageLog;
				_logger = NLog.LogManager.GetLogger("WindowLog");
				this.Title += string.Format(" (Программа от {0})", GetLinkerTime(Assembly.GetExecutingAssembly()).ToString());
			}
			catch (Exception ex)
			{
				MessageBox.Show(this, ex.Message, "", MessageBoxButton.OK);
			}
		}

		private static string GetFullPath(string FileName)
		{
			return System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), FileName);
		}		

		public void EnableRun(bool value)
		{
			btnRun.Dispatcher.Invoke(() => { btnRun.IsEnabled = value; btnStop.IsEnabled = !value; });
		}

		public void MessageLog(string msg, Exception ex)
		{
			tbLog.Dispatcher.InvokeAsync(() =>
			{
				tbLog.Text += string.Format("\r\n{0}", msg);
				Exception exLoc = ex;
				while (exLoc != null)
				{
					tbLog.Text += string.Format("\r\nException message:{0}\r\n Exception StackTrace:{1}", exLoc.Message, exLoc.StackTrace);
					exLoc = exLoc.InnerException;
				}
			});
		}

		public void MessageLog(NLog.LogEventInfo msg)
		{
			if (msg.Level == NLog.LogLevel.Warn)
				MessageLog(string.Format("Внимание!: {0}", msg.Message), msg.Exception);

			else if (msg.Level == NLog.LogLevel.Error)
				MessageLog(string.Format("Ошибка!: {0}", msg.Message), msg.Exception);

			else if (msg.Level == NLog.LogLevel.Trace)
				MessageLog(msg.Message, msg.Exception);
		}

		public string GetLog()
		{
			return this.tbLog.Text;
		}

		private void btnRun_Click(object sender, RoutedEventArgs e)
		{
			try
			{
				OpenFileDialog openFileDialog = new OpenFileDialog();
				openFileDialog.Multiselect = false;
				openFileDialog.Filter = "Test (*.*)| *.xlsx";
				if (openFileDialog.ShowDialog() != true)
				{
					return;
				}

				tbFileName.Text = openFileDialog.FileName;


				if (!System.IO.File.Exists(tbFileName.Text))
				{
					MessageBox.Show(this, "Не указано файд");
					return;
				}

				var setting = new Generator.Setting()
				{
				};
				tbLog.Text = string.Empty;
				Task.Run(() =>
				{
					EnableRun(false);
					NLog.ILogger logger = null;
					string name_log = string.Empty;

					try
					{
						string text_data = DateTime.Now.ToString().Replace(".", "").Replace(":", "");
						name_log = string.Format("log_{0}.txt", text_data);
						string name_exel = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), string.Format("result_{0}.xlsx", text_data));

						logger = NLog.LogManager.GetLogger("WindowLog");
						var gan = new Generator(logger);
						var data = gan.LoadData(openFileDialog.FileName);
						var result = gan.Run(data, setting);
						logger.Trace("Сохранение файла");
						var exelsave = new ExcelUtillite(logger);
						System.Threading.Thread.Sleep(1000);

						exelsave.SaveExcel_Mo(data: result, NameFileExcel: name_exel);
						System.Diagnostics.Process.Start(name_exel);
						logger.Trace("Сохранение файла завершено");
						}
					catch (Exception ex)
					{
						if (logger != null)
							logger.Error(ex, "Ошибка к главной фукции");
					}
					finally
					{
						EnableRun(true);
					}
					try
					{
						System.IO.File.WriteAllText(name_log, this.tbLog.Dispatcher.Invoke<string>(this.GetLog));
					}
					catch (Exception ex)
					{
						if (logger != null)
							logger.Error(ex, "Ошибка при сохранение лога");
					}
				});
			}
			catch (Exception ex)
			{
				MessageBox.Show(this, string.Format("Ошибка {0}", ex.Message));
			}
		}

		public static DateTime GetLinkerTime(Assembly assembly, TimeZoneInfo target = null)
		{
			var filePath = assembly.Location;
			const int c_PeHeaderOffset = 60;
			const int c_LinkerTimestampOffset = 8;

			var buffer = new byte[2048];

			using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
				stream.Read(buffer, 0, 2048);

			var offset = BitConverter.ToInt32(buffer, c_PeHeaderOffset);
			var secondsSince1970 = BitConverter.ToInt32(buffer, offset + c_LinkerTimestampOffset);
			var epoch = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);

			var linkTimeUtc = epoch.AddSeconds(secondsSince1970);

			var tz = target ?? TimeZoneInfo.Local;
			var localTime = TimeZoneInfo.ConvertTimeFromUtc(linkTimeUtc, tz);

			return localTime;
		}
	}

	public class MOInfo
	{
		public int ID_number { get; set; }
		public string ID { get; set; }
		public string Name { get; set; }
		public string ExcelFile { get; set; }
		public string Title
		{
			get
			{
				return string.Format("{0} ({1})", Name, ExcelFile);
			}
		}
	}

	public class CheckedMO : MOInfo
	{
		public bool IsChecked { get; set; }
	};
}
