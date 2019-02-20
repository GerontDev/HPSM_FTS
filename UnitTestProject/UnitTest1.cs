using HPSM_FTS;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace UnitTestProject
{
	[TestClass]
	public class UnitTest1
	{
		[TestMethod]
		public void Test_GeneratorTimeEnd()
		{
			var gen = new Generator(null);
			int minute_max = (int)(2.5 * 60);

			{
				var Opened = new System.DateTime(2019, 2, 18, 10, 10, 10);
				var genDate = gen.GeneratorTimeEnd(Opened, 15, minute_max);
				var s = genDate.Subtract(Opened);
				Assert.IsTrue(Opened.ToShortDateString() == genDate.ToShortDateString());
				Assert.IsTrue(s.TotalMinutes < minute_max);
				Assert.IsTrue(genDate.Hour < Generator.EndWorkHours);
				Assert.IsTrue(Opened < genDate);
			}

			{
				var Opened = new System.DateTime(2019, 2, 20, 17, 45, 10);
				var genDate = gen.GeneratorTimeEnd(Opened, 15, minute_max);
				var s = genDate.Subtract(Opened);

				Assert.IsTrue(Opened.ToShortDateString() == genDate.ToShortDateString());
				Assert.IsTrue(s.TotalMinutes < minute_max);
				Assert.IsTrue(genDate.Hour < Generator.EndWorkHours);
				Assert.IsTrue(Opened < genDate);
			}
		}

		[TestMethod]
		public void Test_GeneratorTimeBegin()
		{
			var gen = new Generator(null);
			int minute_max = (int)(2.5 * 60);

			{
				var Opened = new System.DateTime(2019, 2, 18, 9, 10, 10);
				var genDate = gen.GeneratorTimeBegin(Opened, 15, minute_max);
				var s = Opened.Subtract(genDate);
				Assert.IsTrue(Opened.ToShortDateString() == genDate.ToShortDateString());
				Assert.IsTrue(s.TotalMinutes < minute_max);
				Assert.IsTrue(genDate.Hour < Generator.EndWorkHours);
				Assert.IsTrue(Opened > genDate);
			}

			{
				var Opened = new System.DateTime(2019, 2, 20, 15, 45, 10);
				var genDate = gen.GeneratorTimeBegin(Opened, 15, minute_max);
				var s = Opened.Subtract(genDate);

				Assert.IsTrue(Opened.ToShortDateString() == genDate.ToShortDateString());
				Assert.IsTrue(s.TotalMinutes < minute_max);
				Assert.IsTrue(genDate.Hour < Generator.EndWorkHours);
				Assert.IsTrue(Opened > genDate);
			}
		}

		[TestMethod]
		public void Test_GeneratorDateTimeEnd()
		{
			var gen = new Generator(null);
			int minute_max = 24 * 60;

			{
				var Opened = new System.DateTime(2019, 2, 18, 10, 10, 10);
				var genDate = gen.GeneratorDateTimeEnd(Opened, 15, minute_max);
				var s = genDate.Subtract(Opened);
				Assert.IsTrue(s.TotalMinutes < minute_max);
				Assert.IsTrue(s.TotalMinutes > 15);
			}

			{
				var Opened = new System.DateTime(2019, 2, 20, 17, 45, 10);
				var genDate = gen.GeneratorDateTimeEnd(Opened, 15, minute_max);
				var s = genDate.Subtract(Opened);
				Assert.IsTrue(s.TotalMinutes < minute_max);
				Assert.IsTrue(s.TotalMinutes > 15);
			}
		}

		[TestMethod]
		public void Test_GeneratorIsValideClosedOpened_Высокий()
		{
			var gen = new Generator(null);

			Assert.IsTrue(gen.IsValideClosedOpened(
				new System.DateTime(2019, 2, 20, 16, 45, 10),
				new System.DateTime(2019, 2, 20, 17, 45, 9),
				EPriority.Высокий));


			Assert.IsTrue(gen.IsValideClosedOpened(
				new System.DateTime(2019, 2, 20, 16, 45, 10),
				new System.DateTime(2019, 2, 20, 17, 45, 9),
				EPriority.Высокий));

			Assert.IsFalse(gen.IsValideClosedOpened(
				new System.DateTime(2019, 2, 20, 15, 45, 10),
				new System.DateTime(2019, 2, 20, 17, 45, 9),
				EPriority.Высокий));

			Assert.IsFalse(gen.IsValideClosedOpened(
				new System.DateTime(2019, 2, 20, 12, 45, 10),
				new System.DateTime(2019, 2, 21, 12, 30, 10),
				EPriority.Высокий));
		}
		[TestMethod]
		public void Test_GeneratorIsValideClosedOpened_Важный()
		{
			var gen = new Generator(null);
			Assert.IsTrue(gen.IsValideClosedOpened(
				new System.DateTime(2019, 2, 20, 17, 45, 10),
				new System.DateTime(2019, 2, 20, 17, 45, 9),
				EPriority.Важный));

			Assert.IsTrue(gen.IsValideClosedOpened(
				new System.DateTime(2019, 2, 20, 15, 45, 10),
				new System.DateTime(2019, 2, 20, 17, 45, 10),
				EPriority.Важный));

			Assert.IsFalse(gen.IsValideClosedOpened(
				new System.DateTime(2019, 2, 20, 12, 45, 10),
				new System.DateTime(2019, 2, 20, 17, 45, 10),
				EPriority.Важный));

			Assert.IsFalse(gen.IsValideClosedOpened(
			new System.DateTime(2019, 2, 20, 12, 45, 10),
			new System.DateTime(2019, 2, 21, 12, 30, 10),
			EPriority.Важный));
		}

		[TestMethod]
		public void Test_GeneratorIsValideClosedOpened_Обычный()
		{
			var gen = new Generator(null);
			Assert.IsTrue(gen.IsValideClosedOpened(
				new System.DateTime(2019, 2, 20, 17, 45, 10),
				new System.DateTime(2019, 2, 20, 17, 45, 12),
				EPriority.Обычный));

			Assert.IsTrue(gen.IsValideClosedOpened(
				new System.DateTime(2019, 2, 20, 17, 45, 10),
				new System.DateTime(2019, 2, 20, 18, 0, 0),
				EPriority.Обычный));

			Assert.IsTrue(gen.IsValideClosedOpened(
				new System.DateTime(2019, 2, 22, 17, 45, 10),
				new System.DateTime(2019, 2, 25, 10, 0, 0),
				EPriority.Обычный));

			Assert.IsFalse(gen.IsValideClosedOpened(
				new System.DateTime(2019, 2, 22, 17, 45, 10),
				new System.DateTime(2019, 2, 26, 10, 0, 0),
				EPriority.Обычный));
		}
	
	}
}
