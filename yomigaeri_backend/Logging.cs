using System;
using System.Globalization;
using System.IO;
using System.Text;

namespace yomigaeri_backend
{
	internal static class Logging
	{
		private static StreamWriter s_LogWriter;

		public static string LogFile { get; private set; }

		public static void OpenLog()
		{
			if (s_LogWriter == null)
			{
				LogFile = TempName("YBEDEBUG");
				s_LogWriter = new StreamWriter(LogFile, true, Encoding.UTF8);
			}
		}

		public static void CloseLog(bool dontDelete = false)
		{
			if (s_LogWriter == null)
				return;

			WriteLineToLog("Closing log file.");
			s_LogWriter.Close();

			if (!dontDelete)
				try
				{
					File.Delete(LogFile);
				}
				catch (IOException) { }
				catch (UnauthorizedAccessException) { }
		}

		public static void WriteLineToLog(string value)
		{
			if (s_LogWriter == null)
				return;

			lock (s_LogWriter)
			{
				s_LogWriter.WriteLine("{0}\t{1}", DateTime.Now, value);
				s_LogWriter.Flush();
			}
		}

		public static void WriteBannerToLog(string name)
		{
			WriteLineToLog("------------- {0} -------------", name);
		}

		public static void WriteLineToLog(string format, params object[] args)
		{
			WriteLineToLog(string.Format(CultureInfo.InvariantCulture, format, args));
		}

		private static string TempName(string prefix)
		{
			// Temporary file name format "%TEMP%\{prefix}_{timestamp}.txt".

			string ret = string.Format(CultureInfo.InvariantCulture,
				"{0}{1}{2}_{3}.txt", Path.GetTempPath(), Path.DirectorySeparatorChar, prefix,
				Math.Floor((DateTime.UtcNow - new DateTime(1970, 1, 1, 0, 0, 0)).TotalSeconds));

			ret = Path.GetFullPath(ret); // Normalize

			return ret;
		}
	}
}