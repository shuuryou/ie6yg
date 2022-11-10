using CefSharp;
using CefSharp.WinForms;
using System;
using System.ComponentModel;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Windows.Forms;

namespace yomigaeri_backend
{
	public static class Program
	{
		internal static FuckINI Settings { get; private set; }
		internal static ChromiumWebBrowser WebBrowser { get; private set; }
		

		[STAThread]
		public static int Main(string[] args)
		{
			#region Open Logger
			try
			{
				Logging.OpenLog();
			}
			catch (Exception e)
			{
				string err = string.Format(CultureInfo.CurrentUICulture,
					Resources.Strings.E_InitErrorCouldNotOpenLog, Logging.LogFile, e);

				MessageBox.Show(err, Resources.Strings.E_InitErrorTitle,
					MessageBoxButtons.OK, MessageBoxIcon.Error);

				return 1;
			}

			Logging.WriteBannerToLog("IE6 Yomigaeri's Back Orifice");
			#endregion

			Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
			Application.ThreadException += Application_ThreadException;

			Application.SetCompatibleTextRenderingDefault(false);

			Application.ApplicationExit += Application_ApplicationExit;

			#region Read INI File
			{
				string ini_file_location =
					Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "backend.ini");

				Logging.WriteLineToLog("Now going to load settings from \"{0}\".",
					ini_file_location);

				try
				{
					Settings = new FuckINI(ini_file_location);
				}
				catch (Exception e)
				{
					string err = string.Format(CultureInfo.CurrentUICulture,
						Resources.Strings.E_InitErrorCouldNotLoadSettings,
						ini_file_location, e.Message);

					Logging.WriteLineToLog("Error loading settings: {0}", e);

					MessageBox.Show(err, Resources.Strings.E_InitErrorTitle,
						MessageBoxButtons.OK, MessageBoxIcon.Error);


					return 1;
				}

				Logging.WriteLineToLog("Settings were loaded successfully.");
			}
			#endregion

			#region Initialize RDP Virtual Channel
			{
				Logging.WriteLineToLog("This is a Terminal Server session? {0}", SystemInformation.TerminalServerSession);

				if (!SystemInformation.TerminalServerSession)
				{
					MessageBox.Show(Resources.Strings.E_InitErrorNotTerminalSession,
						Resources.Strings.E_InitErrorTitle, MessageBoxButtons.OK,
						MessageBoxIcon.None);

					return 1;
				}

				Logging.WriteLineToLog("Opening RDP virtual channel to frontend.");

				try
				{
					RDPVirtualChannel.OpenChannel();
				}
				catch (Win32Exception e)
				{
					string err = string.Format(CultureInfo.CurrentUICulture,
						Resources.Strings.E_InitErrorCouldNotOpenRDPVC,
						 e.Message);

					Logging.WriteLineToLog("Error opening virtual channel: {0}", e);

					MessageBox.Show(err, Resources.Strings.E_InitErrorTitle,
						MessageBoxButtons.OK, MessageBoxIcon.Error);
#if !DEBUG
							return 1;
#endif
				}
			}
			#endregion

			#region Request and Apply Styling from Frontend
			{
				Logging.WriteLineToLog("Request styling from frontend.");

				RDPVirtualChannel.WriteChannel("STYLING");

				string response = null;

				try
				{
					response = RDPVirtualChannel.ReadChannelUntilResponse();
				}
				catch (TimeoutException)
				{
					Logging.WriteLineToLog("Time out getting styling from frontend.");
				}

				if (string.IsNullOrEmpty(response))
					goto skipStyling;

				Logging.WriteLineToLog("Frontend styling response is: \"{0}\".", response);

				if (response == "ERROR" || response == "UNSUPPORTED")
					goto skipStyling;

				Logging.WriteLineToLog("Apply frontend styling to this session.");

				try
				{
					FrontendStyling.ApplyStyling(response);
				}
				catch (Exception e)
				{
					Logging.WriteLineToLog("Error applying styling: {0}", e);
				}


			skipStyling:
				;
			}
			#endregion

			#region Request and Apply Cursors from Frontend
			{
				Logging.WriteLineToLog("Request cursors from frontend.");

				RDPVirtualChannel.WriteChannel("CURSORS");

				string response = null;

				try
				{
					response = RDPVirtualChannel.ReadChannelUntilResponse();
				}
				catch (TimeoutException)
				{
					Logging.WriteLineToLog("Time out getting curors from frontend.");
				}

				if (string.IsNullOrEmpty(response))
					goto skipStyling;

				Logging.WriteLineToLog("Frontend curors response is: \"{0}\".", response);

				if (response == "ERROR" || response == "UNSUPPORTED")
					goto skipStyling;

				Logging.WriteLineToLog("Apply frontend curors to this session.");

				try
				{
					FrontendStyling.ApplyCursors(response);
				}
				catch (Exception e)
				{
					Logging.WriteLineToLog("Error applying curors: {0}", e);
				}


			skipStyling:
				;
			}
			#endregion

			string frontendLanguageList = string.Empty;

			#region Request Accept-Language List from Frontend
			{
				Logging.WriteLineToLog("Request browser language list from frontend.");

				RDPVirtualChannel.WriteChannel("LANGLST");

				string response = null;

				try
				{
					response = RDPVirtualChannel.ReadChannelUntilResponse();
				}
				catch (TimeoutException)
				{
					Logging.WriteLineToLog("Time out getting browser language list from frontend.");
				}

				if (string.IsNullOrEmpty(response))
					goto skipLanguageList;

				Logging.WriteLineToLog("Frontend browser language list response is: \"{0}\".", response);

				if (response == "<EMPTY>")
					goto skipLanguageList;

				Logging.WriteLineToLog("Store frontend browser language list for CEF settings.");

				frontendLanguageList = response;

			skipLanguageList:
				;
			}
			#endregion

			int initial_width = 250, initial_height = 25;

			#region Request Initial Window Size from Frontend
			{
				RDPVirtualChannel.WriteChannel("INITSIZ");

				string response = null;

				try
				{
					response = RDPVirtualChannel.ReadChannelUntilResponse();
				}
				catch (TimeoutException)
				{
					Logging.WriteLineToLog("Time out getting initial window size from frontend.");
				}

				if (string.IsNullOrEmpty(response))
					goto skipWinSize;

				Logging.WriteLineToLog("Frontend initial window size response is: \"{0}\".", response);

				int idx = response.IndexOf(',');

				if (idx == -1 || idx + 1 > response.Length)
				{
					Logging.WriteLineToLog("Window size response is incorrect.");
					goto skipWinSize;
				}

				if (!int.TryParse(response.Substring(0, idx), NumberStyles.None, CultureInfo.InvariantCulture, out initial_width) ||
					!int.TryParse(response.Substring(idx + 1), NumberStyles.None, CultureInfo.InvariantCulture, out initial_height))
				{
					Logging.WriteLineToLog("Could not parseinitial  window size.");
				}

				Logging.WriteLineToLog("Parsed and stored initial window size.");

			skipWinSize:
				;
			}
			#endregion

			#region Initialize AdBlock
			try
			{
				// Some day there will be proper µBO grade blocking and a little
				// bit more to do here. :-(

				AdBlock.Initialize();
			}
			catch (Exception e)
			{
				string err = string.Format(CultureInfo.CurrentUICulture,
					Resources.Strings.E_InitErrorAdBlock,
					 e.Message);

				Logging.WriteLineToLog("Error initializing AdBlock: {0}", e);

				MessageBox.Show(err, Resources.Strings.E_InitErrorTitle,
					MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
			#endregion

			#region Initialize Chromium
			try
			{
				CefSettings cefSettings = new CefSettings();

				if (!string.IsNullOrEmpty(frontendLanguageList))
					cefSettings.AcceptLanguageList = frontendLanguageList;

				cefSettings.BackgroundColor = Cef.ColorSetARGB(0, SystemColors.Window.R,
					SystemColors.Window.G, SystemColors.Window.B);

				string cachePath = Environment.ExpandEnvironmentVariables(Settings.Get("CEF", "CachePath", string.Empty));

				if (!Directory.Exists(cachePath))
					Directory.CreateDirectory(cachePath);

				cefSettings.CachePath = cachePath;

				for (int i = 1; i < int.MaxValue; i++)
				{
					string arg = Settings.Get("CEF", string.Format(CultureInfo.InvariantCulture, "CommandLineArg{0}", i));
					if (string.IsNullOrWhiteSpace(arg))
						break;
					cefSettings.CefCommandLineArgs.Add(arg);
				}

				cefSettings.UserAgent = Settings.Get("CEF", "UserAgent", string.Empty);

				cefSettings.PersistUserPreferences = false;
				cefSettings.WindowlessRenderingEnabled = false;

				Cef.EnableWaitForBrowsersToClose();
				if (Settings.Get("CEF", "EnableHighDPISupport") == "1")
					Cef.EnableHighDPISupport();

				Cef.Initialize(cefSettings);

				Program.WebBrowser = new ChromiumWebBrowser("about:blank");
			}
			catch (Exception e)
			{
				string err = string.Format(CultureInfo.CurrentUICulture,
					Resources.Strings.E_InitErrorChromiumEmbeddedFramework,
					 e.Message);

				Logging.WriteLineToLog("Error initializing AdBlock: {0}", e);

				MessageBox.Show(err, Resources.Strings.E_InitErrorTitle,
					MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
			#endregion

			Application.Run(new UI.BrowserForm() { Width = initial_width, Height = initial_height });

			return 0;
		}

		private static void Application_ThreadException(object sender, ThreadExceptionEventArgs e)
		{
			Logging.WriteBannerToLog("CRASH");
			Logging.WriteLineToLog(e.Exception.ToString());
			Application.Exit();
		}

		private static void Application_ApplicationExit(object sender, EventArgs e)
		{
			Cef.Shutdown();
		}

	}
}