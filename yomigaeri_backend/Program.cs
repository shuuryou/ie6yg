using CefSharp;
using CefSharp.WinForms;
using System;
using System.ComponentModel;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using yomigaeri_shared;

namespace yomigaeri_backend
{
	public static class Program
	{
		internal static IniFileReader Settings { get; private set; }
		internal static ChromiumWebBrowser WebBrowser { get; private set; }
		

		[STAThread]
		public static int Main(string[] args)
		{
			#region Open Logger
			try
			{
				Logging.OpenLog("YGBACKEND");
			}
			catch (Exception e)
			{
				string err = string.Format(CultureInfo.CurrentUICulture,
					Resources.E_InitErrorCouldNotOpenLog, Logging.LogFile, e);

				MessageBox.Show(err, Resources.E_InitErrorTitle,
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
					Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config.ini");

				Logging.WriteLineToLog("Main: Now going to load settings from \"{0}\".",
					ini_file_location);

				try
				{
					Settings = new IniFileReader(ini_file_location);
				}
				catch (Exception e)
				{
					string err = string.Format(CultureInfo.CurrentUICulture,
						Resources.E_InitErrorCouldNotLoadSettings,
						ini_file_location, e.Message);

					Logging.WriteLineToLog("Main: Error loading settings: {0}", e);

					MessageBox.Show(err, Resources.E_InitErrorTitle,
						MessageBoxButtons.OK, MessageBoxIcon.Error);

					return 1;
				}

				Logging.WriteLineToLog("Main: Settings were loaded successfully.");
			}
			#endregion

			#region Initialize RDP Virtual Channel
			{
				Logging.WriteLineToLog("Main: Is this a Terminal Server session? {0}", SystemInformation.TerminalServerSession);

				if (!SystemInformation.TerminalServerSession)
				{
					MessageBox.Show(Resources.E_InitErrorNotTerminalSession,
						Resources.E_InitErrorTitle, MessageBoxButtons.OK,
						MessageBoxIcon.None);

					return 1;
				}

				Logging.WriteLineToLog("Main: Opening RDP virtual channel to frontend.");

				try
				{
					RDPVirtualChannel.Open();
				}
				catch (Win32Exception e)
				{
					string err = string.Format(CultureInfo.CurrentUICulture,
						Resources.E_InitErrorCouldNotOpenRDPVC,
						 e.Message);

					Logging.WriteLineToLog("Main: Error opening virtual channel: {0}", e);

					MessageBox.Show(err, Resources.E_InitErrorTitle,
						MessageBoxButtons.OK, MessageBoxIcon.Error);
#if !DEBUG
							return 1;
#endif
				}
			}
			#endregion

			#region Request and Apply Styling from Frontend
			{
				Logging.WriteLineToLog("Main: Request styling from frontend.");

				RDPVirtualChannel.Write("GETINFO STYLING");

				string response = null;

				try
				{
					response = RDPVirtualChannel.ReadUntilResponse();
				}
				catch (TimeoutException)
				{
					Logging.WriteLineToLog("Main: Time out getting styling from frontend.");
				}

				if (string.IsNullOrEmpty(response))
					goto skipStyling;

				Logging.WriteLineToLog("Main: Frontend styling response is: \"{0}\".", response);

				if (response == "ERROR" || response == "UNSUPPORTED")
					goto skipStyling;

				Logging.WriteLineToLog("Main: Apply frontend styling to this session.");

				try
				{
					FrontendStyling.ApplyStyling(response);
				}
				catch (Exception e)
				{
					Logging.WriteLineToLog("Main: Error applying styling: {0}", e);
				}


			skipStyling:
				;
			}
			#endregion

			#region Request and Apply Cursors from Frontend
			{
				Logging.WriteLineToLog("Request cursors from frontend.");

				RDPVirtualChannel.Write("CURSORS");

				string response = null;

				try
				{
					response = RDPVirtualChannel.ReadUntilResponse();
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
				Logging.WriteLineToLog("Main: Request browser language list from frontend.");

				RDPVirtualChannel.Write("GETINFO LANGUAGES");

				string response = null;

				try
				{
					response = RDPVirtualChannel.ReadUntilResponse();
				}
				catch (TimeoutException)
				{
					Logging.WriteLineToLog("Main: Time out getting browser language list from frontend.");
				}

				if (string.IsNullOrEmpty(response))
					goto skipLanguageList;

				Logging.WriteLineToLog("Main: Frontend browser language list response is: \"{0}\".", response);

				if (response == "<EMPTY>")
					goto skipLanguageList;

				Logging.WriteLineToLog("Main: Store frontend browser language list for CEF settings.");

				frontendLanguageList = response;

			skipLanguageList:
				;
			}
			#endregion

			int initial_width = 250, initial_height = 250;

			#region Request Initial Window Size from Frontend
			{
				RDPVirtualChannel.Write("GETINFO INITIALSIZE");

				string response = null;

				try
				{
					response = RDPVirtualChannel.ReadUntilResponse();
				}
				catch (TimeoutException)
				{
					Logging.WriteLineToLog("Main: Time out getting initial window size from frontend.");
				}

				if (string.IsNullOrEmpty(response))
					goto skipWinSize;

				Logging.WriteLineToLog("Main: Frontend initial window size response is: \"{0}\".", response);

				int idx = response.IndexOf(',');

				if (idx == -1 || idx + 1 > response.Length)
				{
					Logging.WriteLineToLog("Main: Window size response is incorrect.");
					goto skipWinSize;
				}

				if (!int.TryParse(response.Substring(0, idx), NumberStyles.None, CultureInfo.InvariantCulture, out initial_width) ||
					!int.TryParse(response.Substring(idx + 1), NumberStyles.None, CultureInfo.InvariantCulture, out initial_height))
				{
					Logging.WriteLineToLog("Main: Could not parse initial window size.");
				}

				Logging.WriteLineToLog("Main: Parsed and stored initial window size.");

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
					Resources.E_InitErrorAdBlock,
					 e.Message);

				Logging.WriteLineToLog("Main: Error initializing AdBlock: {0}", e);

				MessageBox.Show(err, Resources.E_InitErrorTitle,
					MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
			#endregion

			string download_temp_directory = null;

			#region Create Download Temp Dir
			try
			{
				download_temp_directory = Settings.Get("Downloads", "DownloadTempDir");

				if (!Directory.Exists(download_temp_directory))
					Directory.CreateDirectory(download_temp_directory);

			} catch (Exception e)
			{
				string err = string.Format(CultureInfo.CurrentUICulture,
					Resources.E_InitErrorDownloadTempDir,
					download_temp_directory, e.Message);

				Logging.WriteLineToLog("Main: Error creating temporary directory \"{0}\" for downloads: {1}", 
					download_temp_directory, e);

				MessageBox.Show(err, Resources.E_InitErrorTitle,
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
					Resources.E_InitErrorChromiumEmbeddedFramework,
					 e.Message);

				Logging.WriteLineToLog("Main: Error initializing CEF: {0}", e);

				MessageBox.Show(err, Resources.E_InitErrorTitle,
					MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
			#endregion

			Application.Run(new Browser.BrowserForm() { Width = initial_width, Height = initial_height });

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