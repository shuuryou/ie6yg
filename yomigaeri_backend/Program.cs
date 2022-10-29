﻿using CefSharp;
using CefSharp.WinForms;
using Microsoft.Win32;
using System;
using System.ComponentModel;
using System.Drawing;
using System.Globalization;
using System.IO;
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
			Application.SetCompatibleTextRenderingDefault(true);

			Application.ApplicationExit += Application_ApplicationExit;

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
			Logging.WriteLineToLog("This is a Terminal Server session? {0}", SystemInformation.TerminalServerSession);

			if (!SystemInformation.TerminalServerSession)
			{
				MessageBox.Show(Resources.Strings.E_InitErrorNotTerminalSession,
					Resources.Strings.E_InitErrorTitle, MessageBoxButtons.OK,
					MessageBoxIcon.None);

#if !DEBUG
				return 1;
#endif
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

				if (response == null)
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

				if (response == null)
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

			#region Request and Accept-Language List from Frontend
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

				if (response == null)
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

			#region Initialize AdBlock
			AdBlock.Initialize();

			// Some day there will be proper µBO grade blocking and a little
			// bit more to do here. :-(
			#endregion

			#region Initialize Chromium
			{
				CefSettings cefSettings = new CefSettings();

				if (!string.IsNullOrEmpty(frontendLanguageList))
					cefSettings.AcceptLanguageList = frontendLanguageList;

				cefSettings.BackgroundColor = Cef.ColorSetARGB(0, SystemColors.Window.R,
					SystemColors.Window.G, SystemColors.Window.B);

				cefSettings.CachePath =
					Environment.ExpandEnvironmentVariables(Settings.Get("CEF", "CachePath", string.Empty));

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
			}
			#endregion

			Application.Run(new BrowserForm());

			return 0;
		}

		private static void Application_ApplicationExit(object sender, EventArgs e)
		{

			if (Program.WebBrowser != null)
			{
				Program.WebBrowser.Dispose();
				Cef.Shutdown();
			}
		}

	}
}
