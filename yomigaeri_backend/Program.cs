using CefSharp;
using CefSharp.WinForms;
using Microsoft.Win32;
using System;
using System.ComponentModel;
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

			SystemEvents.SessionSwitch += SystemEvents_SessionSwitch;

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

			#region Initialize AdBlock
			AdBlock.Initialize();

			// Some day there will be proper µBO grade blocking and a little
			// bit more to do here. :-(
			#endregion

			#region Initialize Chromium
			{
	
			}
			#endregion

		Application.Run(new BrowserForm());
			return 0;
        }

		private static void SystemEvents_SessionSwitch(object sender, SessionSwitchEventArgs e)
		{
			if (e.Reason != SessionSwitchReason.RemoteDisconnect &&
				e.Reason != SessionSwitchReason.ConsoleDisconnect &&
				e.Reason != SessionSwitchReason.SessionLogoff)
			{
				return;
			}

			// TODO: Shut down CEF

			Application.Exit();
		}
	}
}
