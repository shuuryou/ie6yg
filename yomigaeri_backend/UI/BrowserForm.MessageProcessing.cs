using CefSharp;
using System;
using System.Globalization;

namespace yomigaeri_backend.UI
{
	partial class BrowserForm
	{
		private void ProcessMessage(string message)
		{
			if (message == null)
				throw new ArgumentNullException("message");

			#region NAVIGATE
			if (message.StartsWith("NAVIGATE:", StringComparison.Ordinal))
			{
				Program.WebBrowser.LoadUrl(message.Substring(9));
				return;
			}
			#endregion

			#region WINSIZE
			if (message.StartsWith("WINSIZE", StringComparison.Ordinal))
			{
				string size = message.Trim().Substring(7);
				if (String.IsNullOrEmpty(size))
					goto winsize_err;

				int idx = size.IndexOf(',');

				if (idx == -1 || idx + 1 > size.Length)
					goto winsize_err;

				if (int.TryParse(size.Substring(0, idx), NumberStyles.None, CultureInfo.InvariantCulture, out int width) &&
					int.TryParse(size.Substring(idx + 1), NumberStyles.None, CultureInfo.InvariantCulture, out int height))
				{
					this.Width = width;
					this.Height = height;
					return;
				}

			winsize_err:
					Logging.WriteLineToLog("Illegal WINSIZE message receievd: \"{0}\"", message);

				return;
			}
			#endregion

			#region BTNBACK, BTNFORW
			if (message == "BTNBACK")
			{
				if (!Program.WebBrowser.CanGoBack)
					return;

				Program.WebBrowser.Back();

				return;
			}

			if (message == "BTNFORW")
			{
				if (!Program.WebBrowser.CanGoForward)
					return;

				Program.WebBrowser.Forward();
				return;
			}
			#endregion

			#region MNUBACK, MNUFORW
			if (message.StartsWith("MNUBACK") || message.StartsWith("MNUFORW"))
			{ 
				if (message.Length != 8)
					return;

				if (!int.TryParse(message[7].ToString(), NumberStyles.None, CultureInfo.InvariantCulture, out int offset))
					return;

				if (message.StartsWith("MNUBACK"))
					offset *= -1;

				Logging.WriteLineToLog("History menu. Offset={0}", offset);

				// TODO - This needs to be done another way at some point. Right now
				// CEF doesn't offer another way, but websites that fuck around with
				// the HTML5 history API will probably break this easily.
				Program.WebBrowser.ExecuteScriptAsync("window.history.go", offset);
			}
			#endregion
		}
	}
}
