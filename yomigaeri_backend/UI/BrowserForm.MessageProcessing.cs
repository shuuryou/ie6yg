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

			// There is no technical reason that the messages start with
			// seven letter long commands; they could be longer. It just
			// emerged as a pattern and now it's here to stay like that.

			#region NAVIGAT -- Navigate to URL
			if (message.StartsWith("NAVIGAT", StringComparison.Ordinal))
			{
				string url = message.Substring(7);

				if (string.IsNullOrEmpty(url))
					return;

				Program.WebBrowser.LoadUrl(url);
				return;
			}
			#endregion

			#region WINSIZE -- Adjust window size
			if (message.StartsWith("WINSIZE", StringComparison.Ordinal))
			{
				// e.g. WINSIZE800,600
				//      01234567890123...

				string size = message.Substring(7);

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

			#region BTNBACK, BTNFORW -- Forward and Back buttons
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

			#region MNUBACK, MNUFORW -- Forward and Back travel log
			if (message.StartsWith("MNUBACK") || message.StartsWith("MNUFORW"))
			{ 
				if (message.Length != 8)
					return;

				// e.g. MNUBACK1
				//      01234567

				if (!int.TryParse(message[7].ToString(), NumberStyles.None, CultureInfo.InvariantCulture, out int offset))
					return;

				if (message.StartsWith("MNUBACK"))
					offset *= -1;

				// TODO - This needs to be done another way at some point. Right now
				// CEF doesn't offer another way, but websites that fuck around with
				// the HTML5 history API will probably break this easily.
				Program.WebBrowser.ExecuteScriptAsync("window.history.go", offset);
			}
			#endregion
		}
	}
}
