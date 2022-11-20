using CefSharp;
using System;
using System.Collections.Generic;
using System.IO;
using yomigaeri_shared;

namespace yomigaeri_backend.Browser.Handlers
{
	internal sealed class MyDownloadHandler : CefSharp.Handler.DownloadHandler
	{
		private readonly SynchronizerState m_SyncState;
		private readonly Action m_SyncProc;

		private readonly List<int> m_ActiveDownloads;

		public event EventHandler AllDownloadsCompleted;

		public MyDownloadHandler(SynchronizerState syncState, Action syncProc)
		{
			m_SyncState = syncState ?? throw new ArgumentNullException("syncState");
			m_SyncProc = syncProc ?? throw new ArgumentNullException("syncProc");

			m_ActiveDownloads = new List<int>();
		}

		public bool DownloadsInProgress { get { return m_ActiveDownloads.Count > 0; } }

		protected override void OnBeforeDownload(IWebBrowser chromiumWebBrowser, IBrowser browser, DownloadItem downloadItem, IBeforeDownloadCallback callback)
		{
			Logging.WriteLineToLog("MyDownloadHandler: Starting download for URL {0} (Type: {1}, Size: {2}, ID: {3})",
				downloadItem.Url, downloadItem.MimeType, downloadItem.TotalBytes, downloadItem.Id);

			string random_id;
			string filename;

		// TODO better error handling
		// but what can realistically fail other than writer.Save??

		again:
			random_id = Path.GetRandomFileName()
				.Remove(8, 1); // Get rid of '.' (can't replace it; strings are immutable in C-picketfence)

			filename = Path.GetFullPath(Path.Combine(Program.Settings.Get("Downloads", "DownloadTempDir"), random_id));

			string download_file = filename + ".bin";
			string download_meta = filename + ".ini";

			if (File.Exists(download_file) || File.Exists(download_meta))
				goto again;

			// YGSERVER looks for the INI file so it can relay the information to IE6.

			IniFileWriter writer = new IniFileWriter();
			writer.Add("Download", "CEFID", downloadItem.Id);
			writer.Add("Download", "TotalBytes", downloadItem.TotalBytes);
			writer.Add("Download", "MimeType", downloadItem.MimeType);
			writer.Add("Download", "ContentDisposition", downloadItem.ContentDisposition);
			writer.Add("Download", "URL", downloadItem.Url);
			writer.Add("Download", "SuggestedFileName", downloadItem.SuggestedFileName);
			writer.Add("Download", "LocalFile", download_file);
			writer.Save(download_meta);

			Logging.WriteLineToLog("MyDownloadHandler: Download ID {0}: File: {1}, Meta: {2} ",
				downloadItem.Id, download_file, download_meta);

			callback.Continue(download_file, false);

			AddDownload(downloadItem.Id);

			// The frontend will attempt to download this path from the backend server.

			string download_start;

			if (!string.IsNullOrEmpty(downloadItem.SuggestedFileName))
				download_start = random_id + '/' + downloadItem.SuggestedFileName;
			else
				download_start = random_id;

			m_SyncState.DownloadStart = download_start;
			m_SyncState.StatusProgress = 0;
			m_SyncProc.Invoke();

			Logging.WriteLineToLog("MyDownloadHandler: Notified frontend ({0}).", random_id);
		}

		protected override void OnDownloadUpdated(IWebBrowser chromiumWebBrowser, IBrowser browser, DownloadItem downloadItem, IDownloadItemCallback callback)
		{
			if (string.IsNullOrEmpty(downloadItem.FullPath))
			{
				// Download not yet confirmed, still floating in space; CEF will
				// often call OnDownloadUpdated before OnBeforeDownload. No idea
				// why. The authors of CEF/CEFsharp also don't seem to know.

				return;
			}

			Logging.WriteLineToLog("MyDownloadHandler: Download ID {0}: Progressing {1:n0}/{2:n0} bytes; {3}% complete. Still going? {4}",
				downloadItem.Id, downloadItem.ReceivedBytes, downloadItem.TotalBytes, downloadItem.PercentComplete, downloadItem.IsInProgress);


			// Doing it this way feels wrong but it saves calling the Path.*
			// functions, which basically do very similar things internally.
			string download_meta = downloadItem.FullPath.Substring(0, downloadItem.FullPath.Length - 3) + "ini";

			if (!File.Exists(download_meta))
			{
				// If the download_meta_file is removed while the download is
				// still in progress, it means that the server has cancelled
				// the download (the user aborted it, IE crashed, whatever).
				// This is just a poor man's IPC but dragging in named pipes,
				// or whatever other IPC-du-jour is currently the hottest shit,
				// just makes all of this so much more complicated. :-(

				Logging.WriteLineToLog("MyDownloadHandler: Download ID {0}: Meta {1} no longer exists. Canceling download.",
					downloadItem.Id, download_meta);

				callback.Cancel();

				// Chromium deletes the partially downloaded file.

				RemoveDownload(downloadItem.Id);

				return;
			}

			if (!downloadItem.IsInProgress)
			{
				// download_meta is guaranteed to exist due to above.

				// This is a way for the backend to signal to the server that a
				// download is finished and the file will not keep growing any
				// longer. This is important for downloads where the remote
				// server does not report the overall file size in the response
				// headers (files generated on the fly by a CGI script, etc.).

				File.Delete(download_meta);

				// The server should delete the download file, not us.

				RemoveDownload(downloadItem.Id);
			}
		}
		private void AddDownload(int id)
		{
			if (!m_ActiveDownloads.Contains(id))
				m_ActiveDownloads.Add(id);
		}

		private void RemoveDownload(int id)
		{
			if (m_ActiveDownloads.Contains(id))
				m_ActiveDownloads.Remove(id);

			if (m_ActiveDownloads.Count == 0)
				AllDownloadsCompleted?.Invoke(null, EventArgs.Empty);

		}
	}
}