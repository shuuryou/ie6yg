using CefSharp;
using System;
using System.Collections.Generic;
using yomigaeri_shared;

namespace yomigaeri_backend.Browser.Handlers
{
	internal sealed class MyDialogHandler : CefSharp.Handler.DialogHandler
	{
		private readonly SynchronizerState m_SyncState;
		private readonly Action m_SyncProc;

		public IFileDialogCallback FileDialog_Callback { get; private set; }

		public MyDialogHandler(SynchronizerState syncState, Action syncProc)
		{
			m_SyncState = syncState ?? throw new ArgumentNullException("syncState");
			m_SyncProc = syncProc ?? throw new ArgumentNullException("syncProc");
		}

		protected override bool OnFileDialog(IWebBrowser chromiumWebBrowser, IBrowser browser, CefFileDialogMode mode, string title, string defaultFilePath, List<string> acceptFilters, IFileDialogCallback callback)
		{
			if (mode != CefFileDialogMode.Open)
			{ 	
				Logging.WriteLineToLog("MyDialogHandler: OnFileDialog: Mode {0} not supported.", mode);
				callback.Cancel();
				return true;
			}

			// TODO: Implement this somehow?
			Logging.WriteLineToLog("MyDialogHandler: OnFileDialog: LAME! Title should have been: {0}", title);
			Logging.WriteLineToLog("MyDialogHandler: OnFileDialog: LAME! Filter(s) should have been: {0}", string.Join(", ", acceptFilters));

			FileDialog_Callback = callback;

			m_SyncState.FileUploadPrompt = true;

			m_SyncProc.Invoke();

			return true;
		}

	}
}
