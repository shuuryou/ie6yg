using CefSharp;
using System;

namespace yomigaeri_backend.Browser.Handlers
{
	internal class MyJsDialogHandler : CefSharp.Handler.JsDialogHandler
	{
		private readonly SynchronizerState m_SyncState;
		private readonly Action m_SyncProc;

		public IJsDialogCallback JSDialog_Callback { get; private set; }

		public MyJsDialogHandler(SynchronizerState syncState, Action syncProc)
		{
			m_SyncState = syncState ?? throw new ArgumentNullException("syncState");
			m_SyncProc = syncProc ?? throw new ArgumentNullException("syncProc");
		}

		protected override bool OnBeforeUnloadDialog(IWebBrowser chromiumWebBrowser, IBrowser browser, string messageText, bool isReload, IJsDialogCallback callback)
		{
			if (JSDialog_Callback != null && !JSDialog_Callback.IsDisposed)
			{
				JSDialog_Callback.Continue(false);
				JSDialog_Callback.Dispose();
			}

			JSDialog_Callback = callback;

			m_SyncState.JSDialogPrompt = new SynchronizerState.FrontendJSDialogData(
				SynchronizerState.FrontendJSDialogs.OnBeforeUnload, messageText);

			m_SyncProc.Invoke();

			return true;
		}

		protected override bool OnJSDialog(IWebBrowser chromiumWebBrowser, IBrowser browser, string originUrl, CefJsDialogType dialogType, string messageText, string defaultPromptText, IJsDialogCallback callback, ref bool suppressMessage)
		{
			SynchronizerState.FrontendJSDialogs frontendDialog;

			switch (dialogType)
			{
				case CefJsDialogType.Alert:
					frontendDialog = SynchronizerState.FrontendJSDialogs.Alert;
					break;
				case CefJsDialogType.Confirm:
					frontendDialog = SynchronizerState.FrontendJSDialogs.Confirm;
					break;
				case CefJsDialogType.Prompt:
					frontendDialog = SynchronizerState.FrontendJSDialogs.Prompt;
					break;
				default:
					throw new InvalidOperationException("Dialog type not supported by frontend.");
			}

			if (JSDialog_Callback != null && !JSDialog_Callback.IsDisposed)
			{
				JSDialog_Callback.Continue(false);
				JSDialog_Callback.Dispose();
			}

			JSDialog_Callback = callback;

			m_SyncState.JSDialogPrompt = new SynchronizerState.FrontendJSDialogData(
				frontendDialog, messageText, defaultPromptText);

			m_SyncProc.Invoke();

			return true;
		}

		protected override void OnResetDialogState(IWebBrowser chromiumWebBrowser, IBrowser browser)
		{
			JSDialog_Callback = null;
			m_SyncState.JSDialogPrompt = null;
		}
	}
}