using System;
using System.Collections;

namespace yomigaeri_backend.Browser
{
	internal class SynchronizerState
	{
		private const int CHANGE_ITEMS = 17;

		private readonly BitArray m_Changes = new BitArray(CHANGE_ITEMS);

		#region Enums
		// ATTENTION:
		// Enums have to be shorts so they don't overflow the VB6 Integer type.

		public enum Change : short
		{
			Visible = 0,
			Toolbar = 1,
			StatusText = 2,
			StatusProgress = 3,
			Address = 4,
			PageTitle = 5,
			AddHistoryItem = 6,
			MenuBar = 7,
			TravelLog = 8,
			Cursor = 9,
			Tooltip = 10,
			SSLIcon = 11,
			CertificateState = 12,
			CertificateData = 13,
			CertificatePrompt = 14,
			JSDialogPrompt = 15,
			DownloadStart = 16,
		} // Remember to adjust the CHANGE_ITEMS constant!

		public enum SSLIconState : short
		{
			None = 0,
			Secure = 1,
			SecureBadCert = 2
		}

		[Flags]
		public enum FrontendCertificateStates : short
		{
			None = 0,
			OverallOK = 1 << 0,
			OverallBad = 1 << 1,
			UntrustedIssuer = 1 << 2,
			TrustedIssuer = 1 << 3,
			DateValid = 1 << 4,
			DateInvalid = 1 << 5,
			NameValid = 1 << 6,
			NameInvalid = 1 << 7,
			StrongCert = 1 << 8,
			WeakCert = 1 << 9,
			Revoked = 1 << 10,
			DateInvalidTooLong = 1 << 11,
			ChromeCTFail = 1 << 12
		}

		[Flags]
		public enum FrontendToolbarButtons : short
		{
			None = 0,
			Back = 1 << 0,
			Forward = 1 << 1,
			Stop = 1 << 2,
			Refresh = 1 << 3,
			Home = 1 << 4,
			Media = 1 << 5
		}

		[Flags]
		public enum FrontendMenuBarItems : short
		{
			None = 0,
			Stop = 1 << 0,
			Refresh = 1 << 1,
			Cut = 1 << 2,
			Copy = 1 << 3,
			Paste = 1 << 4
		}

		public enum FrontendJSDialogs : short
		{
			None = 0,
			Prompt = 1,
			Alert = 2,
			Confirm = 3,
			OnBeforeUnload = 4
		}
		#endregion

		#region FrontendJSDialogData Data Holder
		public class FrontendJSDialogData
		{
			public FrontendJSDialogData(FrontendJSDialogs type, string prompt, string defaultText = null)
			{
				Type = type;

				if (prompt == null)
					Prompt = string.Empty;
				else
					Prompt = prompt;

				if (defaultText != null)
					DefaultText = defaultText;
			}

			public FrontendJSDialogs Type { get; private set; }

			public string Prompt { get; private set; }

			public string DefaultText { get; private set; } // only for JS prompt() function.

			public override bool Equals(object obj)
			{
				if (obj == null)
					return false;


				if (!(obj is FrontendJSDialogData other))
					return false;

				return (other.Type == this.Type && other.Prompt == this.Prompt && other.DefaultText == this.DefaultText);
			}

			public override int GetHashCode()
			{
				return (Type, Prompt, DefaultText).GetHashCode();
			}
		}
		#endregion

		#region State Properties
		// ATTENTION:
		// If you add things, modify the Change enum and CHANGE_ITEMS constant.

		private bool m_Visible;
		public bool Visible
		{
			get { return m_Visible; }
			set
			{
				if (m_Visible == value)
					return;

				m_Visible = value;
				m_Changes[(int)Change.Visible] = true;
			}
		}

		private FrontendToolbarButtons m_ToolbarButtons;
		public FrontendToolbarButtons ToolbarButtons
		{
			get { return m_ToolbarButtons; }
			set
			{
				if (m_ToolbarButtons == value)
					return;

				m_ToolbarButtons = value;
				m_Changes[(int)Change.Toolbar] = true;
			}
		}

		private FrontendMenuBarItems m_MenuBarItems;
		public FrontendMenuBarItems MenuBarItems
		{
			get { return m_MenuBarItems; }
			set
			{
				if (m_MenuBarItems == value)
					return;

				m_MenuBarItems = value;
				m_Changes[(int)Change.MenuBar] = true;
			}
		}

		private string m_StatusText;
		public string StatusText
		{
			get { return m_StatusText; }
			set
			{
				if (m_StatusText == value)
					return;
				m_StatusText = value;
				m_Changes[(int)Change.StatusText] = true;
			}
		}

		private int m_StatusProgress;
		public int StatusProgress
		{
			get { return m_StatusProgress; }
			set
			{
				if (m_StatusProgress == value)
					return;
				m_StatusProgress = value;
				m_Changes[(int)Change.StatusProgress] = true;
			}
		}

		private string m_Address;
		public string Address
		{
			get { return m_Address; }
			set
			{
				if (m_Address == value)
					return;
				m_Address = value;
				m_Changes[(int)Change.Address] = true;
			}
		}

		private string m_PageTitle;
		public string PageTitle
		{
			get { return m_PageTitle; }
			set
			{
				if (m_PageTitle == value)
					return;
				m_PageTitle = value;
				m_Changes[(int)Change.PageTitle] = true;
			}
		}

		private bool m_AddHistoryItem;
		public bool AddHistoryItem
		{
			get { return m_AddHistoryItem; }
			set
			{
				if (m_AddHistoryItem == value)
					return;
				m_AddHistoryItem = value;
				m_Changes[(int)Change.AddHistoryItem] = true;
			}
		}

		private bool m_TravelLog;
		public bool TravelLog
		{
			get { return m_TravelLog; }
			set
			{
				if (m_TravelLog == value)
					return;
				m_TravelLog = value;
				m_Changes[(int)Change.TravelLog] = true;
			}
		}

		private int m_Cursor;
		public int Cursor
		{
			get { return m_Cursor; }
			set
			{
				if (m_Cursor == value)
					return;
				m_Cursor = value;
				m_Changes[(int)Change.Cursor] = true;
			}
		}

		private string m_Tooltip;
		public string Tooltip
		{
			get { return m_Tooltip; }
			set
			{
				if (m_Tooltip == value)
					return;

				m_Tooltip = value;
				m_Changes[(int)Change.Tooltip] = true;
			}
		}

		private SSLIconState m_SSLIconState;
		public SSLIconState SSLIcon
		{
			get { return m_SSLIconState; }
			set
			{
				if (m_SSLIconState == value)
					return;

				m_SSLIconState = value;
				m_Changes[(int)Change.SSLIcon] = true;
			}
		}

		private FrontendCertificateStates m_CertificateState;
		public FrontendCertificateStates CertificateState
		{
			get { return m_CertificateState; }
			set
			{
				if (m_CertificateState == value)
					return;

				m_CertificateState = value;
				m_Changes[(int)Change.CertificateState] = true;
			}
		}

		private byte[] m_CertificateData;
		public byte[] CertificateData
		{
			get { return m_CertificateData; }
			set
			{
				if (m_CertificateData == value)
					return;

				m_CertificateData = value;
				m_Changes[(int)Change.CertificateData] = true;
			}
		}

		private bool m_CertificatePrompt;
		public bool CertificatePrompt
		{
			get { return m_CertificatePrompt; }
			set
			{
				m_CertificatePrompt = value;
				m_Changes[(int)Change.CertificatePrompt] = value;
			}
		}
		private FrontendJSDialogData m_JSDialogPrompt;
		public FrontendJSDialogData JSDialogPrompt
		{
			get { return m_JSDialogPrompt; }
			set
			{
				if (m_JSDialogPrompt == value)
					return;

				m_JSDialogPrompt = value;

				m_Changes[(int)Change.JSDialogPrompt] = true;
			}
		}

		private string m_DownloadStart;

		public string DownloadStart
		{
			get { return m_DownloadStart; }
			set
			{
				if (m_DownloadStart == value)
					return;

				m_DownloadStart = value;

				m_Changes[(int)Change.DownloadStart] = true;
			}
		}
		#endregion

		#region Helper Methods
		public void SyncAll()
		{
			m_Changes.SetAll(true);
		}

		public void SyncNone()
		{
			m_Changes.SetAll(false);
		}

		public bool IsChanged(Change what)
		{
			return m_Changes[(int)what];
		}
		#endregion
	}
}