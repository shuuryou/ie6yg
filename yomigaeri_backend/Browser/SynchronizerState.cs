using System.Collections;

namespace yomigaeri_backend.Browser
{
	internal class SynchronizerState
	{
		private readonly BitArray m_Changes;
		public enum Change : int
		{
			Visible = 0,
			CanGoBack = 1,
			CanGoForward = 2,
			CanReload = 3,
			IsLoading = 4,
			StatusText = 5,
			StatusProgress = 6,
			Address = 7,
			PageTitle = 8,
			AddHistoryItem = 9,
			EnableTextCut = 10,
			EnableTextCopy = 11,
			EnableTextPaste = 12,
			TravelLog = 13,
			Cursor = 14,
			Tooltip = 15,
			SSLIcon = 16
		}

		private const int CHANGE_ITEMS = 17;

		public enum SSLIconState
		{
			None = 0,
			Secure = 1,
			SecureBadCert = 2
		}

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

		private bool m_CanGoBack;
		public bool CanGoBack
		{
			get { return m_CanGoBack; }
			set
			{
				if (m_CanGoBack == value)
					return;

				m_CanGoBack = value;
				m_Changes[(int)Change.CanGoBack] = true;
			}
		}

		private bool m_CanGoForward;
		public bool CanGoForward
		{
			get { return m_CanGoForward; }
			set
			{
				if (m_CanGoForward == value)
					return;

				m_CanGoForward = value;
				m_Changes[(int)Change.CanGoForward] = true;
			}
		}

		private bool m_CanReload;
		public bool CanReload
		{
			get { return m_CanReload; }
			set
			{
				if (m_CanReload == value)
					return;

				m_CanReload = value;
				m_Changes[(int)Change.CanReload] = true;
			}
		}

		private bool m_IsLoading;
		public bool IsLoading
		{
			get { return m_IsLoading; }
			set
			{
				if (m_IsLoading == value)
					return;
				m_IsLoading = value;
				m_Changes[(int)Change.IsLoading] = true;
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

		private bool m_EnableTextCut;
		public bool EnableTextCut
		{
			get { return m_EnableTextCut; }
			set
			{
				if (m_EnableTextCut == value)
					return;
				m_EnableTextCut = value;
				m_Changes[(int)Change.EnableTextCut] = true;
			}
		}

		private bool m_EnableTextCopy;
		public bool EnableTextCopy
		{
			get { return m_EnableTextCopy; }
			set
			{
				if (m_EnableTextCopy == value)
					return;
				m_EnableTextCopy = value;
				m_Changes[(int)Change.EnableTextCopy] = true;
			}
		}

		private bool m_EnableTextPaste;
		public bool EnableTextPaste
		{
			get { return m_EnableTextPaste; }
			set
			{
				if (m_EnableTextPaste == value)
					return;
				m_EnableTextPaste = value;
				m_Changes[(int)Change.EnableTextPaste] = true;
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
			get { return m_Tooltip;  }
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

		public SynchronizerState()
		{
			m_Changes = new BitArray(CHANGE_ITEMS);
		}

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
	}
}