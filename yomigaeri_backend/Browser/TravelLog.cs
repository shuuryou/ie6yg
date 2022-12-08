using CefSharp;
using System;
using System.Collections.Generic;
using yomigaeri_shared;

namespace yomigaeri_backend.Browser
{
	internal class TravelLog : INavigationEntryVisitor
	{
		public event EventHandler NewLogPrepared;

		private readonly List<NavigationEntry> m_NavigationEntries;
		private int m_offsetCurrent;

		// @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
		// @ DO NOT CHANGE THIS WITHOUT:                                  @
		// @ 1. Adapting the "parser" in BrowserForm.MessageProcessing.cs @
		// @ 2. Adapting code in the frontend to deal with the new amount @
		// @ 3. Modifying the menu of the frontend using VB6 Menu Editor  @
		// @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
		private const int MAX_ENTRIES = 11;

		public enum TravelDirection
		{
			Back = 1,
			Forward = 2
		}

		public TravelLog()
		{
			m_NavigationEntries = new List<NavigationEntry>(MAX_ENTRIES);
			m_offsetCurrent = -1;
		}

		public void PrepareForVisit()
		{
			m_NavigationEntries.Clear();
			m_offsetCurrent = -1;
		}

		public bool Visit(NavigationEntry entry, bool current, int index, int total)
		{
			m_NavigationEntries.Add(entry);

			if (m_NavigationEntries.Count > MAX_ENTRIES)
				m_NavigationEntries.RemoveAt(0);

			if (m_offsetCurrent == MAX_ENTRIES || index == total - 1)
			{
				for (int i = 0; i < m_NavigationEntries.Count; i++)
					if (m_NavigationEntries[i].IsCurrent)
					{
						m_offsetCurrent = i;
						break;
					}

				NewLogPrepared?.Invoke(null, EventArgs.Empty);
				return false;
			}

			return true;
		}

		public string MakeMenuStringForFrontend(TravelDirection button)
		{
			try
			{
				NavigationEntry[] entries = m_NavigationEntries.ToArray();
				string[] slots = new string[5];

				if (entries.Length == 0)
				{
					Logging.WriteLineToLog("TravelLog: Cannot make menu string because there is no history.");
					return string.Join("\x1", slots);
				}

				if (m_offsetCurrent == -1)
				{
					Logging.WriteLineToLog("TravelLog: Cannot make menu string because offset is -1.");
					return string.Join("\x1", slots);
				}

				int j = 0;

				if (button == TravelDirection.Back)
				{
					for (int i = m_offsetCurrent - 1; i >= 0 && j < slots.Length; i--)
					{
						Logging.WriteLineToLog("TravelLog: Debug Back: Index is {0:n0}, Offset: {1:n0}", i, m_offsetCurrent);
						slots[j++] = string.IsNullOrWhiteSpace(entries[i].Title) ? entries[i].DisplayUrl : entries[i].Title;
					}

					return string.Join("\x1", slots);
				}

				if (button == TravelDirection.Forward)
				{
					for (int i = m_offsetCurrent + 1; i < entries.Length && j < slots.Length; i++)
					{
						Logging.WriteLineToLog("TravelLog: Debug Forward: Index is {0:n0}, Offset: {1:n0}", i, m_offsetCurrent);
						slots[j++] = string.IsNullOrWhiteSpace(entries[i].Title) ? entries[i].DisplayUrl : entries[i].Title;
					}

					return string.Join("\x1", slots);
				}
			}
			catch (Exception e)
			{
				Logging.WriteLineToLog("*** ERROR IN TRAVEL LOG: {0}", e);
				throw;
			}

			throw new InvalidOperationException("Not supposed to get here.");
		}

		public void Dispose()
		{
			// CEF calls this every time it is done walking the history.
			// Correct solution would be *anther* class just for walking
			// the history.
		}
	}
}