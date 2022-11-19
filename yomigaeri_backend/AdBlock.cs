using CefSharp;
using DistillNET;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using yomigaeri_shared;

namespace yomigaeri_backend
{
	public static class AdBlock
	{
		private static bool s_AdBlockEnabled;
		private static DistillNET.FilterDbCollection s_FilterDB;

		public static bool Debug { get; private set; }

		public static void Initialize()
		{
			s_AdBlockEnabled = false;
			s_FilterDB = null;
			Debug = false;

			Logging.WriteBannerToLog("AdBlock Initialize");

			bool enable = Program.Settings.Get("AdBlock", "Enable") == "1";

			if (!enable)
			{
				Logging.WriteLineToLog("AdBlock: Feature is disabled, so I will end my own suffering.");
				return;
			}

			Debug = Program.Settings.Get("AdBlock", "Debug") == "1";

			if (Debug)
				Logging.WriteLineToLog("AdBlock: Debug mode is enabled.");

			Logging.WriteLineToLog("AdBlock: Going to parse AdBlock lists.");

			string lists_str = Program.Settings.Get("AdBlock", "Lists");

			if (string.IsNullOrWhiteSpace(lists_str))
			{
				Logging.WriteLineToLog("AdBlock: Error: lists setting is empty. AdBlock will be disabled.");
				return;
			}

			string db_file = Program.Settings.Get("AdBlock", "CompilationDB");

			if (!string.IsNullOrWhiteSpace(db_file))
			{
				db_file = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
					"AdBlock", db_file);

				Logging.WriteLineToLog("AdBlock: Compiled rules will be stored in \"{0}\".", db_file);
			}
			else
			{
				db_file = null;
				Logging.WriteLineToLog("AdBlock: Compiled rules will be cached in memory only. This is bad for performance.");
			}

			if (File.Exists(db_file))
			{
				Logging.WriteLineToLog("AdBlock: Compiled rules exist, so not parsing filter lists.");
				s_FilterDB = new FilterDbCollection(db_file, false, true);
				goto done;
			}

			string[] lists = lists_str.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
			short category = 0;

			Logging.WriteLineToLog("AdBlock: DistillNET only supports AdBlock Plus format and only domain rules.");
			Logging.WriteLineToLog("AdBlock: Please, contribute a better parser if you want perfect ad blocking.");

			foreach (string list in lists)
			{
				string list_file = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
					"Resources", "AdBlock", list);

				Logging.WriteLineToLog("AdBlock: Processing rules list: \"{0}\" as ID {1}...", list_file, category);

				if (!File.Exists(list_file))
				{
					Logging.WriteLineToLog("AdBlock: Error: File does not exist. Skipping it.");
					continue;
				}

				if (s_FilterDB == null)
				{
					if (db_file != null)
						s_FilterDB = new FilterDbCollection(db_file, true, false);
					else
						s_FilterDB = new FilterDbCollection(null, false, true);
				}

				Tuple<int, int> read_bad;

				using (FileStream fs = File.OpenRead(list_file))
					read_bad = s_FilterDB.ParseStoreRulesFromStream(fs, category++);

				Logging.WriteLineToLog("AdBlock: Read {0:n0} rules and rejected {1:n0} of them.",
					read_bad.Item1, read_bad.Item2);
			}

			if (s_FilterDB == null)
			{
				Logging.WriteLineToLog("AdBlock: Error: Still no filter DB. All files invalid? AdBlock will be disabled.");
				s_AdBlockEnabled = false;
				return;
			}

			Logging.WriteLineToLog("AdBlock: Finalizing compiled rules.");
			s_FilterDB.FinalizeForRead();

		done:
			Logging.WriteLineToLog("AdBlock: Ready.");
		}

		public static bool DomainShouldBeBlocked(string domain)
		{
			if (!s_AdBlockEnabled)
				return false;

			if (domain == null)
				throw new ArgumentNullException("domain");

			if (string.IsNullOrEmpty(domain))
				throw new ArgumentOutOfRangeException("domain");

			bool black = s_FilterDB.GetFiltersForDomain(domain).Any();
			bool white = s_FilterDB.GetWhitelistFiltersForDomain(domain).Any();

			bool decision = white || !black;

			if (Debug)
				Logging.WriteLineToLog("AdBlock: Domain \"{0}\" should be blocked? {1}", domain, decision);

			return decision;
		}

		public static bool ShouldBlockRequest(IRequest request)
		{
			if (request == null)
				throw new ArgumentNullException("request");

			Uri request_uri = null;
			try
			{
				request_uri = new Uri(request.Url);
			}
			catch (UriFormatException)
			{
				if (Debug)
					Logging.WriteLineToLog("AdBlock: ShouldBlockRequest got bad URL: \"{0}\".", request.Url);
			}

			if (request_uri == null)
				return false;

			IEnumerable<UrlFilter> black = s_FilterDB.GetFiltersForDomain(request_uri.Host);
			IEnumerable<UrlFilter> white = s_FilterDB.GetWhitelistFiltersForDomain(request_uri.Host);

			foreach (UrlFilter filter in white)
				if (filter.IsMatch(request_uri, request.Headers))
				{
					if (Debug)
						Logging.WriteLineToLog("ShouldBlockRequest: WHITELIST match for \"{0}\".", request.Url);

					return false;
				}

			foreach (UrlFilter filter in black)
				if (filter.IsMatch(request_uri, request.Headers))
				{
					if (Debug) 
						Logging.WriteLineToLog("ShouldBlockRequest: BLACKLIST match for \"{0}\".", request.Url);

					return true;
				}

			if (Debug) 
				Logging.WriteLineToLog("ShouldBlockRequest: NO match for \"{0}\".", request.Url);

			return false;
		}
	}
}