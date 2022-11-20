using Cassia;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using yomigaeri_shared;

namespace yomigaeri_server
{
	internal static class RDPUsers
	{
		public static void GetNextFreeUser(out string username, out string password)
		{
			Logging.WriteLineToLog("GetNextFreeUser: Enumerating current sessions.");

			List<string> logged_in = new List<string>();

			ITerminalServicesManager manager = new TerminalServicesManager();
			using (ITerminalServer server = manager.GetLocalServer())
			{
				server.Open();

				foreach (ITerminalServicesSession session in server.GetSessions())
					if (!string.IsNullOrEmpty(session.UserName))
						logged_in.Add(session.UserName);
			}

			Logging.WriteLineToLog("GetNextFreeUser: Accounts in use: {0}", string.Join(",", logged_in.ToArray()));

			for (int i = 1; i < int.MaxValue; i++)
			{
				string user = Program.Settings.Get("RDP", string.Format(CultureInfo.InvariantCulture,
					"Username{0}", i));

				if (string.IsNullOrWhiteSpace(user))
					break;

				if (!logged_in.Contains(user, StringComparer.OrdinalIgnoreCase))
				{
					string pass = Program.Settings.Get("RDP", string.Format(CultureInfo.InvariantCulture,
						"Password{0}", i, string.Empty));

					Logging.WriteLineToLog("GetNextFreeUser: User \"{0}\" selected.", user);

					username = user;
					password = pass;
					return;
				}
			}

			Logging.WriteLineToLog("GetNextFreeUser: No more free users available.");

			username = null;
			password = null;

			return;

		}
	}
}
