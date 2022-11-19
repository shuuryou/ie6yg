using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Pipes;
using System.Linq;
using System.Net;
using System.ServiceProcess;
using System.Text;
using yomigaeri_shared;

namespace yomigaeri_server
{
	public static class Program
	{
		internal static IniFileReader Settings { get; private set; }

		public static void Main()
		{
			Logging.OpenLog("YGSERVER");

			string ini_file_location =
				Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config.ini");

			Logging.WriteLineToLog("Main: Now going to load settings from \"{0}\".",
				ini_file_location);

			string server_root;
			IPAddress server_ip;
			int server_port;

			try
			{
				Settings = new IniFileReader(ini_file_location);

				server_ip = IPAddress.Parse(Settings.Get("Server", "IPAddress", string.Empty));
				server_port = int.Parse(Settings.Get("Server", "Port"), NumberStyles.None, CultureInfo.InvariantCulture);
				server_root = Settings.Get("Server", "RootDirectory");
			}
			catch (Exception e)
			{
				Logging.WriteLineToLog("Main: Error loading settings: {0}", e);
				throw;
			}

			Logging.WriteLineToLog("Main: Settings were loaded successfully.");

			Logging.WriteLineToLog("Main: Starting server.");

			try
			{
				Server server = new Server(server_ip, server_port, server_root);

				server.Begin();
			} catch (Exception e)
			{
				Logging.WriteLineToLog("Main: Server crashed: {0}", e);
				throw;
			}

			/*
			ServiceBase[] ServicesToRun;
			ServicesToRun = new ServiceBase[]
			{
				new YGSERVICE()
			};
			ServiceBase.Run(ServicesToRun);
			*/
		}
	}
}