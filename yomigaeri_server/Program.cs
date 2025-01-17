﻿using System;
using System.Globalization;
using System.IO;
using System.Net;
using yomigaeri_shared;

namespace yomigaeri_server
{
	public static class Program
	{
		internal static IniFileReader Settings { get; private set; }

		public static void Main()
		{
			Logging.OpenLog("YGSERVER");

			Logging.WriteBannerToLog("IE6 Yomigaeri's Server");

			string ini_file_location =
				Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config.ini");

			Logging.WriteLineToLog("Main: Now going to load settings from \"{0}\".",
				ini_file_location);

			string server_root;
			IPAddress server_ip;
			int server_port;
			string js_template;

			try
			{
				Settings = new IniFileReader(ini_file_location);

				server_ip = IPAddress.Parse(Settings.Get("Server", "IPAddress", string.Empty));
				server_port = int.Parse(Settings.Get("Server", "Port"), NumberStyles.None, CultureInfo.InvariantCulture);
				server_root = Settings.Get("Server", "RootDirectory");
				js_template = Settings.Get("Server", "JavascriptTemplate");
			}
			catch (Exception e)
			{
				Logging.WriteLineToLog("Main: Error loading settings: {0}", e);
				throw;
			}

			Logging.WriteLineToLog("Main: Settings were loaded successfully.");

			// TODO: Use service instead of console app

			Logging.WriteLineToLog("Main: Starting server.");

			try
			{
				HttpServer server = new HttpServer(server_ip, server_port, server_root, js_template);

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