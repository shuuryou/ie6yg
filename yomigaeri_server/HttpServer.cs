using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Web;
using yomigaeri_shared;

namespace yomigaeri_server
{
	public sealed class HttpServer
	{
		private const string SERVER_NAME = "IE6YG Minimal HTTP Server";

		private readonly TcpListener m_HttpListener;
		private readonly Thread m_HttpServerThread;
		private const int HTTP_SERVER_MAX_BACKLOG = 25;

		private readonly string m_ContentDir;
		private readonly string m_JsTemplateFile;

		private readonly CancellationTokenSource m_Cancel;

		private enum ReplyDownloadResult
		{
			Success = 0,
			Error = 1,
			Bail = 2
		}

		private enum ReplyJsTemplateResult
		{
			Success = 0,
			Error = 1,
			NoMoreUsers = 2
		}

		public HttpServer(IPAddress listenAddr, int listenPort, string contentDir, string jsTemplateFile)
		{
			if (listenAddr == null)
				throw new ArgumentNullException("listenAddr");

			if (contentDir == null)
				throw new ArgumentNullException("contentDir");

			if (listenPort < 0 || listenPort > 65535)
				throw new ArgumentOutOfRangeException("listenPort");

			if (!Directory.Exists(contentDir))
				throw new FileNotFoundException("Content directory does not exist.", contentDir);

			if (!File.Exists(Path.Combine(contentDir, jsTemplateFile)))
				throw new FileNotFoundException("JavaScript template file does not exist.", jsTemplateFile);

			m_ContentDir = Path.GetFullPath(contentDir);
			m_JsTemplateFile = Path.Combine(m_ContentDir, jsTemplateFile);

			m_HttpListener = new TcpListener(listenAddr, listenPort);
			m_HttpServerThread = new Thread(HttpServerMain);

			m_Cancel = new CancellationTokenSource();
		}

		public void Begin()
		{
			m_HttpListener.Start(HTTP_SERVER_MAX_BACKLOG);
			m_HttpServerThread.Start(m_Cancel.Token);

#if DEBUG
			m_HttpServerThread.Join();
#endif
		}

		public void Terminate()
		{
			m_Cancel.Cancel();
		}

		private void HttpServerMain(object cancelToken)
		{
			CancellationToken token = (CancellationToken)cancelToken;

			Logging.WriteLineToLog("HttpServerMain: Entry.");

		again:

			try
			{
				Logging.WriteLineToLog("HttpServerMain: Waiting for connection.");
				TcpClient client = m_HttpListener.AcceptTcpClient();

				Logging.WriteLineToLog("HttpServerMain: Accepted a connection.");
				ThreadPool.QueueUserWorkItem(HttpAnswerConnection, client);
			}
			catch (SocketException e)
			{
				if ((e.SocketErrorCode == SocketError.Interrupted))
				{
					Logging.WriteLineToLog("HttpServerMain: HTTP socket shut down.", e);
				}

				Logging.WriteLineToLog("HttpServerMain: HTTP server socket error: {0}", e);
			}
			catch (ThreadAbortException)
			{
				Logging.WriteLineToLog("HttpServerMain: HTTP server main thread aborted.");
				return;
			}

			if (token.IsCancellationRequested)
			{
				Logging.WriteLineToLog("HttpServerMain: HTTP server main thread exiting cleanly.");
				return;
			}

			goto again;
		}

		private void HttpAnswerConnection(object state)
		{
			if (!(state is TcpClient client))
				return;

			if (!client.Connected)
				return;

			IPAddress client_addr = ((IPEndPoint)client.Client.RemoteEndPoint).Address;

			Logging.WriteLineToLog("HttpAnswerConnection: Start processing connection from {0}", client_addr);

			try
			{
				using (NetworkStream ns = client.GetStream())
				using (StreamReader sr = new StreamReader(ns))
				{
					Dictionary<string, string> client_headers = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

					string request_uri;
					{
						string header = sr.ReadLine();

						if (string.IsNullOrEmpty(header))
							return; // Browser vanished without sending headers

						if (!header.StartsWith("GET ", StringComparison.Ordinal))
							goto badrequest;

						int idx = header.LastIndexOf(' ');

						if (idx == -1)
							request_uri = header.Substring(4);
						else
							request_uri = header.Substring(4, idx - 4);

						while (!string.IsNullOrEmpty(header))
						{
							// Store any other headers that the browser sends.

							header = sr.ReadLine();

							idx = header.IndexOf(": ");

							if (idx != -1)
							{
								string key = header.Substring(0, idx);
								string value = header.Substring(idx + 2);

								if (client_headers.ContainsKey(key))
									client_headers[key] = value;
								else
									client_headers.Add(key, value);

								if (client_headers.Count > 32)
									goto servererror; // Denial of service attack
							}

						}
					}

					// Download proxy

					if (request_uri.StartsWith("/dl/", StringComparison.Ordinal))
					{
						ReplyDownloadResult result = ReplyDownload(request_uri, ns);

						switch (result)
						{
							case ReplyDownloadResult.Success:
								// Everything went well
								goto done;
							case ReplyDownloadResult.Error:
								// Bad random_id from the client or bad data from the backend
								goto servererror;
							case ReplyDownloadResult.Bail:
								// Socket error during download (client vanished/cancelled download)
								goto bail;
						}
					}

					// Arbitrary files. Dangerous. :-(

					if (string.IsNullOrWhiteSpace(request_uri) || request_uri == "/")
						request_uri = "/index.html";

					/* DANGER DANGER DANGER
					 * 
					 * Do not use Path.Combine here!
					 * 
					 * See MSDN:
					 * 
					 * "If an argument other than the first contains a rooted
					 * path, any previous path components are ignored, and the
					 * returned string begins with that rooted path component."
					 * 
					 * Path.Combine("C:\\test\\", "/Windows/System32/calc.exe")
					 * ==> C:\Windows\System32\calc.exe
					 *
					 * *sigh*
					 */

					string request_file = m_ContentDir + Path.DirectorySeparatorChar + request_uri;
					request_file = Path.GetFullPath(request_file);

					// Prevent path traversal attacks
					if (!request_file.StartsWith(m_ContentDir, StringComparison.OrdinalIgnoreCase))
					{
						Logging.WriteLineToLog("HttpAnswerConnection: Path traversal attack blocked for {0}",
							client_addr);

						goto notfound;
					}

					if (!File.Exists(request_file))
					{
						Logging.WriteLineToLog("HttpAnswerConnection: Cannot find file \"{0}\" on disk for {1}",
							request_file, client_addr);

						goto notfound;
					}

					// Special case to replace variables in JS template file
					if (request_file.Equals(m_JsTemplateFile, StringComparison.OrdinalIgnoreCase))
					{
						ReplyJsTemplateResult result = ReplyJsTemplate(ns);

						switch (result)
						{
							case ReplyJsTemplateResult.Success:
								goto done;
							case ReplyJsTemplateResult.Error:
								goto servererror;
							case ReplyJsTemplateResult.NoMoreUsers:
								goto serviceunavail;
						}
					}

					// Just serve a regular file
					Logging.WriteLineToLog("HttpAnswerConnection: Respond with 200 OK, serving {0} for {1}", request_file, client_addr);
					ReplyRequest(200, "OK", ns, request_file);
					goto done;

				badrequest:
					Logging.WriteLineToLog("HttpAnswerConnection: Respond with 401 Bad Request for {0}", client_addr);
					ReplyRequest(401, "Bad Request", ns);
					goto done;

				notfound:
					Logging.WriteLineToLog("HttpAnswerConnection: Respond with 404 Not Found for {0}", client_addr);

					ReplyRequest(404, "Not Found", ns);
					goto done;

				servererror:
					Logging.WriteLineToLog("HttpAnswerConnection: Respond with 500 Internal Server Error for {0}", client_addr);

					ReplyRequest(500, "Internal Server Error", ns);
					goto done;

				serviceunavail:
					Logging.WriteLineToLog("HttpAnswerConnection: Respond with 503 Service Unavailable Error for {0}", client_addr);

					ReplyRequest(503, "Service Unavailable", ns);
					goto done;

				done:
					client.Close();

				bail:
					;
				}
			}
			catch (Exception e)
			{
				Logging.WriteLineToLog("HttpAnswerConnection: Error processing connection for {0}: {1}", client_addr, e);
			}

			Logging.WriteLineToLog("HttpAnswerConnection: End processing connection for {0}", client_addr);
		}

		private ReplyDownloadResult ReplyDownload(string url, Stream output)
		{
			if (!output.CanWrite)
				throw new ArgumentException("Output stream must be writable.", "output");

			string random_id = url.Substring(4);

			{
				// Strip filename at the end that's used to coax IE into
				// showing it in the title bar of the download window, if any.
				int idx = random_id.IndexOf('/');

				if (idx != -1)
					random_id = random_id.Substring(0, idx);
			}

			Logging.WriteLineToLog("ReplyDownload: Download request for download ID {0}", random_id);

			string download_meta = Path.GetFullPath(Path.Combine(Program.Settings.Get("Downloads", "DownloadTempDir"), random_id)) + ".ini";

			Logging.WriteLineToLog("ReplyDownload: Download ID {0} meta: {1}", random_id, download_meta);

			if (!File.Exists(download_meta))
			{
				Logging.WriteLineToLog("ReplyDownload: Download ID {0}: meta not found.", random_id);
				return ReplyDownloadResult.Error;
			}

			IniFileReader reader = new IniFileReader(download_meta);

			/*
			 * "Download", "CEFID", downloadItem.Id
			 * "Download", "TotalBytes", downloadItem.TotalBytes
			 * "Download", "MimeType", downloadItem.MimeType
			 * "Download", "ContentDisposition", downloadItem.ContentDisposition
			 * "Download", "URL", downloadItem.Url
			 * "Download", "SuggestedFileName", downloadItem.SuggestedFileName
			 * "Download", "LocalFile", download_file
			 * 
			 * Not all of them are used (yet) below.
			 */

			string download_file = reader.Get("Download", "LocalFile");

			{
				bool found_download_file = false;

				for (int i = 0; i < 5; i++)
				{
					// Sometimes CEF is a little slow to actually create the
					// file on disk, so give it five seconds to do so before
					// giving up.

					found_download_file = File.Exists(download_file);

					if (found_download_file)
						break;

					Logging.WriteLineToLog("ReplyDownload: Download ID {0}: download file not ready yet. Wait a bit.", random_id);

					Thread.Sleep(1000);
				}

				if (!found_download_file)
				{
					Logging.WriteLineToLog("ReplyDownload: Download ID {0}: download file not found.", random_id);
					return ReplyDownloadResult.Error;
				}
			}

			long total_bytes;

			{
				bool flag = long.TryParse(reader.Get("Download", "TotalBytes"),
					NumberStyles.None, CultureInfo.InvariantCulture, out total_bytes);

				if (!flag)
				{
					Logging.WriteLineToLog("ReplyDownload: Download ID {0}: Could not get file size of download.", random_id);
					return ReplyDownloadResult.Error;
				}
			}

			string mime_type = reader.Get("Download", "MimeType", "application/octet-stream");
			string content_disposition = reader.Get("Download", "ContentDisposition");
			string suggested_file_name = reader.Get("Download", "SuggestedFileName");

			if (string.IsNullOrEmpty(content_disposition))
			{
				if (string.IsNullOrEmpty(suggested_file_name))
					content_disposition = "attachment";
				else
					content_disposition = string.Format(CultureInfo.InvariantCulture, "attachment; filename=\"{0}\"", suggested_file_name);
			}

			/*
			 * Finished with all of the preparations, so send headers.
			 */

			WriteASCIIString(output, "HTTP/1.0 200 OK\r\n");
			WriteASCIIString(output, "Content-Type: {0}\r\n", mime_type);

			if (total_bytes != 0)
				WriteASCIIString(output, "Content-Length: {0}\r\n", total_bytes);

			WriteASCIIString(output, "Content-Disposition: {0}\r\n", content_disposition);
			WriteASCIIString(output, "Date: {0:R}\r\n", DateTime.Now);
			WriteASCIIString(output, "Server: {0}\r\n", SERVER_NAME);
			WriteASCIIString(output, "Connection: Close\r\n\r\n");

			output.Flush();

			/*
			 * Send the actual binary file contents next.
			 */

			Logging.WriteLineToLog("ReplyDownload: Download ID {0}: Starting to serve download to client.", random_id);

			// If unable to write to the downloading browser for 30 seconds,
			// will throw an IOException and in the loop below and cause the
			// method to bail.
			output.WriteTimeout = 30000;

			short bail = 0;

			using (FileStream fs = new FileStream(download_file, FileMode.Open,
				FileAccess.Read, FileShare.ReadWrite, 4096, FileOptions.SequentialScan))
			{
				fs.Seek(0, SeekOrigin.Begin);

				long copied_bytes = 0;
				byte[] buf = new byte[8192];
				int read;

				// Next needs to be an infinite loop rather than Stream.CopyTo
				// or whatever other cool way you normally use with streams,
				// because here we are reading a file that Chromium is still
				// actively writing to. The file is still growing while being
				// read from here!

				// It gets even worse when total_bytes is 0, because it means
				// the server Chromium is reading from doesn't know the size
				// of the download (e.g. it's being generated on the fly).

				for (; ; )
				{
					read = fs.Read(buf, 0, buf.Length);

					if (read > 0)
					{
						try
						{
							output.Write(buf, 0, read);
						}
						catch (IOException)
						{
							// Client probably vanished (download cancelled),
							// or timeout writing to client (DOS attack?)

							Logging.WriteLineToLog("ReplyDownload: Download ID {0}: Client vanished.", random_id);

							// Deleting download_meta should make the backend
							// cancel the download and delete download_file.

							File.Delete(download_meta);

							return ReplyDownloadResult.Bail;
						}

						copied_bytes += read;
						bail = 0;
					}
					else
					{
						// Stall while waiting for more data to arrive in the
						// download file.

						Thread.Sleep(500);

						if (++bail == 600)
						{
							// Something is wrong if there is no data arriving
							// in download_file within 5 minutes.

							Logging.WriteLineToLog("ReplyDownload: Download ID {0}: Breaking out of infinite loop.", random_id);

							// Deleting download_meta should make the backend
							// cancel the download and delete download_file.

							File.Delete(download_meta);

							return ReplyDownloadResult.Bail;
						}
					}

					if (total_bytes != 0 && copied_bytes == total_bytes)
					{
						// This will handle the majority of the cases.

						break;
					}
					else if (total_bytes == 0 && !File.Exists(download_meta))
					{
						// If the size of the download isn't known, the backend
						// will delete the meta file once the download is done.
						// See the comments in MyDownloadHandler.cs about this
						// poor man's IPC.

						break;
					}
				}
			}

			/*
			 * Getting here means the download completed successfully.
			 */

			output.Flush();

			File.Delete(download_meta);
			File.Delete(download_file);

			Logging.WriteLineToLog("ReplyDownload: Download ID {0}: Download served successfully.", random_id);

			return ReplyDownloadResult.Success;
		}

		private void ReplyRequest(int response_code, string response_name, Stream output)
		{
			Encoding __response_encoding = Encoding.ASCII;

			if (!output.CanWrite)
				throw new ArgumentException("Output stream must be writable.", "output");

			byte[] response_bytes = __response_encoding.GetBytes(response_name);

			WriteASCIIString(output, "HTTP/1.0 {0} {1}\r\n", response_code, response_name);
			WriteASCIIString(output, "Content-Type: text/html; charset={0}\r\n", __response_encoding.BodyName);
			WriteASCIIString(output, "Content-Length: {0}\r\n", response_name.Length);
			WriteASCIIString(output, "Date: {0:R}\r\n", DateTime.Now);
			WriteASCIIString(output, "Server: {0}\r\n", SERVER_NAME);
			WriteASCIIString(output, "Connection: Close\r\n\r\n");

			output.Write(response_bytes, 0, response_bytes.Length);

			output.Flush();
		}

		private void ReplyRequest(int response_code, string response_name, Stream output, string content_file)
		{
			if (!output.CanWrite)
				throw new ArgumentException("Output stream must be writable.", "output");

			using (FileStream fs = File.OpenRead(content_file))
			{
				WriteASCIIString(output, "HTTP/1.0 {0} {1}\r\n", response_code, response_name);
				WriteASCIIString(output, "Content-Type: {0}\r\n", GetMimeType(content_file));
				WriteASCIIString(output, "Content-Length: {0}\r\n", fs.Length);
				WriteASCIIString(output, "Date: {0:R}\r\n", DateTime.Now);
				WriteASCIIString(output, "Server: {0}\r\n", SERVER_NAME);
				WriteASCIIString(output, "Connection: Close\r\n\r\n");

				output.Flush();

				fs.CopyTo(output);

				output.Flush();
			}
		}

		private ReplyJsTemplateResult ReplyJsTemplate(Stream output)
		{
			Encoding __response_encoding = Encoding.UTF8;

			if (!output.CanWrite)
				throw new ArgumentException("Output stream must be writable.", "output");

			Logging.WriteLineToLog("ReplyJsTemplate: Processing request for JavaScript template.");

			string js_template = File.ReadAllText(m_JsTemplateFile, __response_encoding);

			{
				bool debug = Program.Settings.Get("Frontend", "Debug", "0") == "1";
				js_template = js_template.Replace("%FRONTENDDEBUG%", debug ? "True" : "False");
			}

			{
				js_template = js_template.Replace("%RDPSERVER%",
					HttpUtility.JavaScriptStringEncode(Program.Settings.Get("RDP", "Server", "")));

				js_template = js_template.Replace("%RDPPORT%",
					HttpUtility.JavaScriptStringEncode(Program.Settings.Get("RDP", "Port", "")));

				js_template = js_template.Replace("%RDPSHELL%",
					HttpUtility.JavaScriptStringEncode(Program.Settings.Get("RDP", "Shell", "")));

				string username, password;

				try
				{
					RDPUsers.GetNextFreeUser(out username, out password);
				}
				catch (Exception e)
				{
					Logging.WriteLineToLog("ReplzJsTemplate: Error getting next free user: {0}", e);
					return ReplyJsTemplateResult.Error;
				}

				if (username == null && password == null)
					return ReplyJsTemplateResult.NoMoreUsers;

				js_template = js_template.Replace("%RDPUSERNAME%",
					HttpUtility.JavaScriptStringEncode(username));

				js_template = js_template.Replace("%RDPPASSWORD%",
					HttpUtility.JavaScriptStringEncode(password));
			}

			{
				js_template = js_template.Replace("%DOWNLOADSERVER%",
					HttpUtility.JavaScriptStringEncode(Program.Settings.Get("Downloads", "Server", "")));

				js_template = js_template.Replace("%DOWNLOADPORT%",
					HttpUtility.JavaScriptStringEncode(Program.Settings.Get("Downloads", "Port", "")));
			}

			byte[] template_bytes = __response_encoding.GetBytes(js_template);

			WriteASCIIString(output, "HTTP/1.0 200 OK\r\n");
			WriteASCIIString(output, "Content-Type: application/javascript; charset=utf-8\r\n");
			WriteASCIIString(output, "Content-Length: {0}\r\n", js_template.Length);
			WriteASCIIString(output, "Date: {0:R}\r\n", DateTime.Now);
			WriteASCIIString(output, "Server: {0}\r\n", SERVER_NAME);
			WriteASCIIString(output, "Connection: Close\r\n\r\n");

			output.Write(template_bytes, 0, template_bytes.Length);

			output.Flush();

			return ReplyJsTemplateResult.Success;
		}

		private void WriteASCIIString(Stream output, string what)
		{
			// Used for response headers and such. Avoid using for anything else.

			if (!output.CanWrite)
				throw new ArgumentException("Output stream must be writable.", "output");

			byte[] bytes = Encoding.ASCII.GetBytes(what);

			output.Write(bytes, 0, bytes.Length);

			Logging.WriteLineToLog("WriteASCIIString: \"{0}\"", what);
		}

		private void WriteASCIIString(Stream output, string format, params object[] args)
		{
			WriteASCIIString(output, string.Format(CultureInfo.InvariantCulture, format, args));
		}

		private string GetMimeType(string fileName)
		{
			const string MIME_TYPE_DEFAULT = "application/unknown";

			string ext = Path.GetExtension(fileName).ToLower();

			if (string.IsNullOrEmpty(ext))
				return MIME_TYPE_DEFAULT;

			if (ext[0] == '.')
				ext = ext.Substring(1);

			return ApacheMimeTypes.MimeTypes.ContainsKey(ext) ?
				ApacheMimeTypes.MimeTypes[ext] : MIME_TYPE_DEFAULT;
		}
	}
}