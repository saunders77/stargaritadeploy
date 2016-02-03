<%@ WebHandler Language="C#" Class="RichApiAgaveWeb.RichApiRestProxy" %>
using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Web;
using System.Net;

namespace RichApiAgaveWeb
{
	/// <summary>
	/// Summary description for RichApiRestProxy
	/// </summary>
	public class RichApiRestProxy : IHttpHandler
	{
		private const string RequestUrl = "RequestUrl";

		public void ProcessRequest(HttpContext context)
		{
			context.Response.TrySkipIisCustomErrors = true;
			string requestUrl = context.Request.QueryString[RequestUrl];
			if (string.IsNullOrWhiteSpace(requestUrl))
			{
				context.Response.StatusCode = (int)HttpStatusCode.BadRequest;
				return;
			}

			HttpWebRequest webReq = CreateHttpWebRequest(context, requestUrl);
			string method = webReq.Method.ToUpperInvariant();
			if (method == "POST" || method == "PATCH" || method == "PUT")
			{
				using (Stream requestStream = webReq.GetRequestStream())
				{
					CopyStream(context.Request.InputStream, requestStream);
				}
			}
			HttpWebResponse webResp = null;

			try
			{
				webResp = webReq.GetResponse() as HttpWebResponse;
			}
			catch(WebException webEx)
			{
				webResp = webEx.Response as HttpWebResponse;
			}

			if (webResp == null)
			{
				context.Response.StatusCode = (int)HttpStatusCode.InternalServerError;
				return;
			}

			CopyHttpWebResponse(webResp, context);

			webResp.Close();
		}

		private HttpWebRequest CreateHttpWebRequest(HttpContext context, string requestUrl)
		{
			List<string> excludedKeys = new List<string>() { "ACCEPT", "CONTENT-TYPE", "CONTENT-LENGTH", "CONNECTION", "HOST", "USER-AGENT", "REFERER" };

			HttpWebRequest webReq = WebRequest.CreateHttp(requestUrl);
			webReq.Method = context.Request.HttpMethod;
			if (context.Request.QueryString["UseFiddler"] == "1")
			{
				webReq.Proxy = new System.Net.WebProxy("127.0.0.1", 8888);
			}

			if (!string.IsNullOrWhiteSpace(context.Request.ContentType))
			{
				webReq.ContentType = context.Request.ContentType;
			}
			else if (webReq.Method != "GET" && webReq.Method != "DELETE")
			{
				webReq.ContentType = "application/json";
			}

			string accept = context.Request.Headers["ACCEPT"];
			if (!string.IsNullOrWhiteSpace(accept))
			{
				webReq.Accept = accept;
			}

			foreach(string key in context.Request.Headers.Keys)
			{
				if (!excludedKeys.Contains(key.ToUpperInvariant()))
				{
					webReq.Headers[key] = context.Request.Headers[key];
				}
			}

			return webReq;
		}

		private void CopyHttpWebResponse(HttpWebResponse webResp, HttpContext context)
		{
			context.Response.StatusCode = (int)webResp.StatusCode;
			context.Response.ContentType = webResp.ContentType;

			string[] excludedKeys = new string[] { "CONTENT-LENGTH", "TRANSFER-ENCODING" };
			foreach (string key in webResp.Headers.Keys)
			{
				if (!excludedKeys.Contains(key.ToUpperInvariant()))
				{
					context.Response.Headers[key] = webResp.Headers[key];
				}
			}

			using (Stream respStream = webResp.GetResponseStream())
			{
				CopyStream(respStream, context.Response.OutputStream);
			}
		}

		private void CopyStream(Stream input, Stream output)
		{
			byte[] buffer = new byte[1024];
			int count;
			while ((count = input.Read(buffer, 0, buffer.Length)) > 0)
			{
				output.Write(buffer, 0, count);
			}
		}

		public bool IsReusable
		{
			get
			{
				return false;
			}
		}
	}
}