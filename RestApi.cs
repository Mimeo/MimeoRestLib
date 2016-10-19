using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Web;
using System.Xml;
using System.Xml.Linq;
using System.Threading.Tasks;

namespace Mimeo.MimeoConnect
{
	public sealed class RestApi
	{
		private static RestApi instance;
        private static XmlDocument _getOrderRequestXml = new XmlDocument();
        private static XmlDocument _getTemplateOrderXml = new XmlDocument();
        private static XDocument _getProductTemplatesXml = new XDocument();
        Dictionary<string, string> savedTemplates = new Dictionary<string, string>();
        Dictionary<string, string> savedProdTemplates = new Dictionary<string, string>();
		private static object syncRoot = new Object();
		public static string authorizationData;
		
		#region Constants
		private static string storageService = "storageservice";
		public static string server = "https://connect.sandbox.mimeo.com/2012/02/"; // default Service
		public static string serverSandbox = "https://connect.sandbox.mimeo.com/2012/02/"; // Sandbox Service
		public static string serverProduction = "https://connect.mimeo.com/2012/02/";				// Production Service
		public static string serverUKProduction = "https://connect.mimeo.co.uk/2012/02/";				// Production Service

		public static XNamespace ns = "http://schemas.mimeo.com/MimeoConnect/2012/02/StorageService";
		public static XNamespace nsOrder = "http://schemas.mimeo.com/MimeoConnect/2012/02/Orders";
		public static XNamespace nsAccountService = "http://schemas.mimeo.com/MimeoConnect/2012/02/AccountService";
		public static XNamespace nsESLOrder = "http://schemas.mimeo.com/EnterpriseServices/2008/09/OrderService";
		public static XNamespace nsESLStorage = "http://schemas.mimeo.com/EnterpriseServices/2008/09/StorageService";
		public static XNamespace nsi = "http://www.w3.org/2001/XMLSchema-instance";
		public static XNamespace nsDoc = "http://schemas.mimeo.com/dom/3.0/Document.xsd";
		#endregion


		private RestApi()
		{

		}

		#region Properties
		public static RestApi GetInstance
		{
			get
			{
				lock(syncRoot)
				{
					if(instance == null)
					{
						instance = new RestApi();
					}
				}

				return instance;
			}
		}
		#endregion

		#region Public Interface
		public void Initialize(string user, string password)
        {
            System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
			this.Initialize(user, password, true);
		}
		public void Initialize(string user, string password, bool sandbox)
        {
            System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
			string inServer = (sandbox) ? serverSandbox : serverProduction;
			this.Initialize(user, password, inServer);
		}
		public void Initialize(string user, string password, string serverEndPoint)
		{
            System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
			ServicePointManager.ServerCertificateValidationCallback += ValidateRemoteCertificate;
			var userName_password = user + ":" + password;
			byte[] encDataByte = System.Text.Encoding.UTF8.GetBytes(userName_password);
			authorizationData = "Basic " + Convert.ToBase64String(encDataByte);

			server = serverEndPoint;
		}

		public string getEnvironment(string environment)
		{

			if (environment.Equals("production", StringComparison.InvariantCultureIgnoreCase))
			{
				return serverProduction;
			}
			else if (environment.Equals("ukproduction", StringComparison.InvariantCultureIgnoreCase))
			{
				return serverUKProduction;
			}
			else
			{
				return serverSandbox;
			}
		}


		#endregion

		#region  Api Calls

		public XmlDocument GetOrderRequestXml()
		{
			if (_getOrderRequestXml.ChildNodes.Count == 0)
			{
				GetNewOrderRequest(_getOrderRequestXml);
			}

			return _getOrderRequestXml;
		}

        public XmlDocument GetTemplateOrderXml(string tmpId)
        {
            if (savedTemplates.ContainsKey(tmpId))
            {
                _getTemplateOrderXml.LoadXml(savedTemplates[tmpId]);
            }
            else
            {
                GetNewTemplateOrder(_getTemplateOrderXml, tmpId);
                savedTemplates.Add(tmpId, _getTemplateOrderXml.InnerXml);
            }
           
            return _getTemplateOrderXml;
        }

        public XDocument GetTemplateProductsXml(string tmpId)
        {
            if (savedProdTemplates.ContainsKey(tmpId))
            {
                _getProductTemplatesXml = XDocument.Parse(savedProdTemplates[tmpId]);
            }
            else
            {
                _getProductTemplatesXml = FindFolderStoreItems(tmpId);
                savedProdTemplates.Add(tmpId, _getProductTemplatesXml.ToString());
            }

            return _getProductTemplatesXml;
        }

        public void ClearTemplateCache(string tmpId)
        {
            if (savedProdTemplates.ContainsKey(tmpId))
            {
                savedProdTemplates.Remove(tmpId);
            }

            return;
        }

		public XDocument GetUser()
		{
			Uri userEndpoint;

			userEndpoint = new Uri(server + "AccountService/user");
			var doc = HttpWebAction(userEndpoint, "GET");


			XDocument xDoc = XDocument.Parse(doc.OuterXml);
			return xDoc;
		}

		public void GetNewOrderRequest(XmlDocument doc)
		{
			Uri ordersEndpoint;

			ordersEndpoint = new Uri(server + "orders/GetOrderRequest");
			HttpWebGet(doc, ordersEndpoint);
		}

        public void GetNewTemplateOrder(XmlDocument doc, string tmpId)
        {
            Uri ordersEndpoint;

            string myURL = String.Format("Orders/NewProduct?template=custom&documentTemplateId={0}", tmpId);
            ordersEndpoint = new Uri(server + myURL);

            HttpWebGet(doc, ordersEndpoint);
        }

		public void GetItemQuoteRequest(XmlDocument doc)
		{
			Uri ordersEndpoint;

			ordersEndpoint = new Uri(server + "orders/GetItemQuoteRequest");
			HttpWebGet(doc, ordersEndpoint);
		}

		public XDocument GetItemQuote(XDocument priceDoc, XDocument content)
		{
			XmlDocument doc = new XmlDocument();
			using (var reader = priceDoc.CreateReader())
			{
				doc.Load(reader);
			}

			XmlNode contents = doc.GetElementsByTagName("content", "http://schemas.mimeo.com/dom/3.0/Document.xsd")[0];
			//contents.InnerXml = content.Document.ToString();
			contents.InnerText = System.Security.SecurityElement.Escape(content.Document.ToString());

			Uri storageEndpoint = new Uri(server + "orders/GetItemQuote");
			XmlDocument xmlDoc = HttpWebPost(doc, storageEndpoint);

			XDocument retXMl = XDocument.Parse(xmlDoc.OuterXml);
			return retXMl;

		}

		public XmlDocument GetQuote(XmlDocument doc)
		{
			var ordersEndpoint = new Uri(server + "orders/GetQuote");
			return HttpWebPost(doc, ordersEndpoint);
		}

        public XmlDocument GetTemplateQuote(XmlDocument doc)
        {
            var ordersEndpoint = new Uri(server + "orders/GetQuote/Template");
            return HttpWebPost(doc, ordersEndpoint);
        }

        public XmlDocument GetTemplateItemQuote(XmlDocument doc)
        {
            var ordersEndpoint = new Uri(server + "orders/GetItemQuote/Template");
            return HttpWebPost(doc, ordersEndpoint);
        }

		public XmlDocument GetShippingOptions(XmlDocument doc)
		{
			var ordersEndpoint = new Uri(server + "orders/GetShippingOptions");
			return HttpWebPost(doc, ordersEndpoint);
		}

		public XDocument GetFolderInfo(string folder)
		{

			string docFolder = "/Document" + folder + "/";
			Uri storageEndpoint = new Uri(server + storageService + docFolder);


			try
			{
				XmlDocument doc = new XmlDocument();
				HttpWebGet(doc, storageEndpoint);

				XDocument xDoc = XDocument.Parse(doc.OuterXml);
				return xDoc;
			}
			catch (WebException we)
			{
				throw;
			}
			catch (Exception ex)
			{
				throw ex;
			}
		}

		public XDocument GetPrintFiles(string folder)
		{

			string docFolder =  folder + "/";
			Uri storageEndpoint = new Uri(server + storageService + docFolder);


			try
			{
				XmlDocument doc = new XmlDocument();
				HttpWebGet(doc, storageEndpoint);

				XDocument xDoc = XDocument.Parse(doc.OuterXml);
				return xDoc;
			}
			catch (WebException we)
			{
				throw;
			}
			catch (Exception ex)
			{
				throw ex;
			}
		}

		public XmlDocument postXml(Uri storageEndpoint, XmlDocument doc)
		{
			return HttpWebPost(doc, storageEndpoint);
		}

        public XmlDocument getXml(Uri storageEndpoint)
        {
            XmlDocument doc = new XmlDocument();
            HttpWebGet(doc, storageEndpoint);

            return doc;
        }

		public XmlDocument FindStoreItems(XmlDocument doc)
		{
			Uri storageEndpoint = new Uri(server + storageService + "/FindStoreItems");
			return HttpWebPost(doc, storageEndpoint);
		}

		public XDocument DeleteDocument(string folder, string fileId)
		{
			var storageEndpoint = new Uri(server + storageService + "/Document/" + folder);

			// REMOVE
			var deleteEndpoint = new Uri(storageEndpoint + "/" + fileId);

			var doc = HttpWebAction(deleteEndpoint, "DELETE");

			XDocument xDoc = XDocument.Parse(doc.OuterXml);
			return xDoc;

		}

		public XDocument DeletePrintFile(string folder, string fileId)
		{
			var storageEndpoint = new Uri(server + storageService + "/" + folder);

			// REMOVE
			var deleteEndpoint = new Uri(storageEndpoint + "/" + fileId);

			var doc = HttpWebAction(deleteEndpoint, "DELETE");

			XDocument xDoc = XDocument.Parse(doc.OuterXml);
			return xDoc;

		}

		public XDocument updateDocument(string docId, string fileId, string templateId)
		{

			// Get Print File Information
			var fileXML = GetStoreItem(fileId);
			string printFileId = (from file in fileXML.Descendants(nsESLStorage + "StoreItem")
							   select file.Element(nsESLStorage + "Id").Value).FirstOrDefault();
			string pageCount = (from file in fileXML.Descendants(nsESLStorage + "ItemDetails")
								select file.Element(nsESLStorage + "PageCount").Value).FirstOrDefault();

			// Get Template Information
			var templateXML = GetStoreItem(templateId);
			string templateName = (from file in templateXML.Descendants(nsESLStorage + "StoreItem")
							   select file.Element(nsESLStorage + "Name").Value).FirstOrDefault();

			// Get Document Information
			var docXML = GetStoreItem(docId, StoreItemLevelOfDetail.IncludeFolder);
			string docFolderName = (from file in docXML.Descendants(nsESLStorage + "Folder")
								   select file.Element(nsESLStorage + "Name").Value).FirstOrDefault();
			string docName = (from file in docXML.Descendants(nsESLStorage + "StoreItem")
									select file.Element(nsESLStorage + "Name").Value).FirstOrDefault();

			var documentXML = GetDocument(docId);

			var docTempId = documentXML.Descendants(nsOrder + "DocumentTemplateId").FirstOrDefault();
			docTempId.Value = templateId;
			var docTempName = documentXML.Descendants(nsOrder + "DocumentTemplateName").FirstOrDefault();
			docTempName.Value = templateName;
			var source = documentXML.Descendants(nsOrder + "Source").FirstOrDefault();
			source.Value = printFileId;
			var range = documentXML.Descendants(nsOrder + "Range").FirstOrDefault();
			range.Value = string.Format("[1,{0}]", pageCount);

			// Update Document
			string createDocument = string.Format("/Document");
			Uri storageEndpoint = new Uri(server + storageService + createDocument);
			XmlDocument inXml = new XmlDocument();
			inXml.Load(documentXML.CreateReader());
			XmlDocument newDoc = HttpWebPost(inXml, storageEndpoint, "PUT");

			documentXML = XDocument.Parse(newDoc.OuterXml);
			return documentXML;

		}


		public XDocument createDocument(string newDocName, string newDocFolder, string fileId, string templateId)
		{

			// Get Print File Information
			var fileXML = GetStoreItem(fileId);
			string printFileId = (from file in fileXML.Descendants(nsESLStorage + "StoreItem")
								  select file.Element(nsESLStorage + "Id").Value).FirstOrDefault();
			string pageCount = (from file in fileXML.Descendants(nsESLStorage + "ItemDetails")
								select file.Element(nsESLStorage + "PageCount").Value).FirstOrDefault();

			// Get Template Information
			var templateXML = GetStoreItem(templateId);
			string templateName = (from file in templateXML.Descendants(nsESLStorage + "StoreItem")
								   select file.Element(nsESLStorage + "Name").Value).FirstOrDefault();

			// Get Document Information
			string docFolderName = newDocFolder;

			var documentXML = GetNewDocument(templateId);

			var docName = documentXML.Descendants(nsOrder + "Name").FirstOrDefault();
			docName.Value = newDocName;
			var docTempId = documentXML.Descendants(nsOrder + "DocumentTemplateId").FirstOrDefault();
			docTempId.Value = templateId;
			var docTempName = documentXML.Descendants(nsOrder + "DocumentTemplateName").FirstOrDefault();
			docTempName.Value = templateName;
			var source = documentXML.Descendants(nsOrder + "Source").FirstOrDefault();
			source.Value = printFileId;
			var range = documentXML.Descendants(nsOrder + "Range").FirstOrDefault();
			range.Value = string.Format("[1,{0}]", pageCount);

			// Create Document
			string createDocument = string.Format("/Document/{0}", newDocFolder);
			Uri storageEndpoint = new Uri(server + storageService + createDocument);
			XmlDocument inXml = new XmlDocument();
			inXml.Load(documentXML.CreateReader());
			XmlDocument newDoc = HttpWebPost(inXml, storageEndpoint, "POST");

			documentXML = XDocument.Parse(newDoc.OuterXml);
			return documentXML;

		}


        public XDocument createNewDocument(string newDocName, string newDocFolder, string fileId, string templateId, string pages)
        {

            // Get Print File Information
            var fileXML = GetStoreItem(fileId);
            string printFileId = (from file in fileXML.Descendants(nsESLStorage + "StoreItem")
                                  select file.Element(nsESLStorage + "Id").Value).FirstOrDefault();
            string pageCount = (from file in fileXML.Descendants(nsESLStorage + "ItemDetails")
                                select file.Element(nsESLStorage + "PageCount").Value).FirstOrDefault();

            // Get Document Information
            string docFolderName = newDocFolder;

            var documentXML = GetNewDocument(templateId);

            var docName = documentXML.Descendants(nsOrder + "Name").FirstOrDefault();
            docName.Value = newDocName;

            XmlDocument inXml = new XmlDocument();
            inXml.LoadXml(documentXML.ToString());
            PopulateProductSections(inXml, printFileId, pages, "1");

            // Create Document
            string createDocument = string.Format("/Document/{0}", newDocFolder);
            Uri storageEndpoint = new Uri(server + storageService + createDocument);
            XmlDocument newDoc = HttpWebPost(inXml, storageEndpoint, "POST");

            documentXML = XDocument.Parse(newDoc.OuterXml);
            return documentXML;

        }


        public void PopulateProductSection(XmlDocument orderRequest, string printFileId, string pageCount, string copies)
        {
            var qty = orderRequest.GetElementsByTagName("Quantity")[0];
            qty.InnerText = copies;
            var source = orderRequest.GetElementsByTagName("Source")[0];
            source.InnerText = printFileId;
            var range = orderRequest.GetElementsByTagName("Range")[0];
            range.InnerText = string.Format("[1,{0}]", pageCount);
        }

        public void PopulateProductSections(XmlDocument orderRequest, string printFileId, string pageCount, string copies)
        {
            var qty = orderRequest.GetElementsByTagName("Quantity")[0];
            qty.InnerText = copies;
            XmlNodeList ContentRootNode = orderRequest.GetElementsByTagName("DocumentSection");
            int nodesCnt = ContentRootNode.Count;
            foreach (XmlNode node in ContentRootNode)
            {
                node["Source"].InnerText = printFileId;
                string range = node["Range"].InnerText;
                if (range != "1")
                {
                    String rangeIs = String.Format("[{0}, {1}]", (nodesCnt > 1) ? "2" : "1", (nodesCnt > 1) ? ((int.Parse(pageCount) - nodesCnt)+1).ToString() : pageCount);
                    node["Range"].InnerText = rangeIs;
                }
            }
        }

        public void PopulateAddress(XmlDocument orderRequest, Address address)
        {
            var firstName = orderRequest.GetElementsByTagName("FirstName")[0];
            firstName.InnerText = address.firstName;

            var lastName = orderRequest.GetElementsByTagName("LastName")[0];
            lastName.InnerText = address.lastName;

            var street = orderRequest.GetElementsByTagName("Street")[0];
            street.InnerText = address.street;

            var apt = orderRequest.GetElementsByTagName("ApartmentOrSuite")[0];
            apt.InnerText = address.apartmentOrSuite;

            var careof = orderRequest.GetElementsByTagName("CareOf")[0];
            careof.InnerText = address.careOf;

            var city = orderRequest.GetElementsByTagName("City")[0];
            city.InnerText = address.city;

            var state = orderRequest.GetElementsByTagName("StateOrProvince")[0];
            state.InnerText = address.state;

            var country = orderRequest.GetElementsByTagName("Country")[0];
            country.InnerText = address.country;

            var postal = orderRequest.GetElementsByTagName("PostalCode")[0];
            postal.InnerText = address.postalCode;

            var phone = orderRequest.GetElementsByTagName("TelephoneNumber")[0];
            phone.InnerText = address.telephone;
        }

		public XDocument GetInfo(string friendlyId, string action)
		{
			Uri ordersEndpoint;
			XmlDocument resultDoc = new XmlDocument();
			string statusPath = string.Format("orders/{0}/{1}", friendlyId, action);
			ordersEndpoint = new Uri(server + statusPath);
			HttpWebGet(resultDoc, ordersEndpoint);

			return XDocument.Parse(resultDoc.OuterXml);
		}
		public void AddDiscount(XmlDocument orderRequest, string discount)
		{
			XmlNode node = orderRequest.CreateNode(XmlNodeType.Element, "DiscountCodes", nsOrder.NamespaceName);

			XmlNode nodeString = orderRequest.CreateElement("string", nsOrder.NamespaceName);
			nodeString.InnerText = discount;
			node.AppendChild(nodeString);

			orderRequest.DocumentElement.AppendChild(node);
		}
		public void AddPackagingSlip(XmlDocument orderRequest, string salutation, string memo)
		{
			XmlNode node = orderRequest.CreateNode(XmlNodeType.Element, "PackagingSlip", nsOrder.NamespaceName);

			XmlNode nodeSalutationType = orderRequest.CreateElement("SalutationType", nsOrder.NamespaceName);
			nodeSalutationType.InnerText = (string.IsNullOrEmpty(salutation)) ? "None" : salutation;
			node.AppendChild(nodeSalutationType);

			XmlNode nodeMemo = orderRequest.CreateElement("Memo", nsOrder.NamespaceName);
			nodeMemo.InnerText = (string.IsNullOrEmpty(memo)) ? "None" : memo;
			node.AppendChild(nodeMemo);

			orderRequest.DocumentElement.AppendChild(node);
		}

		#endregion

		#region Protocol
		public XmlDocument HttpWebPost(XmlDocument doc, Uri ordersEndpoint)
		{
			return HttpWebPost(doc, ordersEndpoint, "POST");
		}

		public XmlDocument HttpWebPost(XmlDocument doc, Uri ordersEndpoint, string action)
		{
			var webrequest = (HttpWebRequest)WebRequest.Create(ordersEndpoint);
			webrequest.Headers.Add(HttpRequestHeader.Authorization, authorizationData);
			webrequest.Method = action;
            webrequest.ProtocolVersion = HttpVersion.Version11;
			// Set the ContentType property of the WebRequest.
			webrequest.ContentType = "application/xml";

			var result = new XmlDocument();
			result.XmlResolver = null;

			using(var sw = new StringWriter())
			{
				using(var xtw = new XmlTextWriter(sw))
				{
					doc.WriteTo(xtw);

					byte[] byteArray = Encoding.UTF8.GetBytes(sw.ToString());
					webrequest.ContentLength = byteArray.Length;

					Stream dataStream = webrequest.GetRequestStream();
					// Write the data to the request stream.
					dataStream.Write(byteArray, 0, byteArray.Length);
					// Close the Stream object.
					dataStream.Close();

					WebResponse response = GetWebResponseWithFaultException(webrequest);
					Stream s = response.GetResponseStream();
					result.Load(s);

					dataStream.Close();
					response.Close();
				}
			}

			return result;
		}

		public void HttpWebGet(XmlDocument doc, Uri ordersEndpoint)
		{
			var encoding = new UTF8Encoding();

			HttpWebRequest objRequest;
			HttpWebResponse objResponse;
			StreamReader srResponse;

			// Initialize request object  
			objRequest = (HttpWebRequest)WebRequest.Create(ordersEndpoint);
			objRequest.Headers.Add(HttpRequestHeader.Authorization, authorizationData);
			objRequest.Method = "GET";
			objRequest.AllowWriteStreamBuffering = true;

			// Get response
			objResponse = (HttpWebResponse)objRequest.GetResponse();
			srResponse = new StreamReader(objResponse.GetResponseStream(), Encoding.ASCII);
			string xmlOut = srResponse.ReadToEnd();
			srResponse.Close();

			if(xmlOut != null && xmlOut.Length > 0)
			{
				doc.LoadXml(xmlOut);
			}
		}

		private XmlDocument HttpWebAction(Uri ordersEndpoint, string action)
		{
			XmlDocument retDocument = new XmlDocument();

			var encoding = new UTF8Encoding();

			HttpWebRequest objRequest;
			HttpWebResponse objResponse;
			StreamReader srResponse;

			// Initialize request object  
			objRequest = (HttpWebRequest)WebRequest.Create(ordersEndpoint);
			objRequest.Headers.Add(HttpRequestHeader.Authorization, authorizationData);
			objRequest.Method = action;
			objRequest.AllowWriteStreamBuffering = true;

			// Get response
			objResponse = (HttpWebResponse)objRequest.GetResponse();
			srResponse = new StreamReader(objResponse.GetResponseStream(), Encoding.ASCII);
			string xmlOut = srResponse.ReadToEnd();
			srResponse.Close();

			if (xmlOut != null && xmlOut.Length > 0)
			{
				retDocument.LoadXml(xmlOut);
			}

			return retDocument;
		}


		private static XDocument doUploadPDF(Uri uri, string fileName)
		{
			XDocument myResult = new XDocument();

			string boundary = "----------" + DateTime.Now.Ticks.ToString("x");

			try
			{
				HttpWebRequest webrequest = (HttpWebRequest)WebRequest.Create(uri);
				webrequest.Headers.Add(HttpRequestHeader.Authorization, authorizationData);
				webrequest.ContentType = "multipart/form-data; boundary=" + boundary;
				webrequest.Method = "POST";

				// Build up the post message header
				StringBuilder sb = new StringBuilder();
				sb.Append("--");
				sb.Append(boundary);
				sb.Append("\r\n");
				sb.Append("Content-Disposition: form-data; name=\"");
				sb.Append("file");
				sb.Append("\"; filename=\"" + fileName + "\"");
				sb.Append("\r\n");
				sb.Append("Content-Type: application/octet-stream");
				sb.Append("\r\n");
				sb.Append("\r\n");

				string postHeader = sb.ToString();

				byte[] postHeaderBytes = Encoding.UTF8.GetBytes(postHeader);

				// Build the trailing boundary string as a byte array
				// ensuring the boundary appears on a line by itself
				byte[] boundaryBytes = Encoding.ASCII.GetBytes("\r\n--" + boundary + "\r\n");

				using (Stream fileStream = File.OpenRead(fileName))
				{
					if (fileStream != null)
					{
						long length = postHeaderBytes.Length + fileStream.Length + boundaryBytes.Length;

						webrequest.ContentLength = length;

						Stream requestStream = webrequest.GetRequestStream();

						// Write out our post header
						requestStream.Write(postHeaderBytes, 0, postHeaderBytes.Length);

						// Write out the file contents
						byte[] buffer = new Byte[checked((uint)Math.Min(4096, (int)fileStream.Length))];

						int bytesRead = 0;

						int i = 0;

						while ((bytesRead = fileStream.Read(buffer, 0, buffer.Length)) != 0)
						{
							requestStream.Write(buffer, 0, bytesRead);
							i++;
						}

						// Write out the trailing boundary
						requestStream.Write(boundaryBytes, 0, boundaryBytes.Length);

						WebResponse response = GetWebResponseWithFaultException(webrequest);
						Stream s = response.GetResponseStream();
						StreamReader sr = new StreamReader(s);
						myResult = XDocument.Load(sr);
					}
					else
					{
						throw new Exception("File stream is null");
					}
				}
			}
			catch (WebException we)
			{
				throw;
			}
			catch (Exception)
			{
				throw;
			}

			return myResult;
		}


		private static WebResponse GetWebResponseWithFaultException(HttpWebRequest httpWebRequest)
		{
			WebResponse response = null;

			try
			{
				response = httpWebRequest.GetResponse();
			}
			catch(WebException we)
			{
				String restError = null;
				if (we.Status == WebExceptionStatus.ProtocolError)
				{
					using (Stream stream = we.Response.GetResponseStream())
					{
						var doc = new XmlDocument();
						doc.XmlResolver = null;
						doc.Load(stream);
						restError = doc.InnerXml;
					}
				}
				throw new System.Exception(restError, we.InnerException);
			}
			return response;
		}
		#endregion

		#region Helpers

		public XDocument UploadPDF(string folder, string filePath)
		{
			XDocument retDoc = new XDocument();

			try
			{
				Uri storageEndpoint = new Uri(server + storageService + "/" + folder);
				retDoc = doUploadPDF(storageEndpoint, filePath);
			}
			catch (Exception ex)
			{
				throw ex;
			}

			return retDoc;
		}

		public Guid FindDocumentIdbyName(string name)
		{
			String retDocId = "00000000-0000-0000-0000-000000000001";
			XDocument docs = FindStoreItem(name);

			retDocId = (from file in docs.Descendants(nsESLStorage + "StoreItem")
						select file.Element(nsESLStorage + "Id").Value).FirstOrDefault();

			retDocId = (String.IsNullOrEmpty(retDocId) == false) ? retDocId : "00000000-0000-0000-0000-000000000001";
			return Guid.Parse(retDocId);
		}

		public XDocument FindStoreItem(string name)
		{
			return FindStoreItem(name, "Document");
		}

		public XDocument FindStoreItem(string name, string type)
		{

			string xmlRequest =
		   "<StoreItemSearchCriteria xmlns=\"http://schemas.mimeo.com/EnterpriseServices/2008/09/StorageService\">" +
		   "<PageInfo xmlns=\"http://schemas.mimeo.com/EnterpriseServices/2008/09/Common/Search\"><PageSize>20</PageSize><PageNumber>1</PageNumber></PageInfo>" +
		   "<Name>" + name + "</Name>" +
		   "<Type>"+ type + "</Type>" +
		   "</StoreItemSearchCriteria>";

			XmlDocument doc = new XmlDocument();
			doc.LoadXml(xmlRequest);

			XmlDocument apiResult = FindStoreItems(doc);

			XDocument retXMl = XDocument.Parse(apiResult.OuterXml);
			return retXMl;
		}


        public XDocument FindFolderStoreItems(string folderId)
        {

            string xmlRequest =
           "<StoreItemSearchCriteria xmlns=\"http://schemas.mimeo.com/EnterpriseServices/2008/09/StorageService\">" +
           "<PageInfo xmlns=\"http://schemas.mimeo.com/EnterpriseServices/2008/09/Common/Search\"><PageSize>20</PageSize><PageNumber>1</PageNumber></PageInfo>" +
           "<FolderId>" + folderId + "</FolderId>" +
           "<LevelOfDetail>All</LevelOfDetail>" + 
           "</StoreItemSearchCriteria>";

            XmlDocument doc = new XmlDocument();
            doc.LoadXml(xmlRequest);

            XmlDocument apiResult = FindStoreItems(doc);

            XDocument retXMl = XDocument.Parse(apiResult.OuterXml);
            return retXMl;
        }

		public String GetRootFolder()
		{

			string xmlRequest =
			"<GetTopLevelFoldersRequest xmlns=\"http://schemas.mimeo.com/EnterpriseServices/2008/09/StorageService\">" +
			"<Scope>Default</Scope>" +
			"<LevelOfDetail>Basic</LevelOfDetail>" +
			"</GetTopLevelFoldersRequest>";

			XmlDocument doc = new XmlDocument();
			doc.LoadXml(xmlRequest);

			Uri storageEndpoint = new Uri(server + storageService + "/GetTopLevelFolders");
            XmlDocument apiResult = HttpWebPost(doc, storageEndpoint); 
            XDocument retXMl = XDocument.Parse(apiResult.OuterXml);

            var folderId = (from folder in retXMl.Descendants(nsESLStorage + "StoreFolder")
                            where folder.Element(nsESLStorage + "WellKnownFolderType").Value == "PersonalStorageRootFolder"
                            select folder.Element(nsESLStorage + "Id").Value).FirstOrDefault();


            return folderId;
		}
        public XDocument GetTopFolders()
        {

            string xmlRequest =
            "<GetTopLevelFoldersRequest xmlns=\"http://schemas.mimeo.com/EnterpriseServices/2008/09/StorageService\">" +
            "<Scope>Default</Scope>" +
            "<LevelOfDetail>IncludeSubFolders</LevelOfDetail>" +
            "</GetTopLevelFoldersRequest>";

            XmlDocument doc = new XmlDocument();
            doc.LoadXml(xmlRequest);

            Uri storageEndpoint = new Uri(server + storageService + "/GetTopLevelFolders");
            XmlDocument apiResult = HttpWebPost(doc, storageEndpoint);
            XDocument retXMl = XDocument.Parse(apiResult.OuterXml);

            return retXMl;
        }

        public XDocument GetFolderSubFolders(string folderId)
        {

            string xmlRequest =
            "<GetFolderRequest xmlns=\"http://schemas.mimeo.com/EnterpriseServices/2008/09/StorageService\">" +
            "<FolderId>" + folderId + "</FolderId>" +
            "<LevelOfDetail>All</LevelOfDetail>" +
            "</GetFolderRequest>";

            XmlDocument doc = new XmlDocument();
            doc.LoadXml(xmlRequest);

            Uri storageEndpoint = new Uri(server + storageService + "/GetFolderSubFolders");
            XmlDocument apiResult = HttpWebPost(doc, storageEndpoint);
            XDocument retXMl = XDocument.Parse(apiResult.OuterXml);

            return retXMl;
        }

		public XDocument GetDocument(string Id)
		{
			XmlDocument retDoc = new XmlDocument();

			try
			{
				string docFinder = "/Document/GetDocument?DocumentId=" + Id;
				var storageEndpoint = new Uri(server + storageService + docFinder);

				HttpWebGet(retDoc, storageEndpoint);
			}
			catch (Exception ex)
			{
				throw ex;
			}

			XDocument retXMl = XDocument.Parse(retDoc.OuterXml);
			return retXMl;
		}


		public XDocument GetNewDocument(string templateId)
		{
			XmlDocument retDoc = new XmlDocument();

			try
			{
				string docFinder = "/NewDocument?DocumentTemplateId=" + templateId;
				var storageEndpoint = new Uri(server + storageService + docFinder);

				HttpWebGet(retDoc, storageEndpoint);
			}
			catch (Exception ex)
			{
				throw ex;
			}

			XDocument retXMl = XDocument.Parse(retDoc.OuterXml);
			return retXMl;
		}

		public XDocument GetStoreItem(string Id)
		{
			return GetStoreItem(Id, StoreItemLevelOfDetail.IncludeItemDetails);
		}

		public XDocument GetStoreItem(string Id, StoreItemLevelOfDetail LevelOfDetail)
		{

			string xmlRequest =
		   "<GetStoreItemRequest xmlns=\"http://schemas.mimeo.com/EnterpriseServices/2008/09/StorageService\">" +
		   "<ItemId>" + Id + "</ItemId>" +
		   "<LevelOfDetail>" + LevelOfDetail +"</LevelOfDetail>" +
		   "</GetStoreItemRequest>";

			XmlDocument doc = new XmlDocument();
			doc.LoadXml(xmlRequest);

			Uri storageEndpoint = new Uri(server + storageService + "/GetStoreItem");
			XmlDocument xmlDoc = HttpWebPost(doc, storageEndpoint);

			XDocument retXMl = XDocument.Parse(xmlDoc.OuterXml);
			return retXMl;
		}

		/// <summary>
		/// Get Document XML 
		/// </summary>
		/// <param name="inDocument"></param>
		/// <returns></returns>
		public XDocument GetDocumentXml(XDocument inDocument)
		{
			XmlDocument doc = new XmlDocument();
			doc.Load(inDocument.CreateReader());
			Uri storageEndpoint = new Uri(server + storageService + "/Document/GetDocumentXml");
			XmlDocument xmlDoc = HttpWebPost(doc, storageEndpoint);

			XDocument retXMl = XDocument.Parse(xmlDoc.OuterXml);
			return retXMl;
		}

		public void setShippingOption(XmlDocument orderRequest, string shipOption, int idx)
		{
			XmlNode addRecipientRequestRoot = orderRequest.GetElementsByTagName("AddRecipientRequest")[idx];
			addRecipientRequestRoot.ChildNodes[1].InnerText = shipOption;
		}

		public void SpecifyProcessingHours(XmlDocument orderRequest, int hours)
		{
			XmlNode node = orderRequest.CreateNode(XmlNodeType.Element, "Options", nsESLOrder.NamespaceName);
			XmlNode nodeAdditionalProcessingHours = orderRequest.CreateElement("AdditionalProcessingHours",
				"http://schemas.mimeo.com/EnterpriseServices/2008/09/OrderService");
			nodeAdditionalProcessingHours.InnerText = hours.ToString();
			node.AppendChild(nodeAdditionalProcessingHours);
			orderRequest.DocumentElement.AppendChild(node);
		}
		
		public void AddProcessingHours(XmlDocument orderRequest, int hours)
		{
			XmlNode nodeAdditionalProcessingHours = orderRequest.CreateElement("AdditionalProcessingHours",
				"http://schemas.mimeo.com/EnterpriseServices/2008/09/OrderService");
			nodeAdditionalProcessingHours.InnerText = hours.ToString();
			orderRequest.DocumentElement.AppendChild(nodeAdditionalProcessingHours);
		}

		public void AddLineItems(XmlDocument orderRequest, List<Document> documents)
		{
			XmlNode lineItemsRootNode = orderRequest.GetElementsByTagName("LineItems")[0];
			lineItemsRootNode.RemoveAll();

			foreach(Document doc in documents)
			{
				XmlNode addLineItemRequest = orderRequest.CreateElement("AddLineItemRequest", nsESLOrder.NamespaceName);

				if (doc.id == Guid.Empty)
				{
					// Let's get GUID by Name
					doc.id = FindDocumentIdbyName(doc.Name);
				}

				XmlNode nameNode = orderRequest.CreateElement("Name", nsESLOrder.NamespaceName);
				XmlNode nameTextNode = orderRequest.CreateTextNode(doc.Name);
				nameNode.AppendChild(nameTextNode);

				XmlNode storeItemReferenceNode = orderRequest.CreateElement("StoreItemReference", nsESLOrder.NamespaceName);
				XmlNode idNode = orderRequest.CreateElement("Id", nsESLOrder.NamespaceName);
				XmlNode idTextNode = orderRequest.CreateTextNode(doc.id.ToString());
				idNode.AppendChild(idTextNode);
				storeItemReferenceNode.AppendChild(idNode);


				XmlNode quantitydNode = orderRequest.CreateElement("Quantity", nsESLOrder.NamespaceName);
				XmlNode quantityTextNode = orderRequest.CreateTextNode(doc.Quantity.ToString());
				quantitydNode.AppendChild(quantityTextNode);


				addLineItemRequest.AppendChild(nameNode);
				addLineItemRequest.AppendChild(storeItemReferenceNode);
				addLineItemRequest.AppendChild(quantitydNode);


				lineItemsRootNode.AppendChild(addLineItemRequest);
			}
		}
		public void PopulatePaymentMethod(XmlDocument orderRequest)
		{
			XmlNode paymentMethodCodesRoot = orderRequest.GetElementsByTagName("PaymentMethod")[0];
			paymentMethodCodesRoot.Attributes["i:type"].Value = "UserCreditLimitPaymentMethod";

			for(; paymentMethodCodesRoot.HasChildNodes; )
			{
				paymentMethodCodesRoot.RemoveChild(paymentMethodCodesRoot.FirstChild);
			}

			XmlNode nodeId = orderRequest.CreateElement("Id",
				"http://schemas.mimeo.com/EnterpriseServices/2008/09/OrderService");
			nodeId.InnerText = "00000000-0000-0000-0000-000000000001";
			paymentMethodCodesRoot.AppendChild(nodeId);
		}

        public void PopulateRecipients(XmlDocument orderRequest, List<Address> addresses)
		{
			//Address inAddress = addresses.FirstOrDefault();
			XmlNode recipientsRootNode = orderRequest.GetElementsByTagName("Recipients")[0];
			recipientsRootNode.RemoveAll();

			foreach(Address inAddress in addresses)
			{

				if (inAddress.firstName == inAddress.lastName)
				{                
					//Some companies will be sending us a full name:  First\bLast Name
					//Let address that rule
					string[] tmpName = inAddress.firstName.Split(' ');
					inAddress.firstName = tmpName[0];
					inAddress.lastName = (tmpName.Length > 1) ? tmpName[1] : inAddress.lastName;
					inAddress.lastName = inAddress.lastName.Replace(",", "");
				}

				XmlNode addRecipientRequest = orderRequest.CreateElement("AddRecipientRequest", nsESLOrder.NamespaceName);
				XmlNode address = orderRequest.CreateElement("Address", nsESLOrder.NamespaceName);

				XmlNode newFirstnameNode = orderRequest.CreateElement("FirstName", nsESLOrder.NamespaceName);
				XmlNode firstnameTextNode = orderRequest.CreateTextNode(inAddress.firstName);
				newFirstnameNode.AppendChild(firstnameTextNode);
				address.AppendChild(newFirstnameNode);

				XmlNode newLastnameNode = orderRequest.CreateElement("LastName", nsESLOrder.NamespaceName);
				XmlNode lastnameTextNode = orderRequest.CreateTextNode(inAddress.lastName);
				newLastnameNode.AppendChild(lastnameTextNode);
				address.AppendChild(newLastnameNode);

				XmlNode newStreetNode = orderRequest.CreateElement("Street", nsESLOrder.NamespaceName);
				XmlNode streetTextNode = orderRequest.CreateTextNode(inAddress.street);
				newStreetNode.AppendChild(streetTextNode);
				address.AppendChild(newStreetNode);

				XmlNode newAptOrSuiteNode = orderRequest.CreateElement("ApartmentOrSuite", nsESLOrder.NamespaceName);
				XmlNode aptOrSuiteTextNode = orderRequest.CreateTextNode(inAddress.apartmentOrSuite);
				newAptOrSuiteNode.AppendChild(aptOrSuiteTextNode);
				address.AppendChild(newAptOrSuiteNode);

				XmlNode newCareOfNode = orderRequest.CreateElement("CareOf", nsESLOrder.NamespaceName);
				XmlNode careOfTextNode = orderRequest.CreateTextNode(inAddress.careOf);
				newCareOfNode.AppendChild(careOfTextNode);
				address.AppendChild(newCareOfNode);

				XmlNode newCityNode = orderRequest.CreateElement("City", nsESLOrder.NamespaceName);
				XmlNode cityTextNode = orderRequest.CreateTextNode(inAddress.city);
				newCityNode.AppendChild(cityTextNode);
				address.AppendChild(newCityNode);

				XmlNode newStateNode = orderRequest.CreateElement("StateOrProvince", nsESLOrder.NamespaceName);
				XmlNode stateTextNode = orderRequest.CreateTextNode(inAddress.state);
				newStateNode.AppendChild(stateTextNode);
				address.AppendChild(newStateNode);

				XmlNode newCountryNode = orderRequest.CreateElement("Country", nsESLOrder.NamespaceName);
				XmlNode countryTextNode = orderRequest.CreateTextNode(inAddress.country);
				newCountryNode.AppendChild(countryTextNode);
				address.AppendChild(newCountryNode);

				XmlNode newPostalCodeNode = orderRequest.CreateElement("PostalCode", nsESLOrder.NamespaceName);
				XmlNode postalCodeTextNode = orderRequest.CreateTextNode(inAddress.postalCode);
				newPostalCodeNode.AppendChild(postalCodeTextNode);
				address.AppendChild(newPostalCodeNode);

				XmlNode newTelephoneNumberNode = orderRequest.CreateElement("TelephoneNumber", nsESLOrder.NamespaceName);
				XmlNode TelephoneNumberTextNode = orderRequest.CreateTextNode(inAddress.telephone);
				newTelephoneNumberNode.AppendChild(TelephoneNumberTextNode);
				address.AppendChild(newTelephoneNumberNode);

				XmlNode newResidentialNode = orderRequest.CreateElement("IsResidential", nsESLOrder.NamespaceName);
				XmlNode residentialTextNode = orderRequest.CreateTextNode("true");
				newResidentialNode.AppendChild(residentialTextNode);
				address.AppendChild(newResidentialNode);


				XmlNode shippingMethodIdNode = orderRequest.CreateElement("ShippingMethodId", nsESLOrder.NamespaceName);
				XmlNode shippingMethodIdTextNode = orderRequest.CreateTextNode(inAddress.shipping);
				shippingMethodIdNode.AppendChild(shippingMethodIdTextNode);


				addRecipientRequest.AppendChild(address);
				addRecipientRequest.AppendChild(shippingMethodIdNode);
				recipientsRootNode.AppendChild(addRecipientRequest);

			}
		}

		public void AddReferenceNumber(XmlDocument orderRequest, string refNbr)
		{
			XmlNode node = orderRequest.CreateNode(XmlNodeType.Element, "ReferenceNumber", "http://schemas.mimeo.com/EnterpriseServices/2008/09/OrderService");
			node.InnerText = refNbr;
			orderRequest.DocumentElement.AppendChild(node);
		}

		public void PopulateSpecialInstructionCodes(XmlDocument orderRequest, List<string> SIs)
		{
			if (SIs.Count > 0)
			{
				XmlNode specialInstructionCodesRoot = orderRequest.GetElementsByTagName("SpecialInstructionCodes")[0];

				bool firstOne = true;
				foreach (string si in SIs)
				{
					if (si.Length > 0)
					{
						if (firstOne)
						{
							// 1st SI
							specialInstructionCodesRoot.ChildNodes[0].InnerText = si;
							firstOne = false;
						}
						else
						{
							// 2nd SI
							XmlNode importNode = specialInstructionCodesRoot.OwnerDocument.ImportNode(specialInstructionCodesRoot.ChildNodes[0], true);
							importNode.InnerText = si;
							specialInstructionCodesRoot.AppendChild(importNode);
						}
					}
				}

			}
		}

		public static string findShippingId(XmlDocument doc, string shippingMethodName)
		{
			XDocument shippingDoc = XDocument.Parse(doc.InnerXml);
			string retShipId = "-1";

			retShipId = (from file in shippingDoc.Descendants(nsESLOrder + "ShippingMethodDetail")
						 where file.Element(nsESLOrder + "Name").Value == shippingMethodName
						 select file.Element(nsESLOrder + "Id").Value).FirstOrDefault();

			return retShipId;
		}

		public void PlaceOrder(XmlDocument doc)
		{
			var ordersEndpoint = new Uri(server + "orders/PlaceOrder");
			var webrequest = (HttpWebRequest)WebRequest.Create(ordersEndpoint);
			webrequest.Headers.Add(HttpRequestHeader.Authorization, authorizationData);
			webrequest.Method = "POST";
			// Set the ContentType property of the WebRequest.
			webrequest.ContentType = "application/xml";


			using(var sw = new StringWriter())
			{
				using(var xtw = new XmlTextWriter(sw))
				{
					doc.WriteTo(xtw);

					byte[] byteArray = Encoding.UTF8.GetBytes(sw.ToString());
					webrequest.ContentLength = byteArray.Length;

					Stream dataStream = webrequest.GetRequestStream();
					// Write the data to the request stream.
					dataStream.Write(byteArray, 0, byteArray.Length);
					// Close the Stream object.
					dataStream.Close();

					WebResponse response = GetWebResponseWithFaultException(webrequest);
					Stream s = response.GetResponseStream();
					doc.Load(s);
					dataStream.Close();
					response.Close();
				}
			}
		}

		private bool ValidateRemoteCertificate(object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors policyErrors)
		{
			return true;
		}
		#endregion
	}
}
