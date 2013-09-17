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
		private static object syncRoot = new Object();
		public static string authorizationData;
		
        #region Constants
        private static string storageService = "storageservice";
        public static string server = "https://connect.sandbox.mimeo.com/2012/02/"; // default Service
        public static string serverSandbox = "https://connect.sandbox.mimeo.com/2012/02/"; // Sandbox Service
		public static string serverProduction = "https://connect.mimeo.com/2012/02/";				// Production Service
		public static XNamespace ns = "http://schemas.mimeo.com/MimeoConnect/2012/02/StorageService";
		public static XNamespace nsOrder = "http://schemas.mimeo.com/MimeoConnect/2012/02/Orders";
		public static XNamespace nsESLOrder = "http://schemas.mimeo.com/EnterpriseServices/2008/09/OrderService";
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
            this.Initialize(user, password, true);
        }
		public void Initialize(string user, string password, bool sandbox)
		{
			ServicePointManager.ServerCertificateValidationCallback += ValidateRemoteCertificate;
			var userName_password = user + ":" + password;
			byte[] encDataByte = System.Text.Encoding.UTF8.GetBytes(userName_password);
			authorizationData = "Basic " + Convert.ToBase64String(encDataByte);

            server = (sandbox) ? serverSandbox : serverProduction;
		}
		#endregion

		#region  Api Calls

		public void GetNewOrderRequest(XmlDocument doc)
		{
			Uri ordersEndpoint;

			ordersEndpoint = new Uri(server + "orders/GetOrderRequest");
			HttpWebGet(doc, ordersEndpoint);
		}

		public XmlDocument GetQuote(XmlDocument doc)
		{
			var ordersEndpoint = new Uri(server + "orders/GetQuote");
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

        public XmlDocument FindStoreItems(XmlDocument doc)
        {
            Uri storageEndpoint = new Uri(server + storageService + "/FindStoreItems");
            return HttpWebPost(doc, storageEndpoint);
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

		private static WebResponse GetWebResponseWithFaultException(HttpWebRequest httpWebRequest)
		{
			WebResponse response = null;

			try
			{
				response = httpWebRequest.GetResponse();
			}
			catch(WebException we)
			{
				if(we.Status == WebExceptionStatus.ProtocolError)
				{
					using(Stream stream = we.Response.GetResponseStream())
					{
						var doc = new XmlDocument();
						doc.XmlResolver = null;
						doc.Load(stream);
					}
				}
				throw;
			}
			return response;
		}
		#endregion

		#region Helpers
		
		public void setShippingOption(XmlDocument orderRequest, string shipOption, int idx)
		{
			// Select 2nd day delivery for Recipient 1
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
		public void AddLineItems(XmlDocument orderRequest, List<Document> documents)
		{

			XmlNode lineItemsRootNode = orderRequest.GetElementsByTagName("LineItems")[0];
			lineItemsRootNode.RemoveAll();

			XmlNode addLineItemRequest = orderRequest.CreateElement("AddLineItemRequest", nsESLOrder.NamespaceName);


			foreach(Document doc in documents)
			{
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
			nodeId.InnerText = "524D8417-6E1D-466F-A0BE-627A4FA742DF";
			paymentMethodCodesRoot.AppendChild(nodeId);
		}
		public void PopulateRecipients(XmlDocument orderRequest, List<Address> addresses)
		{
			//Address inAddress = addresses.FirstOrDefault();
			XmlNode recipientsRootNode = orderRequest.GetElementsByTagName("Recipients")[0];
			recipientsRootNode.RemoveAll();

			foreach(Address inAddress in addresses)
			{
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
				XmlNode shippingMethodIdTextNode = orderRequest.CreateTextNode("00000000-0000-0000-0000-000000000001");
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
