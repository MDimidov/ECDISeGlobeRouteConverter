using System;
using System.ComponentModel.DataAnnotations;
using System.Xml.Serialization;

namespace ECDIS_eGloebe___RouteConverter.DTOs
{
	[XmlRoot("HomeInfo")]
	public class HomeInfoDto
	{
		public HomeInfoDto()
		{
			WKO1Name = "Dimidov, Mariyan";
			VesselName = "Eurocargo Salerno";
		}

		[XmlElement(nameof(MasterName))]
		public string MasterName { get; set; }

		[XmlElement(nameof(ChiefMateName))]
		public string ChiefMateName { get; set; }

		[XmlElement(nameof(DelegateOfficerName))]
		public string DelegateOfficerName { get; set; }

		[XmlElement(nameof(WKO1Name))]
		public string WKO1Name { get; set; }

		[XmlElement(nameof(WKO2Name))]
		public string WKO2Name { get; set; }

		[XmlElement(nameof(VesselName))]
		public string VesselName { get; set; }

		[XmlElement(nameof(PortFrom))]
		public string PortFrom { get; set; }

		[XmlElement(nameof(PortTo))]
		public string PortTo { get; set; }

		[XmlElement(nameof(Ets))]
		public string Ets { get; set; }

		[XmlElement(nameof(Eta))]
		public string Eta { get; set; }

		[Range(0.1, 30.0)]
		[XmlElement(nameof(draftFWD))]
		public double draftFWD { get; set; }

		[Range(0.1, 30.0)]
		[XmlElement(nameof(draftAFT))]
		public double draftAFT { get; set; }

		public double draftMiddle => (draftFWD + draftAFT) / 2;

		[XmlElement(nameof(Voyage))]
		public string Voyage { get; set; }

		[XmlElement(nameof(SafetyCountDepth))]
		public string SafetyCountDepth { get; set; }

		public DateTime CreationDate { get; set;}

	}
}
