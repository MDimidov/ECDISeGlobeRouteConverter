using System.Xml.Serialization;

namespace ECDIS_eGloebe___RouteConverter.DTOs
{
	[XmlType("xte")]
	public class ImportXteDto
	{
		[XmlAttribute("merge")]
		public bool Merge {  get; set; }

		[XmlElement("both-sided")]
		public double BothSided {  get; set; }

		[XmlElement("port")]
		public double Port { get; set; }

		[XmlElement("starboard")]
		public double Starboard { get; set; }
	}
}
