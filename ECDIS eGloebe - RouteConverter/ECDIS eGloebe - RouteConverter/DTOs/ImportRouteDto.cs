using System.Xml.Serialization;

namespace ECDIS_eGloebe___RouteConverter.DTOs
{
	[XmlType("route")]
	public class ImportRouteDto
	{
		//[XmlAttribute("xmlns")]
		//public string Xmlns { get; set; }

		[XmlArray("waypoints")]
		public ImportWaipointDto[] Waipoints;
	}
}
