using System.Collections.Generic;
using System.Xml.Serialization;

namespace ECDIS_eGloebe___RouteConverter.DTOs
{
	[XmlType("route")]
	public class ImportRouteDto
	{
		//[XmlAttribute("xmlns")]
		//public string Xmlns { get; set; }

		[XmlArray("waypoints")]
		public List<ImportWaipointDto> Waipoints { get; set; } = new List<ImportWaipointDto>();
	}
}
