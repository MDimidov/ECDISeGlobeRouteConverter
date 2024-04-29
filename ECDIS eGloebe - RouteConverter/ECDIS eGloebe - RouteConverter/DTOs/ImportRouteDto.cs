using System;
using System.Xml.Serialization;

namespace ECDIS_eGloebe___RouteConverter.DTOs
{
	[XmlType("route")]
	public class ImportRouteDto
	{
		[XmlArray("waypoints")]
		public ImportWaipointDto[] Waipoints;
	}
}
