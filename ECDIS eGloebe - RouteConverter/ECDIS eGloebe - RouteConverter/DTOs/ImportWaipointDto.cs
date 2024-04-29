using System.Xml.Serialization;

namespace ECDIS_eGloebe___RouteConverter.DTOs
{
	[XmlType("waypoint")]
	public class ImportWaipointDto
	{
		[XmlElement("position")]
		public ImportPositionDto Position { get; set; }

		//[XmlElement("turningRadius")]
		//public double TurningRadius { get; set; }

		//[XmlElement("tidalLevel")]
		//public double TidalLevel { get; set; }

		[XmlElement("legType")]
		public string LegType { get; set; }

		[XmlElement("xte")]
		public ImportXteDto Xte { get; set; }

		//[XmlElement("antiGroundingStatus")]
		//public string AntiGroundingStatus { get; set; }

		//[XmlElement("weatherOptimizable")]
		//public bool WeatherOptimizable { get; set; }

		[XmlElement("timeZone")]
		public int TimeZone { get; set; }

		public double DistanceFromLastWp { get; set; }

		public double Course { get; set; }
	}
}
