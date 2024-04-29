using System.Xml.Serialization;

namespace ECDIS_eGloebe___RouteConverter.DTOs
{
	[XmlType("position")]
	public class ImportPositionDto
	{
		[XmlAttribute("latitude")]
		public double Latitude { get; set; }


		[XmlAttribute("longitude")]
		public double Longtitude { get; set; }



		public char LatDir
		{
			get
			{
				if(Latitude < 0)
				{
					return 'S';
				}

				return 'N';
			}
		}

		public int LatDegrees => (int)Latitude;

		public double LatMinutes => (Latitude - LatDegrees) * 60;



		public char LongDir
		{
			get
			{
				if (Longtitude < 0)
				{
					return 'W';
				}

				return 'E';
			}
		}

		public int LongDegrees => (int)Longtitude;

		public double LongMinutes => (Longtitude - LongDegrees) * 60;
	}
}
