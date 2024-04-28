using System.ComponentModel.DataAnnotations;

namespace ECDIS_eGloebe___RouteConverter.Classes
{
	public class HomeInfo
	{
		public HomeInfo()
		{
			WKO1Name = "Dimidov, Mariyan";
			VesselName = "Eurocargo Salerno";
		}
		public string MasterName { get; set; }

		public string ChiefMateName { get; set; }

		public string DelegateOfficerName { get; set; }

		public string WKO1Name { get; set; }

		public string WKO2Name { get; set; }

		public string VesselName { get; set; }

		public string PortFrom { get; set; }

		public string PortTo { get; set; }

		[Range(0.1, 30.0)]
		public double draftFWD { get; set; }

		[Range(0.1, 30.0)]
		public double draftAFT { get; set; }

		public double draftMiddle => (draftFWD * draftAFT) / 2;

	}
}
