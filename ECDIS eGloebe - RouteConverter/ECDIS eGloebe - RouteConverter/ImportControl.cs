using ECDIS_eGloebe___RouteConverter.DTOs;
using ECDIS_eGloebe___RouteConverter.Utilities;
using System;
using System.IO;
using System.Windows.Forms;

namespace ECDIS_eGloebe___RouteConverter
{
	public partial class ImportControl : UserControl
	{
		ImportRouteDto routeDto = new ImportRouteDto();

		public ImportControl()
		{
			InitializeComponent();
		}

		private void btnImportRoute_Click(object sender, EventArgs e)
		{
			ImportFromXml();
		}

		private void ImportFromXml()
		{
			string xmlStirng = File.ReadAllText(GetProjectDirectory());
			xmlStirng = xmlStirng.Replace("http://www.sevencs.com/eglobe/route", "");
			routeDto = new XmlHelper()
				.Deserialize<ImportRouteDto>(xmlStirng, "");

			CalculateAllDistancesBetweenWp();
			;
		}

		private static string GetProjectDirectory()
		{
			string relativePath = @"../../../";

			return relativePath + "route.rte";
		}

		private (double, double) GetDistanceFromLastWp(
			ImportPositionDto lastPsn,
			ImportPositionDto currentPsn)
		{
			double RSH = (currentPsn.Latitude - lastPsn.Latitude) * 60 + 0.0000000000001;

			double RD = (currentPsn.Longtitude - lastPsn.Longtitude) * 60;

			double latM = lastPsn.Latitude + (RSH / 120);

			double OTSH = RD * Math.Cos(latM * Math.PI / 180);

			double course = Math.Abs(Math.Atan(OTSH / RSH) * 180 / Math.PI);

			if (RSH < 0 && RD >= 0)
			{
				course = 180 - course;
			}
			else if (RSH < 0 && RD < 0)
			{
				course = 180 + course;
			}
			else if (RSH >= 0 && RD < 0)
			{
				course = 360 - course;
			}

			double distance = Math.Round(RSH / Math.Cos(course * Math.PI / 180), 1);

			course = Math.Round(course, 1);

			return (course, distance);
		}

		private void CalculateAllDistancesBetweenWp()
		{
			for (int i = 1; i < routeDto.Waipoints.Length; i++)
			{
				(routeDto.Waipoints[i].Course, routeDto.Waipoints[i].DistanceFromLastWp) =
					GetDistanceFromLastWp(
					routeDto.Waipoints[i - 1].Position,
					routeDto.Waipoints[i].Position);
			}
		}
	}
}
