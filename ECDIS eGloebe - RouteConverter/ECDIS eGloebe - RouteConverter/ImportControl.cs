using ECDIS_eGloebe___RouteConverter.DTOs;
using ECDIS_eGloebe___RouteConverter.Utilities;
using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using static ECDIS_eGloebe___RouteConverter.Common.Common;

namespace ECDIS_eGloebe___RouteConverter
{
	public partial class ImportControl : UserControl
	{
		OpenFileDialog openFileDialog = new OpenFileDialog()
		{
			Filter = "ECDIS Route file (*.rte)|*.rte|All files (*.*)|*.*",
			FilterIndex = 1,
			RestoreDirectory = true,
		};

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
			if (openFileDialog.ShowDialog() == DialogResult.OK)
			{
				try
				{
					// Отваряне на избрания файл за четене
					using (StreamReader sr = new StreamReader(openFileDialog.OpenFile()))
					{
						// Четене на съдържание от файла и извеждане на конзолата
						string xmlStirng = sr.ReadToEnd();

						xmlStirng = xmlStirng.Replace("http://www.sevencs.com/eglobe/route", "");
						RouteDto = new XmlHelper()
							.Deserialize<ImportRouteDto>(xmlStirng, "");
					}

					CalculateAllDistancesBetweenWp();
					CalculateDistanceToGo();
					MessageBox.Show($"You import route {openFileDialog.FileName} successfuly");
				}
				catch (Exception ex)
				{
					MessageBox.Show("Error while loading the file: " + ex.Message);
				}
			}
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
			for (int i = 1; i < RouteDto.Waipoints.Count; i++)
			{
				(RouteDto.Waipoints[i].Course, RouteDto.Waipoints[i].DistanceFromLastWp) =
					GetDistanceFromLastWp(
					RouteDto.Waipoints[i - 1].Position,
					RouteDto.Waipoints[i].Position);
			}
		}

		private void CalculateDistanceToGo()
		{
			double totalDistance = RouteDto
				.Waipoints
				.Sum(w => w.DistanceFromLastWp);

			foreach (var wp in RouteDto.Waipoints)
			{
				totalDistance -= wp.DistanceFromLastWp;
				wp.DistanceToGo = Math.Round(totalDistance, 1);
			}
		}



	}
}
