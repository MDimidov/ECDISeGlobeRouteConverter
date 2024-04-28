using ECDIS_eGloebe___RouteConverter.DTOs;
using ECDIS_eGloebe___RouteConverter.Utilities;
using System;
using System.IO;
using System.Windows.Forms;

namespace ECDIS_eGloebe___RouteConverter
{
	public partial class HomeControl : UserControl
	{
		HomeInfo homeInfo = new HomeInfo();

		public HomeControl()
		{
			InitializeComponent();
			ShowFormInfo();
		}

		private void btnExportHomeInfo_Click(object sender, EventArgs e)
		{
			SetFormInfo();
			File.WriteAllText(GetProjectDirectory(), ExportToStringXml());
			MessageBox.Show($"You successfuly exported file dir: {GetProjectDirectory()}");
		}

		private void btnImportHomeInfo_Click(object sender, EventArgs e)
		{
			ImportFromXml();
			ShowFormInfo();
		}

		private void SetFormInfo()
		{
			homeInfo.MasterName = tbMaster.Text;
			homeInfo.ChiefMateName = tbChiefMate.Text;
			homeInfo.DelegateOfficerName = tbDelegateOfficer.Text;
			homeInfo.WKO1Name = tbWKO1.Text;
			homeInfo.WKO2Name = tbWKO2.Text;
			homeInfo.VesselName = tbVessel.Text;
			homeInfo.PortFrom = tbPortFrom.Text;
			homeInfo.PortTo = tbPortTo.Text;

			if (double.TryParse(tbDraftFwd.Text, out double draftFwd))
			{
				homeInfo.draftFWD = draftFwd;
			}
			else
			{
				MessageBox.Show($"Please enter valid number for draft");
			}

			if (double.TryParse(tbDraftAft.Text, out double draftAft))
			{
				homeInfo.draftAFT = draftAft;
			}
			else
			{
				MessageBox.Show($"Please enter valid number for draft");
			}
		}

		private void ShowFormInfo()
		{
			tbMaster.Text = homeInfo.MasterName;
			tbChiefMate.Text = homeInfo.ChiefMateName;
			tbDelegateOfficer.Text = homeInfo.DelegateOfficerName;
			tbWKO1.Text = homeInfo.WKO1Name;
			tbWKO2.Text = homeInfo.WKO2Name;
			tbVessel.Text = homeInfo.VesselName;
			tbPortFrom.Text = homeInfo.PortFrom;
			tbPortTo.Text = homeInfo.PortTo;
			tbDraftFwd.Text = homeInfo.draftFWD.ToString();
			tbDraftAft.Text = homeInfo.draftAFT.ToString();
		}

		private string ExportToStringXml()
		{
			return new XmlHelper().Serialize(homeInfo, "homeInfo");
		}

		private void ImportFromXml()
		{
			string xmlStirng= File.ReadAllText(GetProjectDirectory());
			homeInfo = new XmlHelper().Deserialize<HomeInfo>(xmlStirng, "homeInfo");
		}

		private static string GetProjectDirectory()
		{
			string relativePath = @"../../../";

			return relativePath + "HomeInfo.xml";
		}

		
	}
}
