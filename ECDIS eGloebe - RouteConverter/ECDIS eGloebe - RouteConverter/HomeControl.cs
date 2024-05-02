using ECDIS_eGloebe___RouteConverter.DTOs;
using ECDIS_eGloebe___RouteConverter.Utilities;
using System;
using System.IO;
using System.Windows.Forms;

namespace ECDIS_eGloebe___RouteConverter
{
	public partial class HomeControl : UserControl
	{
		OpenFileDialog openFileDialog = new OpenFileDialog()
		{
			Filter = "Ship info (*.xml)|*.xml|All files (*.*)|*.*",
			FilterIndex = 1,
			RestoreDirectory = true,
		};

		SaveFileDialog saveFileDialog = new SaveFileDialog()
		{
			Filter = "Ship info (*.xml)|*.xml|All files (*.*)|*.*",
			FilterIndex = 1,
			RestoreDirectory = true,
		};


		HomeInfoDto homeInfo = new HomeInfoDto();

		public HomeControl()
		{
			InitializeComponent();
			ShowFormInfo();
		}

		private void btnExportHomeInfo_Click(object sender, EventArgs e)
		{
			SetFormInfo();
			//File.WriteAllText(GetProjectDirectory(), ExportToStringXml());
			ExportToStringXml();
			//MessageBox.Show($"You successfuly exported file dir: {GetProjectDirectory()}");
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
			tbDraftFwd.Text = homeInfo.draftFWD.ToString("f1");
			tbDraftAft.Text = homeInfo.draftAFT.ToString("f1");
		}

		private void ExportToStringXml()
		{
			string homeInfoXml =  new XmlHelper().Serialize(homeInfo, "homeInfo");
			if (saveFileDialog.ShowDialog() == DialogResult.OK)
			{
				try
				{
					using (StreamWriter sw = new StreamWriter(saveFileDialog.FileName))
					{
						sw.Write(homeInfoXml);
					}
					MessageBox.Show("Your file is succesfully saved");
				}
				catch (Exception ex)
				{
					MessageBox.Show("Eror while saving" + ex.Message);
				}
			}
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

						homeInfo = new XmlHelper()
							.Deserialize<HomeInfoDto>(xmlStirng, "homeInfo");
					}
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

			return relativePath + "HomeInfo.xml";
		}


	}
}
