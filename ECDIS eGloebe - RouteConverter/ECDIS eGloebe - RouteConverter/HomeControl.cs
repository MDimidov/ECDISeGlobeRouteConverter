using ECDIS_eGloebe___RouteConverter.DTOs;
using ECDIS_eGloebe___RouteConverter.Utilities;
using System;
using System.IO;
using System.Windows.Forms;
using static ECDIS_eGloebe___RouteConverter.Common.Common;

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


		public HomeControl()
		{
			InitializeComponent();
			ShowFormInfo();
		}

		private void btnConfirmHomeInfo_Click(object sender, EventArgs e)
		{
			SetFormInfo();
		}

		private void btnImportHomeInfo_Click(object sender, EventArgs e)
		{
			ImportFromXml();
			ShowFormInfo();
		}

		private void btnExportHomeInfo_Click(object sender, EventArgs e)
		{
			SetFormInfo();
			ExportToStringXml();
		}

		private void SetFormInfo()
		{
			HomeInfo.MasterName = tbMaster.Text;
			HomeInfo.ChiefMateName = tbChiefMate.Text;
			HomeInfo.DelegateOfficerName = tbDelegateOfficer.Text;
			HomeInfo.WKO1Name = tbWKO1.Text;
			HomeInfo.WKO2Name = tbWKO2.Text;
			HomeInfo.VesselName = tbVessel.Text;
			HomeInfo.PortFrom = tbPortFrom.Text;
			HomeInfo.PortTo = tbPortTo.Text;
			HomeInfo.Voyage = tbVoyage.Text;
			HomeInfo.SafetyCountDepth = tbSafetyContDepth.Text;
			HomeInfo.Ets = tbEts.Text;
			HomeInfo.Eta = tbEta.Text;
			HomeInfo.CreationDate = dateOfCreatingRoute.MinDate;

			if (double.TryParse(tbDraftFwd.Text, out double draftFwd))
			{
				HomeInfo.draftFWD = draftFwd;
			}
			else
			{
				MessageBox.Show($"Please enter valid number for draft");
			}

			if (double.TryParse(tbDraftAft.Text, out double draftAft))
			{
				HomeInfo.draftAFT = draftAft;
			}
			else
			{
				MessageBox.Show($"Please enter valid number for draft");
			}
		}

		private void ShowFormInfo()
		{
			tbMaster.Text = HomeInfo.MasterName;
			tbChiefMate.Text = HomeInfo.ChiefMateName;
			tbDelegateOfficer.Text = HomeInfo.DelegateOfficerName;
			tbWKO1.Text = HomeInfo.WKO1Name;
			tbWKO2.Text = HomeInfo.WKO2Name;
			tbVessel.Text = HomeInfo.VesselName;
			tbPortFrom.Text = HomeInfo.PortFrom;
			tbPortTo.Text = HomeInfo.PortTo;
			tbVoyage.Text = HomeInfo.Voyage;
			tbSafetyContDepth.Text = HomeInfo.SafetyCountDepth;
			tbEts.Text = HomeInfo.Ets;
			tbEta.Text = HomeInfo.Eta;
			//Skip show creation date
			tbDraftFwd.Text = HomeInfo.draftFWD.ToString("f1");
			tbDraftAft.Text = HomeInfo.draftAFT.ToString("f1");
		}

		private void ExportToStringXml()
		{
			string homeInfoXml = new XmlHelper().Serialize(HomeInfo, "homeInfo");
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

						HomeInfo = new XmlHelper()
							.Deserialize<HomeInfoDto>(xmlStirng, "homeInfo");
					}
				}
				catch (Exception ex)
				{
					MessageBox.Show("Error while loading the file: " + ex.Message);
				}
			}
		}

		
	}
}
