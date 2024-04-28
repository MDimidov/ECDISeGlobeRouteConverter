using System;
using System.Windows.Forms;

namespace ECDIS_eGloebe___RouteConverter
{
	public partial class EcdisForm : Form
	{
		public EcdisForm()
		{
			InitializeComponent();
			sidePanel.Top = btnHome.Top;
			sidePanel.Height = btnHome.Height;
			homeControl1.BringToFront();

		}

		private void textBox1_TextChanged(object sender, EventArgs e)
		{
			
		}

		private void textBox1_TextChanged_1(object sender, EventArgs e)
		{

		}

		private void btnHome_Click(object sender, EventArgs e)
		{
			sidePanel.Top = btnHome.Top;
			sidePanel.Height = btnHome.Height;
			homeControl1.BringToFront();
		}

		private void btnExit_Click(object sender, EventArgs e)
		{
			//MessageBox.Show("Do you want to exit this program?");
			Application.Exit();
		}

		private void btnExport_Click(object sender, EventArgs e)
		{
			sidePanel.Top = btnExport.Top;
			sidePanel.Height = btnExport.Height;
			exportControl1.BringToFront();
		}

		private void btnImport_Click(object sender, EventArgs e)
		{
			sidePanel.Top = btnImport.Top;
			sidePanel.Height = btnImport.Height;
		}
	}
}
