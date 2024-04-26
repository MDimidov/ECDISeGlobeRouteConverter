namespace ECDIS_eGloebe___RouteConverter
{
	partial class Form1
	{
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		/// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing && (components != null))
			{
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region Windows Form Designer generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.master = new System.Windows.Forms.TextBox();
			this.chiefMate = new System.Windows.Forms.TextBox();
			this.wko1 = new System.Windows.Forms.TextBox();
			this.wko2 = new System.Windows.Forms.TextBox();
			this.delegateOfficer = new System.Windows.Forms.TextBox();
			this.draftFWD = new System.Windows.Forms.TextBox();
			this.draftAFT = new System.Windows.Forms.TextBox();
			this.SuspendLayout();
			// 
			// master
			// 
			this.master.AccessibleName = "";
			this.master.Location = new System.Drawing.Point(42, 36);
			this.master.Name = "master";
			this.master.Size = new System.Drawing.Size(196, 20);
			this.master.TabIndex = 1;
			this.master.Text = "Master Name";
			// 
			// chiefMate
			// 
			this.chiefMate.AccessibleName = "";
			this.chiefMate.Location = new System.Drawing.Point(42, 62);
			this.chiefMate.Name = "chiefMate";
			this.chiefMate.Size = new System.Drawing.Size(196, 20);
			this.chiefMate.TabIndex = 2;
			this.chiefMate.Text = "Chief Mate Name";
			// 
			// wko1
			// 
			this.wko1.AccessibleName = "";
			this.wko1.Location = new System.Drawing.Point(42, 114);
			this.wko1.Name = "wko1";
			this.wko1.Size = new System.Drawing.Size(196, 20);
			this.wko1.TabIndex = 3;
			this.wko1.Text = "WKO1 Name";
			this.wko1.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
			// 
			// wko2
			// 
			this.wko2.AccessibleName = "";
			this.wko2.Location = new System.Drawing.Point(42, 140);
			this.wko2.Name = "wko2";
			this.wko2.Size = new System.Drawing.Size(196, 20);
			this.wko2.TabIndex = 4;
			this.wko2.Text = "WKO2 Name";
			// 
			// delegateOfficer
			// 
			this.delegateOfficer.AccessibleName = "";
			this.delegateOfficer.Location = new System.Drawing.Point(42, 88);
			this.delegateOfficer.Name = "delegateOfficer";
			this.delegateOfficer.Size = new System.Drawing.Size(196, 20);
			this.delegateOfficer.TabIndex = 5;
			this.delegateOfficer.Text = "Delegate Officer Name";
			this.delegateOfficer.TextChanged += new System.EventHandler(this.textBox1_TextChanged_1);
			// 
			// draftFWD
			// 
			this.draftFWD.AccessibleName = "";
			this.draftFWD.Location = new System.Drawing.Point(42, 176);
			this.draftFWD.Name = "draftFWD";
			this.draftFWD.Size = new System.Drawing.Size(196, 20);
			this.draftFWD.TabIndex = 6;
			this.draftFWD.Text = "Draft FWD";
			// 
			// draftAFT
			// 
			this.draftAFT.AccessibleName = "";
			this.draftAFT.Location = new System.Drawing.Point(42, 202);
			this.draftAFT.Name = "draftAFT";
			this.draftAFT.Size = new System.Drawing.Size(196, 20);
			this.draftAFT.TabIndex = 7;
			this.draftAFT.Text = "Draft AFT";
			// 
			// Form1
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(797, 442);
			this.Controls.Add(this.draftAFT);
			this.Controls.Add(this.draftFWD);
			this.Controls.Add(this.delegateOfficer);
			this.Controls.Add(this.wko2);
			this.Controls.Add(this.wko1);
			this.Controls.Add(this.chiefMate);
			this.Controls.Add(this.master);
			this.Name = "Form1";
			this.Text = "Form1";
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion
		private System.Windows.Forms.TextBox master;
		private System.Windows.Forms.TextBox chiefMate;
		private System.Windows.Forms.TextBox wko1;
		private System.Windows.Forms.TextBox wko2;
		private System.Windows.Forms.TextBox delegateOfficer;
		private System.Windows.Forms.TextBox draftFWD;
		private System.Windows.Forms.TextBox draftAFT;
	}
}

