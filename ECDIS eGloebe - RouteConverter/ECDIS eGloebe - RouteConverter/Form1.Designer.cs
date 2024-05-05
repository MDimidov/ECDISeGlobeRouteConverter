namespace ECDIS_eGloebe___RouteConverter
{
	partial class EcdisForm
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
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(EcdisForm));
			this.leftMenu = new System.Windows.Forms.Panel();
			this.creator = new System.Windows.Forms.Label();
			this.sidePanel = new System.Windows.Forms.Panel();
			this.btnImport = new System.Windows.Forms.Button();
			this.btnHome = new System.Windows.Forms.Button();
			this.btnExport = new System.Windows.Forms.Button();
			this.header = new System.Windows.Forms.Panel();
			this.btnExit = new System.Windows.Forms.Button();
			this.mainPanel = new System.Windows.Forms.Panel();
			this.importControl1 = new ECDIS_eGloebe___RouteConverter.ImportControl();
			this.exportControl1 = new ECDIS_eGloebe___RouteConverter.ExportControl();
			this.homeControl1 = new ECDIS_eGloebe___RouteConverter.HomeControl();
			this.leftMenu.SuspendLayout();
			this.header.SuspendLayout();
			this.mainPanel.SuspendLayout();
			this.SuspendLayout();
			// 
			// leftMenu
			// 
			this.leftMenu.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(41)))), ((int)(((byte)(39)))), ((int)(((byte)(40)))));
			this.leftMenu.Controls.Add(this.creator);
			this.leftMenu.Controls.Add(this.sidePanel);
			this.leftMenu.Controls.Add(this.btnImport);
			this.leftMenu.Controls.Add(this.btnHome);
			this.leftMenu.Controls.Add(this.btnExport);
			this.leftMenu.Dock = System.Windows.Forms.DockStyle.Left;
			this.leftMenu.Location = new System.Drawing.Point(0, 0);
			this.leftMenu.Name = "leftMenu";
			this.leftMenu.Size = new System.Drawing.Size(191, 584);
			this.leftMenu.TabIndex = 9;
			// 
			// creator
			// 
			this.creator.AutoSize = true;
			this.creator.Font = new System.Drawing.Font("Comic Sans MS", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(204)));
			this.creator.ForeColor = System.Drawing.SystemColors.ButtonShadow;
			this.creator.Location = new System.Drawing.Point(3, 566);
			this.creator.Name = "creator";
			this.creator.Size = new System.Drawing.Size(116, 15);
			this.creator.TabIndex = 12;
			this.creator.Text = "Designed by Dimidov";
			// 
			// sidePanel
			// 
			this.sidePanel.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(178)))), ((int)(((byte)(8)))), ((int)(((byte)(55)))));
			this.sidePanel.Location = new System.Drawing.Point(1, 65);
			this.sidePanel.Name = "sidePanel";
			this.sidePanel.Size = new System.Drawing.Size(10, 62);
			this.sidePanel.TabIndex = 10;
			// 
			// btnImport
			// 
			this.btnImport.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(41)))), ((int)(((byte)(39)))), ((int)(((byte)(40)))));
			this.btnImport.FlatAppearance.BorderSize = 0;
			this.btnImport.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
			this.btnImport.Font = new System.Drawing.Font("Century Gothic", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
			this.btnImport.ForeColor = System.Drawing.SystemColors.ButtonFace;
			this.btnImport.Location = new System.Drawing.Point(0, 134);
			this.btnImport.Name = "btnImport";
			this.btnImport.Size = new System.Drawing.Size(190, 62);
			this.btnImport.TabIndex = 1;
			this.btnImport.Text = "Import file";
			this.btnImport.UseCompatibleTextRendering = true;
			this.btnImport.UseVisualStyleBackColor = false;
			this.btnImport.Click += new System.EventHandler(this.btnImport_Click);
			// 
			// btnHome
			// 
			this.btnHome.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(41)))), ((int)(((byte)(39)))), ((int)(((byte)(40)))));
			this.btnHome.FlatAppearance.BorderSize = 0;
			this.btnHome.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
			this.btnHome.Font = new System.Drawing.Font("Century Gothic", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
			this.btnHome.ForeColor = System.Drawing.SystemColors.ButtonFace;
			this.btnHome.Location = new System.Drawing.Point(0, 65);
			this.btnHome.Name = "btnHome";
			this.btnHome.Size = new System.Drawing.Size(190, 62);
			this.btnHome.TabIndex = 0;
			this.btnHome.Text = "Home";
			this.btnHome.UseCompatibleTextRendering = true;
			this.btnHome.UseVisualStyleBackColor = false;
			this.btnHome.Click += new System.EventHandler(this.btnHome_Click);
			// 
			// btnExport
			// 
			this.btnExport.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(41)))), ((int)(((byte)(39)))), ((int)(((byte)(40)))));
			this.btnExport.FlatAppearance.BorderSize = 0;
			this.btnExport.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
			this.btnExport.Font = new System.Drawing.Font("Century Gothic", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
			this.btnExport.ForeColor = System.Drawing.SystemColors.ButtonFace;
			this.btnExport.Location = new System.Drawing.Point(0, 203);
			this.btnExport.Name = "btnExport";
			this.btnExport.Size = new System.Drawing.Size(190, 62);
			this.btnExport.TabIndex = 11;
			this.btnExport.Text = "Export file";
			this.btnExport.UseCompatibleTextRendering = true;
			this.btnExport.UseVisualStyleBackColor = false;
			this.btnExport.Click += new System.EventHandler(this.btnExport_Click);
			// 
			// header
			// 
			this.header.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(178)))), ((int)(((byte)(8)))), ((int)(((byte)(55)))));
			this.header.Controls.Add(this.btnExit);
			this.header.Dock = System.Windows.Forms.DockStyle.Top;
			this.header.Location = new System.Drawing.Point(191, 0);
			this.header.Name = "header";
			this.header.Size = new System.Drawing.Size(872, 32);
			this.header.TabIndex = 10;
			// 
			// btnExit
			// 
			this.btnExit.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExit.FlatAppearance.BorderSize = 0;
			this.btnExit.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
			this.btnExit.Font = new System.Drawing.Font("Century", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
			this.btnExit.ForeColor = System.Drawing.SystemColors.ButtonFace;
			this.btnExit.Location = new System.Drawing.Point(843, 0);
			this.btnExit.Name = "btnExit";
			this.btnExit.Size = new System.Drawing.Size(26, 29);
			this.btnExit.TabIndex = 0;
			this.btnExit.Text = "X";
			this.btnExit.UseVisualStyleBackColor = true;
			this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
			// 
			// mainPanel
			// 
			this.mainPanel.Controls.Add(this.importControl1);
			this.mainPanel.Controls.Add(this.exportControl1);
			this.mainPanel.Controls.Add(this.homeControl1);
			this.mainPanel.Dock = System.Windows.Forms.DockStyle.Fill;
			this.mainPanel.Location = new System.Drawing.Point(191, 32);
			this.mainPanel.Name = "mainPanel";
			this.mainPanel.Size = new System.Drawing.Size(872, 552);
			this.mainPanel.TabIndex = 11;
			// 
			// importControl1
			// 
			this.importControl1.Location = new System.Drawing.Point(0, 0);
			this.importControl1.Name = "importControl1";
			this.importControl1.Size = new System.Drawing.Size(872, 552);
			this.importControl1.TabIndex = 2;
			// 
			// exportControl1
			// 
			this.exportControl1.Location = new System.Drawing.Point(0, 0);
			this.exportControl1.Name = "exportControl1";
			this.exportControl1.Size = new System.Drawing.Size(872, 552);
			this.exportControl1.TabIndex = 1;
			// 
			// homeControl1
			// 
			this.homeControl1.Location = new System.Drawing.Point(0, 0);
			this.homeControl1.Name = "homeControl1";
			this.homeControl1.Size = new System.Drawing.Size(870, 550);
			this.homeControl1.TabIndex = 0;
			// 
			// EcdisForm
			// 
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
			this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
			this.ClientSize = new System.Drawing.Size(1063, 584);
			this.Controls.Add(this.mainPanel);
			this.Controls.Add(this.header);
			this.Controls.Add(this.leftMenu);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.Name = "EcdisForm";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "eGlobe Passage Plan Generator";
			this.leftMenu.ResumeLayout(false);
			this.leftMenu.PerformLayout();
			this.header.ResumeLayout(false);
			this.mainPanel.ResumeLayout(false);
			this.ResumeLayout(false);

		}

		#endregion
		private System.Windows.Forms.Panel leftMenu;
		private System.Windows.Forms.Button btnImport;
		private System.Windows.Forms.Button btnHome;
		private System.Windows.Forms.Panel sidePanel;
		private System.Windows.Forms.Panel header;
		private System.Windows.Forms.Button btnExit;
		private System.Windows.Forms.Panel mainPanel;
		private HomeControl homeControl1;
		private System.Windows.Forms.Button btnExport;
		private ExportControl exportControl1;
		private System.Windows.Forms.Label creator;
		private ImportControl importControl1;
	}
}

