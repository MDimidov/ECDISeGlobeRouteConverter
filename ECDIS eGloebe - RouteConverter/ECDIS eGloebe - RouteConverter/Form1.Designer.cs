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
			this.leftMenu = new System.Windows.Forms.Panel();
			this.sidePanel = new System.Windows.Forms.Panel();
			this.btnImport = new System.Windows.Forms.Button();
			this.btnHome = new System.Windows.Forms.Button();
			this.header = new System.Windows.Forms.Panel();
			this.button2 = new System.Windows.Forms.Button();
			this.mainPanel = new System.Windows.Forms.Panel();
			this.homeControl1 = new ECDIS_eGloebe___RouteConverter.HomeControl();
			this.leftMenu.SuspendLayout();
			this.header.SuspendLayout();
			this.mainPanel.SuspendLayout();
			this.SuspendLayout();
			// 
			// leftMenu
			// 
			this.leftMenu.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(41)))), ((int)(((byte)(39)))), ((int)(((byte)(40)))));
			this.leftMenu.Controls.Add(this.sidePanel);
			this.leftMenu.Controls.Add(this.btnImport);
			this.leftMenu.Controls.Add(this.btnHome);
			this.leftMenu.Dock = System.Windows.Forms.DockStyle.Left;
			this.leftMenu.Location = new System.Drawing.Point(0, 0);
			this.leftMenu.Name = "leftMenu";
			this.leftMenu.Size = new System.Drawing.Size(191, 584);
			this.leftMenu.TabIndex = 9;
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
			this.btnImport.Location = new System.Drawing.Point(0, 136);
			this.btnImport.Name = "btnImport";
			this.btnImport.Size = new System.Drawing.Size(190, 62);
			this.btnImport.TabIndex = 1;
			this.btnImport.Text = "Import file";
			this.btnImport.UseCompatibleTextRendering = true;
			this.btnImport.UseVisualStyleBackColor = false;
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
			// header
			// 
			this.header.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(178)))), ((int)(((byte)(8)))), ((int)(((byte)(55)))));
			this.header.Controls.Add(this.button2);
			this.header.Dock = System.Windows.Forms.DockStyle.Top;
			this.header.Location = new System.Drawing.Point(191, 0);
			this.header.Name = "header";
			this.header.Size = new System.Drawing.Size(872, 32);
			this.header.TabIndex = 10;
			// 
			// button2
			// 
			this.button2.FlatAppearance.BorderSize = 0;
			this.button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
			this.button2.Font = new System.Drawing.Font("Century", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
			this.button2.ForeColor = System.Drawing.SystemColors.ButtonFace;
			this.button2.Location = new System.Drawing.Point(843, 0);
			this.button2.Name = "button2";
			this.button2.Size = new System.Drawing.Size(26, 29);
			this.button2.TabIndex = 0;
			this.button2.Text = "X";
			this.button2.UseVisualStyleBackColor = true;
			// 
			// mainPanel
			// 
			this.mainPanel.Controls.Add(this.homeControl1);
			this.mainPanel.Dock = System.Windows.Forms.DockStyle.Fill;
			this.mainPanel.Location = new System.Drawing.Point(191, 32);
			this.mainPanel.Name = "mainPanel";
			this.mainPanel.Size = new System.Drawing.Size(872, 552);
			this.mainPanel.TabIndex = 11;
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
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(1063, 584);
			this.Controls.Add(this.mainPanel);
			this.Controls.Add(this.header);
			this.Controls.Add(this.leftMenu);
			this.Name = "EcdisForm";
			this.Text = "eGlobe Passage Plan Generator";
			this.leftMenu.ResumeLayout(false);
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
		private System.Windows.Forms.Button button2;
		private System.Windows.Forms.Panel mainPanel;
		private HomeControl homeControl1;
	}
}

