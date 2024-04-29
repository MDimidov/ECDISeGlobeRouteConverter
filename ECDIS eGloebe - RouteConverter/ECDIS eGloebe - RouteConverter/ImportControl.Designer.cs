namespace ECDIS_eGloebe___RouteConverter
{
	partial class ImportControl
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

		#region Component Designer generated code

		/// <summary> 
		/// Required method for Designer support - do not modify 
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.btnImportRoute = new System.Windows.Forms.Button();
			this.SuspendLayout();
			// 
			// btnImportRoute
			// 
			this.btnImportRoute.Location = new System.Drawing.Point(38, 63);
			this.btnImportRoute.Name = "btnImportRoute";
			this.btnImportRoute.Size = new System.Drawing.Size(245, 69);
			this.btnImportRoute.TabIndex = 37;
			this.btnImportRoute.Text = "Import route file";
			this.btnImportRoute.UseVisualStyleBackColor = true;
			this.btnImportRoute.Click += new System.EventHandler(this.btnImportRoute_Click);
			// 
			// ImportControl
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.Controls.Add(this.btnImportRoute);
			this.Name = "ImportControl";
			this.Size = new System.Drawing.Size(872, 552);
			this.ResumeLayout(false);

		}

		#endregion

		private System.Windows.Forms.Button btnImportRoute;
	}
}
