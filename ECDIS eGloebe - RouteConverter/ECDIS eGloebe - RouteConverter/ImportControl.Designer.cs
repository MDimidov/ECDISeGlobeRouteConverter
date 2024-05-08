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
			this.lbImportRte = new System.Windows.Forms.Label();
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
			// lbImportRte
			// 
			this.lbImportRte.AutoSize = true;
			this.lbImportRte.Font = new System.Drawing.Font("Century Gothic", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lbImportRte.Location = new System.Drawing.Point(79, 41);
			this.lbImportRte.Name = "lbImportRte";
			this.lbImportRte.Size = new System.Drawing.Size(175, 19);
			this.lbImportRte.TabIndex = 38;
			this.lbImportRte.Text = "Import Route File (.rte)";
			// 
			// ImportControl
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.Controls.Add(this.lbImportRte);
			this.Controls.Add(this.btnImportRoute);
			this.Name = "ImportControl";
			this.Size = new System.Drawing.Size(872, 552);
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.Button btnImportRoute;
		private System.Windows.Forms.Label lbImportRte;
	}
}
