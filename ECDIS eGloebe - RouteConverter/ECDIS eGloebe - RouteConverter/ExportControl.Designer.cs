namespace ECDIS_eGloebe___RouteConverter
{
	partial class ExportControl
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
			this.lbExportHomeInfo = new System.Windows.Forms.Label();
			this.btnExportHomeInfo = new System.Windows.Forms.Button();
			this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			this.SuspendLayout();
			// 
			// lbExportHomeInfo
			// 
			this.lbExportHomeInfo.AutoSize = true;
			this.lbExportHomeInfo.Font = new System.Drawing.Font("Century Gothic", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lbExportHomeInfo.Location = new System.Drawing.Point(51, 38);
			this.lbExportHomeInfo.Name = "lbExportHomeInfo";
			this.lbExportHomeInfo.Size = new System.Drawing.Size(216, 19);
			this.lbExportHomeInfo.TabIndex = 0;
			this.lbExportHomeInfo.Text = "Export Passage Plan to .xlsx";
			// 
			// btnExportHomeInfo
			// 
			this.btnExportHomeInfo.Location = new System.Drawing.Point(79, 83);
			this.btnExportHomeInfo.Name = "btnExportHomeInfo";
			this.btnExportHomeInfo.Size = new System.Drawing.Size(165, 56);
			this.btnExportHomeInfo.TabIndex = 1;
			this.btnExportHomeInfo.Text = "Export";
			this.btnExportHomeInfo.UseVisualStyleBackColor = true;
			this.btnExportHomeInfo.Click += new System.EventHandler(this.btnExportHomeInfo_Click);
			// 
			// ExportControl
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.Controls.Add(this.btnExportHomeInfo);
			this.Controls.Add(this.lbExportHomeInfo);
			this.Name = "ExportControl";
			this.Size = new System.Drawing.Size(872, 552);
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.Label lbExportHomeInfo;
		private System.Windows.Forms.Button btnExportHomeInfo;
		private System.Windows.Forms.SaveFileDialog saveFileDialog1;
	}
}
