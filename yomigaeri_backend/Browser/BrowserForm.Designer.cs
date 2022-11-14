
namespace yomigaeri_backend.Browser
{
	partial class BrowserForm
	{
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		#region Windows Form Designer generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.components = new System.ComponentModel.Container();
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(BrowserForm));
			this.VirtualChannelTimer = new System.Windows.Forms.Timer(this.components);
			this.SuspendLayout();
			// 
			// VirtualChannelTimer
			// 
			this.VirtualChannelTimer.Enabled = true;
			this.VirtualChannelTimer.Interval = 500;
			this.VirtualChannelTimer.Tick += new System.EventHandler(this.VirtualChannelTimer_Tick);
			// 
			// BrowserForm
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.BackColor = System.Drawing.Color.White;
			this.ClientSize = new System.Drawing.Size(250, 250);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.Name = "BrowserForm";
			this.ShowIcon = false;
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
			this.Text = "IE6YG Browser Window";
			this.Shown += new System.EventHandler(this.BrowserForm_Shown);
			this.ResizeBegin += new System.EventHandler(this.BrowserForm_ResizeBegin);
			this.ResizeEnd += new System.EventHandler(this.BrowserForm_ResizeEnd);
			this.ResumeLayout(false);

		}

		#endregion

		private System.Windows.Forms.Timer VirtualChannelTimer;
	}
}