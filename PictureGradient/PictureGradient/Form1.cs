using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

namespace PictureGradient
{
	/// <summary>
	/// Summary description for Form1.
	/// </summary>
	public class Form1 : System.Windows.Forms.Form
	{
		private System.Windows.Forms.PictureBox pctGradient;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.HScrollBar scrWidth;
		private System.Windows.Forms.HScrollBar scrHeight;
		private System.Windows.Forms.HScrollBar scrRedTop;
		private System.Windows.Forms.HScrollBar scrGreenTop;
		private System.Windows.Forms.HScrollBar scrBlueTop;
		private System.Windows.Forms.HScrollBar scrRedBottom;
		private System.Windows.Forms.HScrollBar scrGreenBottom;
		private System.Windows.Forms.HScrollBar scrBlueBottom;
		private System.Windows.Forms.TextBox txtHeight;
		private System.Windows.Forms.TextBox txtRedTop;
		private System.Windows.Forms.TextBox txtGreenTop;
		private System.Windows.Forms.TextBox txtBlueTop;
		private System.Windows.Forms.TextBox txtRedBottom;
		private System.Windows.Forms.TextBox txtGreenBottom;
		private System.Windows.Forms.TextBox txtBlueBottom;
		private System.Windows.Forms.Button cmdGenerateFile;
		private System.Windows.Forms.TextBox txtWidth;
		private System.Windows.Forms.TextBox txtAlphaTop;
		private System.Windows.Forms.HScrollBar scrAlphaBottom;
		private System.Windows.Forms.HScrollBar scrAlphaTop;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.TextBox txtAlphaBottom;
		private System.Windows.Forms.CheckBox chkHorizontal;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public Form1()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if (components != null) 
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(Form1));
			this.pctGradient = new System.Windows.Forms.PictureBox();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.scrWidth = new System.Windows.Forms.HScrollBar();
			this.scrHeight = new System.Windows.Forms.HScrollBar();
			this.scrRedTop = new System.Windows.Forms.HScrollBar();
			this.scrGreenTop = new System.Windows.Forms.HScrollBar();
			this.scrBlueTop = new System.Windows.Forms.HScrollBar();
			this.scrRedBottom = new System.Windows.Forms.HScrollBar();
			this.scrGreenBottom = new System.Windows.Forms.HScrollBar();
			this.scrBlueBottom = new System.Windows.Forms.HScrollBar();
			this.txtHeight = new System.Windows.Forms.TextBox();
			this.txtRedTop = new System.Windows.Forms.TextBox();
			this.txtGreenTop = new System.Windows.Forms.TextBox();
			this.txtBlueTop = new System.Windows.Forms.TextBox();
			this.txtRedBottom = new System.Windows.Forms.TextBox();
			this.txtGreenBottom = new System.Windows.Forms.TextBox();
			this.txtBlueBottom = new System.Windows.Forms.TextBox();
			this.cmdGenerateFile = new System.Windows.Forms.Button();
			this.txtWidth = new System.Windows.Forms.TextBox();
			this.txtAlphaTop = new System.Windows.Forms.TextBox();
			this.scrAlphaBottom = new System.Windows.Forms.HScrollBar();
			this.scrAlphaTop = new System.Windows.Forms.HScrollBar();
			this.label6 = new System.Windows.Forms.Label();
			this.txtAlphaBottom = new System.Windows.Forms.TextBox();
			this.chkHorizontal = new System.Windows.Forms.CheckBox();
			this.SuspendLayout();
			// 
			// pctGradient
			// 
			this.pctGradient.BackColor = System.Drawing.Color.Transparent;
			this.pctGradient.BackgroundImage = ((System.Drawing.Bitmap)(resources.GetObject("pctGradient.BackgroundImage")));
			this.pctGradient.Image = ((System.Drawing.Bitmap)(resources.GetObject("pctGradient.Image")));
			this.pctGradient.Location = new System.Drawing.Point(8, 112);
			this.pctGradient.Name = "pctGradient";
			this.pctGradient.Size = new System.Drawing.Size(192, 176);
			this.pctGradient.TabIndex = 1;
			this.pctGradient.TabStop = false;
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(8, 8);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(80, 16);
			this.label1.TabIndex = 2;
			this.label1.Text = "Width";
			// 
			// label2
			// 
			this.label2.Location = new System.Drawing.Point(8, 32);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(80, 16);
			this.label2.TabIndex = 3;
			this.label2.Text = "Height";
			// 
			// label3
			// 
			this.label3.Location = new System.Drawing.Point(376, 8);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(80, 16);
			this.label3.TabIndex = 4;
			this.label3.Text = "Red";
			// 
			// label4
			// 
			this.label4.Location = new System.Drawing.Point(376, 32);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(80, 16);
			this.label4.TabIndex = 5;
			this.label4.Text = "Green";
			// 
			// label5
			// 
			this.label5.Location = new System.Drawing.Point(376, 56);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(80, 16);
			this.label5.TabIndex = 6;
			this.label5.Text = "Blue";
			// 
			// scrWidth
			// 
			this.scrWidth.Location = new System.Drawing.Point(88, 8);
			this.scrWidth.Maximum = 1289;
			this.scrWidth.Minimum = 1;
			this.scrWidth.Name = "scrWidth";
			this.scrWidth.Size = new System.Drawing.Size(184, 16);
			this.scrWidth.TabIndex = 7;
			this.scrWidth.Value = 100;
			this.scrWidth.Scroll += new System.Windows.Forms.ScrollEventHandler(this.scrWidth_Scroll);
			// 
			// scrHeight
			// 
			this.scrHeight.Location = new System.Drawing.Point(88, 32);
			this.scrHeight.Maximum = 1033;
			this.scrHeight.Minimum = 1;
			this.scrHeight.Name = "scrHeight";
			this.scrHeight.Size = new System.Drawing.Size(184, 16);
			this.scrHeight.TabIndex = 8;
			this.scrHeight.Value = 100;
			this.scrHeight.Scroll += new System.Windows.Forms.ScrollEventHandler(this.scrHeight_Scroll);
			// 
			// scrRedTop
			// 
			this.scrRedTop.Location = new System.Drawing.Point(456, 8);
			this.scrRedTop.Maximum = 264;
			this.scrRedTop.Name = "scrRedTop";
			this.scrRedTop.Size = new System.Drawing.Size(88, 16);
			this.scrRedTop.TabIndex = 9;
			this.scrRedTop.Scroll += new System.Windows.Forms.ScrollEventHandler(this.scrRedTop_Scroll);
			// 
			// scrGreenTop
			// 
			this.scrGreenTop.Location = new System.Drawing.Point(456, 32);
			this.scrGreenTop.Maximum = 264;
			this.scrGreenTop.Name = "scrGreenTop";
			this.scrGreenTop.Size = new System.Drawing.Size(88, 16);
			this.scrGreenTop.TabIndex = 10;
			this.scrGreenTop.Scroll += new System.Windows.Forms.ScrollEventHandler(this.scrGreenTop_Scroll);
			// 
			// scrBlueTop
			// 
			this.scrBlueTop.Location = new System.Drawing.Point(456, 56);
			this.scrBlueTop.Maximum = 264;
			this.scrBlueTop.Name = "scrBlueTop";
			this.scrBlueTop.Size = new System.Drawing.Size(88, 16);
			this.scrBlueTop.TabIndex = 11;
			this.scrBlueTop.Scroll += new System.Windows.Forms.ScrollEventHandler(this.scrBlueTop_Scroll);
			// 
			// scrRedBottom
			// 
			this.scrRedBottom.Location = new System.Drawing.Point(552, 8);
			this.scrRedBottom.Maximum = 264;
			this.scrRedBottom.Name = "scrRedBottom";
			this.scrRedBottom.Size = new System.Drawing.Size(88, 16);
			this.scrRedBottom.TabIndex = 12;
			this.scrRedBottom.Value = 255;
			this.scrRedBottom.Scroll += new System.Windows.Forms.ScrollEventHandler(this.scrRedBottom_Scroll);
			// 
			// scrGreenBottom
			// 
			this.scrGreenBottom.Location = new System.Drawing.Point(552, 32);
			this.scrGreenBottom.Maximum = 264;
			this.scrGreenBottom.Name = "scrGreenBottom";
			this.scrGreenBottom.Size = new System.Drawing.Size(88, 16);
			this.scrGreenBottom.TabIndex = 13;
			this.scrGreenBottom.Value = 255;
			this.scrGreenBottom.Scroll += new System.Windows.Forms.ScrollEventHandler(this.scrGreenBottom_Scroll);
			// 
			// scrBlueBottom
			// 
			this.scrBlueBottom.Location = new System.Drawing.Point(552, 56);
			this.scrBlueBottom.Maximum = 264;
			this.scrBlueBottom.Name = "scrBlueBottom";
			this.scrBlueBottom.Size = new System.Drawing.Size(88, 16);
			this.scrBlueBottom.TabIndex = 14;
			this.scrBlueBottom.Value = 255;
			this.scrBlueBottom.Scroll += new System.Windows.Forms.ScrollEventHandler(this.hScrollBar8_Scroll);
			// 
			// txtHeight
			// 
			this.txtHeight.Location = new System.Drawing.Point(296, 32);
			this.txtHeight.Name = "txtHeight";
			this.txtHeight.Size = new System.Drawing.Size(64, 20);
			this.txtHeight.TabIndex = 16;
			this.txtHeight.Text = "100";
			this.txtHeight.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtHeight_KeyPress);
			this.txtHeight.Leave += new System.EventHandler(this.txtHeight_Leave);
			// 
			// txtRedTop
			// 
			this.txtRedTop.Location = new System.Drawing.Point(664, 8);
			this.txtRedTop.Name = "txtRedTop";
			this.txtRedTop.Size = new System.Drawing.Size(64, 20);
			this.txtRedTop.TabIndex = 17;
			this.txtRedTop.Text = "0";
			this.txtRedTop.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtRedTop_KeyPress);
			this.txtRedTop.Leave += new System.EventHandler(this.txtRedTop_Leave);
			// 
			// txtGreenTop
			// 
			this.txtGreenTop.Location = new System.Drawing.Point(664, 32);
			this.txtGreenTop.Name = "txtGreenTop";
			this.txtGreenTop.Size = new System.Drawing.Size(64, 20);
			this.txtGreenTop.TabIndex = 18;
			this.txtGreenTop.Text = "0";
			this.txtGreenTop.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtGreenTop_KeyPress);
			this.txtGreenTop.Leave += new System.EventHandler(this.txtGreenTop_Leave);
			// 
			// txtBlueTop
			// 
			this.txtBlueTop.Location = new System.Drawing.Point(664, 56);
			this.txtBlueTop.Name = "txtBlueTop";
			this.txtBlueTop.Size = new System.Drawing.Size(64, 20);
			this.txtBlueTop.TabIndex = 19;
			this.txtBlueTop.Text = "0";
			this.txtBlueTop.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtBlueTop_KeyPress);
			this.txtBlueTop.Leave += new System.EventHandler(this.txtBlueTop_Leave);
			// 
			// txtRedBottom
			// 
			this.txtRedBottom.Location = new System.Drawing.Point(736, 8);
			this.txtRedBottom.Name = "txtRedBottom";
			this.txtRedBottom.Size = new System.Drawing.Size(64, 20);
			this.txtRedBottom.TabIndex = 20;
			this.txtRedBottom.Text = "255";
			this.txtRedBottom.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtRedBottom_KeyPress);
			this.txtRedBottom.Leave += new System.EventHandler(this.txtRedBottom_Leave);
			// 
			// txtGreenBottom
			// 
			this.txtGreenBottom.Location = new System.Drawing.Point(736, 32);
			this.txtGreenBottom.Name = "txtGreenBottom";
			this.txtGreenBottom.Size = new System.Drawing.Size(64, 20);
			this.txtGreenBottom.TabIndex = 21;
			this.txtGreenBottom.Text = "255";
			this.txtGreenBottom.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtGreenBottom_KeyPress);
			this.txtGreenBottom.Leave += new System.EventHandler(this.txtGreenBottom_Leave);
			// 
			// txtBlueBottom
			// 
			this.txtBlueBottom.Location = new System.Drawing.Point(736, 56);
			this.txtBlueBottom.Name = "txtBlueBottom";
			this.txtBlueBottom.Size = new System.Drawing.Size(64, 20);
			this.txtBlueBottom.TabIndex = 22;
			this.txtBlueBottom.Text = "255";
			this.txtBlueBottom.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtBlueBottom_KeyPress);
			this.txtBlueBottom.Leave += new System.EventHandler(this.txtBlueBottom_Leave);
			// 
			// cmdGenerateFile
			// 
			this.cmdGenerateFile.Location = new System.Drawing.Point(8, 56);
			this.cmdGenerateFile.Name = "cmdGenerateFile";
			this.cmdGenerateFile.Size = new System.Drawing.Size(96, 40);
			this.cmdGenerateFile.TabIndex = 23;
			this.cmdGenerateFile.Text = "Generate File";
			this.cmdGenerateFile.Click += new System.EventHandler(this.cmdGenerateFile_Click);
			// 
			// txtWidth
			// 
			this.txtWidth.Location = new System.Drawing.Point(296, 8);
			this.txtWidth.Name = "txtWidth";
			this.txtWidth.Size = new System.Drawing.Size(64, 20);
			this.txtWidth.TabIndex = 15;
			this.txtWidth.Text = "100";
			this.txtWidth.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtWidth_KeyPress);
			this.txtWidth.Leave += new System.EventHandler(this.txtWidth_Leave);
			// 
			// txtAlphaTop
			// 
			this.txtAlphaTop.Location = new System.Drawing.Point(664, 80);
			this.txtAlphaTop.Name = "txtAlphaTop";
			this.txtAlphaTop.Size = new System.Drawing.Size(64, 20);
			this.txtAlphaTop.TabIndex = 27;
			this.txtAlphaTop.Text = "0";
			this.txtAlphaTop.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtAlphaTop_KeyPress);
			// 
			// scrAlphaBottom
			// 
			this.scrAlphaBottom.Location = new System.Drawing.Point(552, 80);
			this.scrAlphaBottom.Maximum = 264;
			this.scrAlphaBottom.Name = "scrAlphaBottom";
			this.scrAlphaBottom.Size = new System.Drawing.Size(88, 16);
			this.scrAlphaBottom.TabIndex = 26;
			this.scrAlphaBottom.Value = 255;
			this.scrAlphaBottom.Scroll += new System.Windows.Forms.ScrollEventHandler(this.scrAlphaBottom_Scroll);
			// 
			// scrAlphaTop
			// 
			this.scrAlphaTop.Location = new System.Drawing.Point(456, 80);
			this.scrAlphaTop.Maximum = 264;
			this.scrAlphaTop.Name = "scrAlphaTop";
			this.scrAlphaTop.Size = new System.Drawing.Size(88, 16);
			this.scrAlphaTop.TabIndex = 25;
			this.scrAlphaTop.Scroll += new System.Windows.Forms.ScrollEventHandler(this.scrAlphaTop_Scroll);
			// 
			// label6
			// 
			this.label6.Location = new System.Drawing.Point(376, 80);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(80, 16);
			this.label6.TabIndex = 24;
			this.label6.Text = "Alpha";
			// 
			// txtAlphaBottom
			// 
			this.txtAlphaBottom.Location = new System.Drawing.Point(736, 80);
			this.txtAlphaBottom.Name = "txtAlphaBottom";
			this.txtAlphaBottom.Size = new System.Drawing.Size(64, 20);
			this.txtAlphaBottom.TabIndex = 28;
			this.txtAlphaBottom.Text = "255";
			this.txtAlphaBottom.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtAlphaBottom_KeyPress);
			// 
			// chkHorizontal
			// 
			this.chkHorizontal.Location = new System.Drawing.Point(120, 64);
			this.chkHorizontal.Name = "chkHorizontal";
			this.chkHorizontal.Size = new System.Drawing.Size(136, 16);
			this.chkHorizontal.TabIndex = 29;
			this.chkHorizontal.Text = "Horizontal";
			this.chkHorizontal.Click += new System.EventHandler(this.chkHorizontal_Click);
			// 
			// Form1
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(784, 614);
			this.Controls.AddRange(new System.Windows.Forms.Control[] {
																		  this.chkHorizontal,
																		  this.txtAlphaBottom,
																		  this.txtAlphaTop,
																		  this.scrAlphaBottom,
																		  this.scrAlphaTop,
																		  this.label6,
																		  this.cmdGenerateFile,
																		  this.txtBlueBottom,
																		  this.txtGreenBottom,
																		  this.txtRedBottom,
																		  this.txtBlueTop,
																		  this.txtGreenTop,
																		  this.txtRedTop,
																		  this.txtHeight,
																		  this.txtWidth,
																		  this.scrBlueBottom,
																		  this.scrGreenBottom,
																		  this.scrRedBottom,
																		  this.scrBlueTop,
																		  this.scrGreenTop,
																		  this.scrRedTop,
																		  this.scrHeight,
																		  this.scrWidth,
																		  this.label5,
																		  this.label4,
																		  this.label3,
																		  this.label2,
																		  this.label1,
																		  this.pctGradient});
			this.Name = "Form1";
			this.Text = "Form1";
			this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
			this.Load += new System.EventHandler(this.Form1_Load);
			this.ResumeLayout(false);

		}
		#endregion

		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main() 
		{
			Application.Run(new Form1());
		}


		private void cmdGo_Click(object sender, System.EventArgs e)
		{
			RenderGradient();
		}

		private void RenderGradient() 
		{
			Bitmap b = new Bitmap(pctGradient.Width,pctGradient.Height);

			Graphics oCanvas = pctGradient.CreateGraphics();
			pctGradient.Refresh();
			
			if (!chkHorizontal.Checked) 
			{
				for (int iYCoord = 0; iYCoord < pctGradient.Height; iYCoord++) 
				{	
					int iRed = scrRedTop.Value+iYCoord*(scrRedBottom.Value - scrRedTop.Value)/pctGradient.Height;
					int iGreen = scrGreenTop.Value+iYCoord*(scrGreenBottom.Value - scrGreenTop.Value)/pctGradient.Height;
					int iBlue = scrBlueTop.Value+iYCoord*(scrBlueBottom.Value - scrBlueTop.Value)/pctGradient.Height;
					int iAlpha = scrAlphaTop.Value+iYCoord*(scrAlphaBottom.Value - scrAlphaTop.Value)/pctGradient.Height;

					Pen oPen = new System.Drawing.Pen(Color.FromArgb(iAlpha,iRed,iGreen,iBlue));
					oCanvas.DrawLine(oPen, 0, iYCoord, pctGradient.Width, iYCoord);
				}
			} 
			else 
			{
				for (int iXCoord = 0; iXCoord < pctGradient.Width; iXCoord++) 
				{	
					int iRed = scrRedTop.Value+iXCoord*(scrRedBottom.Value - scrRedTop.Value)/pctGradient.Width;
					int iGreen = scrGreenTop.Value+iXCoord*(scrGreenBottom.Value - scrGreenTop.Value)/pctGradient.Width;
					int iBlue = scrBlueTop.Value+iXCoord*(scrBlueBottom.Value - scrBlueTop.Value)/pctGradient.Width;
					int iAlpha = scrAlphaTop.Value+iXCoord*(scrAlphaBottom.Value - scrAlphaTop.Value)/pctGradient.Width;

					Pen oPen = new System.Drawing.Pen(Color.FromArgb(iAlpha,iRed,iGreen,iBlue));
					oCanvas.DrawLine(oPen, iXCoord, 0, iXCoord, pctGradient.Height);
				}
			}
		}

		private void cmdGenerateFile_Click(object sender, System.EventArgs e)
		{
			Bitmap b = new Bitmap(pctGradient.Width,pctGradient.Height);

			Graphics oFile = Graphics.FromImage(b);


			if (!chkHorizontal.Checked) 
			{
				for (int iYCoord = 0; iYCoord < pctGradient.Height; iYCoord++) 
				{	
					int iRed = scrRedTop.Value+iYCoord*(scrRedBottom.Value - scrRedTop.Value)/pctGradient.Height;
					int iGreen = scrGreenTop.Value+iYCoord*(scrGreenBottom.Value - scrGreenTop.Value)/pctGradient.Height;
					int iBlue = scrBlueTop.Value+iYCoord*(scrBlueBottom.Value - scrBlueTop.Value)/pctGradient.Height;
					int iAlpha = scrAlphaTop.Value+iYCoord*(scrAlphaBottom.Value - scrAlphaTop.Value)/pctGradient.Height;

					Pen oPen = new System.Drawing.Pen(Color.FromArgb(iAlpha,iRed,iGreen,iBlue));
					oFile.DrawLine(oPen, 0, iYCoord, pctGradient.Width, iYCoord);
				}
			} 
			else 
			{
				for (int iXCoord = 0; iXCoord < pctGradient.Width; iXCoord++) 
				{	
					int iRed = scrRedTop.Value+iXCoord*(scrRedBottom.Value - scrRedTop.Value)/pctGradient.Width;
					int iGreen = scrGreenTop.Value+iXCoord*(scrGreenBottom.Value - scrGreenTop.Value)/pctGradient.Width;
					int iBlue = scrBlueTop.Value+iXCoord*(scrBlueBottom.Value - scrBlueTop.Value)/pctGradient.Width;
					int iAlpha = scrAlphaTop.Value+iXCoord*(scrAlphaBottom.Value - scrAlphaTop.Value)/pctGradient.Width;

					Pen oPen = new System.Drawing.Pen(Color.FromArgb(iAlpha,iRed,iGreen,iBlue));
					oFile.DrawLine(oPen, iXCoord, 0, iXCoord, pctGradient.Height);
				}
			}

			b.Save("gradient.png", System.Drawing.Imaging.ImageFormat.Png);
					
		}

		private void scrWidth_Scroll(object sender, System.Windows.Forms.ScrollEventArgs e)
		{
			txtWidth.Text = scrWidth.Value.ToString();
			pctGradient.Width = scrWidth.Value;
			RenderGradient();
		}

		private void scrHeight_Scroll(object sender, System.Windows.Forms.ScrollEventArgs e)
		{
			txtHeight.Text = scrHeight.Value.ToString();
			pctGradient.Height = scrHeight.Value;
			RenderGradient();
		}

		private void scrRedTop_Scroll(object sender, System.Windows.Forms.ScrollEventArgs e)
		{
			txtRedTop.Text = scrRedTop.Value.ToString();
			RenderGradient();
		}

		private void scrRedBottom_Scroll(object sender, System.Windows.Forms.ScrollEventArgs e)
		{
			txtRedBottom.Text = scrRedBottom.Value.ToString();
			RenderGradient();
		}

		private void scrGreenTop_Scroll(object sender, System.Windows.Forms.ScrollEventArgs e)
		{
			txtGreenTop.Text = scrGreenTop.Value.ToString();
			RenderGradient();
		}

		private void scrGreenBottom_Scroll(object sender, System.Windows.Forms.ScrollEventArgs e)
		{
			txtGreenBottom.Text = scrGreenBottom.Value.ToString();
			RenderGradient();
		}

		private void scrBlueTop_Scroll(object sender, System.Windows.Forms.ScrollEventArgs e)
		{
			txtBlueTop.Text = scrBlueTop.Value.ToString();
			RenderGradient();
		}

		private void hScrollBar8_Scroll(object sender, System.Windows.Forms.ScrollEventArgs e)
		{
			txtBlueBottom.Text = scrBlueBottom.Value.ToString();
			RenderGradient();
		}

		private void Form1_Load(object sender, System.EventArgs e)
		{
			RenderGradient();
		}


		private void txtWidth_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar==13) 
			{
				scrWidth.Value=System.Int32.Parse(txtWidth.Text);
				pctGradient.Width=scrWidth.Value;
				RenderGradient();
			}
	}

		private void txtWidth_Leave(object sender, System.EventArgs e)
		{
			scrWidth.Value=System.Int32.Parse(txtWidth.Text);	
			pctGradient.Width=scrWidth.Value;
			RenderGradient();
		}

		private void txtHeight_Leave(object sender, System.EventArgs e)
		{
			scrHeight.Value=System.Int32.Parse(txtHeight.Text);	
			pctGradient.Height=scrHeight.Value;
			RenderGradient();		
		}

		private void txtHeight_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar==13) 
			{
				scrHeight.Value=System.Int32.Parse(txtHeight.Text);
				pctGradient.Height=scrHeight.Value;
				RenderGradient();
			}		
		}

		private void txtRedTop_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar==13) 
			{
				scrRedTop.Value=System.Int32.Parse(txtRedTop.Text);
				RenderGradient();
			}		
		}

		private void txtRedTop_Leave(object sender, System.EventArgs e)
		{
			scrRedTop.Value=System.Int32.Parse(txtRedTop.Text);
			RenderGradient();				
		}

		private void txtRedBottom_Leave(object sender, System.EventArgs e)
		{
			scrRedBottom.Value=System.Int32.Parse(txtRedBottom.Text);
			RenderGradient();					
		}

		private void txtRedBottom_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar==13) 
			{
				scrRedBottom.Value=System.Int32.Parse(txtRedBottom.Text);
				RenderGradient();
			}		
		}

		private void txtGreenTop_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar==13) 
			{
				scrGreenTop.Value=System.Int32.Parse(txtGreenTop.Text);
				RenderGradient();
			}		
		}

		private void txtGreenTop_Leave(object sender, System.EventArgs e)
		{
			scrGreenTop.Value=System.Int32.Parse(txtGreenTop.Text);
			RenderGradient();				
		}

		private void txtGreenBottom_Leave(object sender, System.EventArgs e)
		{
			scrGreenBottom.Value=System.Int32.Parse(txtGreenBottom.Text);
			RenderGradient();				
		}

		private void txtGreenBottom_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar==13) 
			{
				scrGreenBottom.Value=System.Int32.Parse(txtGreenBottom.Text);
				RenderGradient();
			}		
		}

		private void txtBlueTop_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar==13) 
			{
				scrBlueTop.Value=System.Int32.Parse(txtBlueTop.Text);
				RenderGradient();
			}		
		}

		private void txtBlueTop_Leave(object sender, System.EventArgs e)
		{
			scrBlueTop.Value=System.Int32.Parse(txtBlueTop.Text);
			RenderGradient();						
		}

		private void txtBlueBottom_Leave(object sender, System.EventArgs e)
		{
			scrBlueBottom.Value=System.Int32.Parse(txtBlueBottom.Text);
			RenderGradient();	
		}

		private void txtBlueBottom_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar==13) 
			{
				scrBlueBottom.Value=System.Int32.Parse(txtBlueBottom.Text);
				RenderGradient();
			}		
		}

		private void scrAlphaTop_Scroll(object sender, System.Windows.Forms.ScrollEventArgs e)
		{
			txtAlphaTop.Text = scrAlphaTop.Value.ToString();
			RenderGradient();		
		}

		private void scrAlphaBottom_Scroll(object sender, System.Windows.Forms.ScrollEventArgs e)
		{
			txtAlphaBottom.Text = scrAlphaBottom.Value.ToString();
			RenderGradient();	
		}

		private void txtAlphaTop_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar==13) 
			{
				scrBlueBottom.Value=System.Int32.Parse(txtAlphaTop.Text);
				RenderGradient();
			}				
		}

		private void txtAlphaBottom_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar==13) 
			{
				scrBlueBottom.Value=System.Int32.Parse(txtAlphaBottom.Text);
				RenderGradient();
			}				
		}

		private void chkHorizontal_Click(object sender, System.EventArgs e)
		{
			RenderGradient();
		}

	}
}
