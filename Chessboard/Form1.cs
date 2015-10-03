using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

namespace RoundCorners
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
		private System.Windows.Forms.TextBox txtHeight;
		private System.Windows.Forms.TextBox txtWidth;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.TextBox txtRadius;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.HScrollBar scrRadius;
		private System.Windows.Forms.HScrollBar scrRed1;
		private System.Windows.Forms.HScrollBar scrGreen1;
		private System.Windows.Forms.HScrollBar scrBlue1;
		private System.Windows.Forms.HScrollBar scrAlpha1;
		private System.Windows.Forms.TextBox txtRed1;
		private System.Windows.Forms.TextBox txtGreen1;
		private System.Windows.Forms.TextBox txtBlue1;
		private System.Windows.Forms.TextBox txtAlpha1;
		private System.Windows.Forms.Button cmdSelector1;
		private System.Windows.Forms.Button cmdSelector2;
		private System.Windows.Forms.Button cmdSelector3;
		private System.Windows.Forms.Button cmdSelector4;

		private int[] miRed = new int[4];
		private int[] miGreen = new int[4];
		private int[] miBlue = new int[4];
		private int[] miAlpha = new int[4];
		private int miSelectedColour = 0;

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
			this.pctGradient = new System.Windows.Forms.PictureBox();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.scrWidth = new System.Windows.Forms.HScrollBar();
			this.scrHeight = new System.Windows.Forms.HScrollBar();
			this.scrRed1 = new System.Windows.Forms.HScrollBar();
			this.scrGreen1 = new System.Windows.Forms.HScrollBar();
			this.scrBlue1 = new System.Windows.Forms.HScrollBar();
			this.txtHeight = new System.Windows.Forms.TextBox();
			this.txtRed1 = new System.Windows.Forms.TextBox();
			this.txtGreen1 = new System.Windows.Forms.TextBox();
			this.txtBlue1 = new System.Windows.Forms.TextBox();
			this.txtWidth = new System.Windows.Forms.TextBox();
			this.txtAlpha1 = new System.Windows.Forms.TextBox();
			this.scrAlpha1 = new System.Windows.Forms.HScrollBar();
			this.label6 = new System.Windows.Forms.Label();
			this.txtRadius = new System.Windows.Forms.TextBox();
			this.scrRadius = new System.Windows.Forms.HScrollBar();
			this.label7 = new System.Windows.Forms.Label();
			this.cmdSelector1 = new System.Windows.Forms.Button();
			this.cmdSelector2 = new System.Windows.Forms.Button();
			this.cmdSelector3 = new System.Windows.Forms.Button();
			this.cmdSelector4 = new System.Windows.Forms.Button();
			this.SuspendLayout();
			// 
			// pctGradient
			// 
			this.pctGradient.BackColor = System.Drawing.SystemColors.HighlightText;
			this.pctGradient.Location = new System.Drawing.Point(8, 128);
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
			this.label3.Size = new System.Drawing.Size(40, 16);
			this.label3.TabIndex = 4;
			this.label3.Text = "Red";
			// 
			// label4
			// 
			this.label4.Location = new System.Drawing.Point(376, 32);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(40, 16);
			this.label4.TabIndex = 5;
			this.label4.Text = "Green";
			// 
			// label5
			// 
			this.label5.Location = new System.Drawing.Point(376, 56);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(40, 16);
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
			// scrRed1
			// 
			this.scrRed1.Location = new System.Drawing.Point(416, 8);
			this.scrRed1.Maximum = 264;
			this.scrRed1.Name = "scrRed1";
			this.scrRed1.Size = new System.Drawing.Size(88, 16);
			this.scrRed1.TabIndex = 9;
			this.scrRed1.Scroll += new System.Windows.Forms.ScrollEventHandler(this.scrRed1_Scroll);
			// 
			// scrGreen1
			// 
			this.scrGreen1.Location = new System.Drawing.Point(416, 32);
			this.scrGreen1.Maximum = 264;
			this.scrGreen1.Name = "scrGreen1";
			this.scrGreen1.Size = new System.Drawing.Size(88, 16);
			this.scrGreen1.TabIndex = 10;
			this.scrGreen1.Scroll += new System.Windows.Forms.ScrollEventHandler(this.scrGreen1_Scroll);
			// 
			// scrBlue1
			// 
			this.scrBlue1.Location = new System.Drawing.Point(416, 56);
			this.scrBlue1.Maximum = 264;
			this.scrBlue1.Name = "scrBlue1";
			this.scrBlue1.Size = new System.Drawing.Size(88, 16);
			this.scrBlue1.TabIndex = 11;
			this.scrBlue1.Scroll += new System.Windows.Forms.ScrollEventHandler(this.scrBlue1_Scroll);
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
			// txtRed1
			// 
			this.txtRed1.Location = new System.Drawing.Point(512, 8);
			this.txtRed1.Name = "txtRed1";
			this.txtRed1.Size = new System.Drawing.Size(32, 20);
			this.txtRed1.TabIndex = 17;
			this.txtRed1.Text = "0";
			this.txtRed1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtRed1_KeyPress);
			this.txtRed1.Leave += new System.EventHandler(this.txtRed1_Leave);
			// 
			// txtGreen1
			// 
			this.txtGreen1.Location = new System.Drawing.Point(512, 32);
			this.txtGreen1.Name = "txtGreen1";
			this.txtGreen1.Size = new System.Drawing.Size(32, 20);
			this.txtGreen1.TabIndex = 18;
			this.txtGreen1.Text = "0";
			this.txtGreen1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtGreen1_KeyPress);
			this.txtGreen1.Leave += new System.EventHandler(this.txtGreen1_Leave);
			// 
			// txtBlue1
			// 
			this.txtBlue1.Location = new System.Drawing.Point(512, 56);
			this.txtBlue1.Name = "txtBlue1";
			this.txtBlue1.Size = new System.Drawing.Size(32, 20);
			this.txtBlue1.TabIndex = 19;
			this.txtBlue1.Text = "0";
			this.txtBlue1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtBlue1_KeyPress);
			this.txtBlue1.Leave += new System.EventHandler(this.txtBlue1_Leave);
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
			// txtAlpha1
			// 
			this.txtAlpha1.Location = new System.Drawing.Point(512, 80);
			this.txtAlpha1.Name = "txtAlpha1";
			this.txtAlpha1.Size = new System.Drawing.Size(32, 20);
			this.txtAlpha1.TabIndex = 27;
			this.txtAlpha1.Text = "255";
			this.txtAlpha1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtAlpha1_KeyPress);
			this.txtAlpha1.Leave += new System.EventHandler(this.txtAlpha1_Leave);
			// 
			// scrAlpha1
			// 
			this.scrAlpha1.Location = new System.Drawing.Point(416, 80);
			this.scrAlpha1.Maximum = 264;
			this.scrAlpha1.Name = "scrAlpha1";
			this.scrAlpha1.Size = new System.Drawing.Size(88, 16);
			this.scrAlpha1.TabIndex = 25;
			this.scrAlpha1.Value = 255;
			this.scrAlpha1.Scroll += new System.Windows.Forms.ScrollEventHandler(this.scrAlpha1_Scroll);
			// 
			// label6
			// 
			this.label6.Location = new System.Drawing.Point(376, 80);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(40, 16);
			this.label6.TabIndex = 24;
			this.label6.Text = "Alpha";
			// 
			// txtRadius
			// 
			this.txtRadius.Location = new System.Drawing.Point(328, 64);
			this.txtRadius.Name = "txtRadius";
			this.txtRadius.Size = new System.Drawing.Size(32, 20);
			this.txtRadius.TabIndex = 32;
			this.txtRadius.Text = "32";
			this.txtRadius.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtRadius_KeyPress);
			// 
			// scrRadius
			// 
			this.scrRadius.Location = new System.Drawing.Point(88, 56);
			this.scrRadius.Maximum = 264;
			this.scrRadius.Minimum = 1;
			this.scrRadius.Name = "scrRadius";
			this.scrRadius.Size = new System.Drawing.Size(88, 16);
			this.scrRadius.TabIndex = 31;
			this.scrRadius.Value = 32;
			this.scrRadius.Scroll += new System.Windows.Forms.ScrollEventHandler(this.scrRadius_Scroll);
			// 
			// label7
			// 
			this.label7.Location = new System.Drawing.Point(8, 56);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(40, 16);
			this.label7.TabIndex = 30;
			this.label7.Text = "Radius";
			// 
			// cmdSelector1
			// 
			this.cmdSelector1.Location = new System.Drawing.Point(552, 8);
			this.cmdSelector1.Name = "cmdSelector1";
			this.cmdSelector1.Size = new System.Drawing.Size(24, 24);
			this.cmdSelector1.TabIndex = 45;
			this.cmdSelector1.Text = "1";
			this.cmdSelector1.Click += new System.EventHandler(this.cmdSelector1_Click);
			// 
			// cmdSelector2
			// 
			this.cmdSelector2.Location = new System.Drawing.Point(584, 8);
			this.cmdSelector2.Name = "cmdSelector2";
			this.cmdSelector2.Size = new System.Drawing.Size(24, 24);
			this.cmdSelector2.TabIndex = 46;
			this.cmdSelector2.Text = "2";
			this.cmdSelector2.Click += new System.EventHandler(this.cmdSelector2_Click);
			// 
			// cmdSelector3
			// 
			this.cmdSelector3.Location = new System.Drawing.Point(616, 8);
			this.cmdSelector3.Name = "cmdSelector3";
			this.cmdSelector3.Size = new System.Drawing.Size(24, 24);
			this.cmdSelector3.TabIndex = 47;
			this.cmdSelector3.Text = "3";
			this.cmdSelector3.Click += new System.EventHandler(this.cmdSelector3_Click);
			// 
			// cmdSelector4
			// 
			this.cmdSelector4.Location = new System.Drawing.Point(648, 8);
			this.cmdSelector4.Name = "cmdSelector4";
			this.cmdSelector4.Size = new System.Drawing.Size(24, 24);
			this.cmdSelector4.TabIndex = 48;
			this.cmdSelector4.Text = "4";
			this.cmdSelector4.Click += new System.EventHandler(this.cmdSelector4_Click);
			// 
			// Form1
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(896, 614);
			this.Controls.AddRange(new System.Windows.Forms.Control[] {
																		  this.cmdSelector4,
																		  this.cmdSelector3,
																		  this.cmdSelector2,
																		  this.cmdSelector1,
																		  this.txtRadius,
																		  this.scrRadius,
																		  this.label7,
																		  this.txtAlpha1,
																		  this.scrAlpha1,
																		  this.label6,
																		  this.txtBlue1,
																		  this.txtGreen1,
																		  this.txtRed1,
																		  this.txtHeight,
																		  this.txtWidth,
																		  this.scrBlue1,
																		  this.scrGreen1,
																		  this.scrRed1,
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
			RenderCorners();
		}

		private void RenderCorners() 
		{
			Bitmap b = new Bitmap(pctGradient.Width,pctGradient.Height);

			Graphics oCanvas = pctGradient.CreateGraphics();
			oCanvas.SmoothingMode = SmoothingMode.HighQuality;

			//pctGradient.Refresh();
			
			int iDiameter = scrRadius.Value*2;
			int iRadius = scrRadius.Value;

			SolidBrush[] oBrush = new SolidBrush[4];
			oBrush[0] = new System.Drawing.SolidBrush(Color.FromArgb(miAlpha[0],miRed[0],miGreen[0],miBlue[0]));
			oBrush[1] = new System.Drawing.SolidBrush(Color.FromArgb(miAlpha[1],miRed[1],miGreen[1],miBlue[1]));
			oBrush[2] = new System.Drawing.SolidBrush(Color.FromArgb(miAlpha[2],miRed[2],miGreen[2],miBlue[3]));
			oBrush[3] = new System.Drawing.SolidBrush(Color.FromArgb(miAlpha[3],miRed[3],miGreen[3],miBlue[3]));

			int iColourIndexY = 0;
			int iColourIndexX = 0;
			int iBrush;

			for (int y=0; y<pctGradient.Height; y+=iRadius) 
			{
				iColourIndexX = 0;
				for (int x=0; x<pctGradient.Width; x+=iRadius) 
				{
					iBrush = 	(iColourIndexX+iColourIndexY)%2;
					oCanvas.FillRectangle(oBrush[iBrush],x,y,iRadius+1,iRadius+1);
					iColourIndexX++;
				}
				iColourIndexY++;
			}

			oCanvas.FillRectangle(oBrush[3],75,75,400,400);

			oCanvas.FillRectangle(oBrush[2],5*iRadius,6*iRadius,iRadius+1,iRadius+1);
		}


		private void cmdGenerateFile_Click(object sender, System.EventArgs e)
		{
		/*
			int iRed1 = scrRed1.Value;
			int iGreen1 = scrGreen1.Value;
			int iBlue1 = scrBlue1.Value;
			int iAlpha1 = scrAlpha1.Value;

			int iDiameter = scrRadius.Value*2;
			int iRadius = scrRadius.Value;
			Bitmap b;
			Graphics oFile;
			Pen oPen = new System.Drawing.Pen(Color.FromArgb(iAlpha1,iRed1,iGreen1,iBlue1));
			SolidBrush oBrush = new System.Drawing.SolidBrush(Color.FromArgb(iAlpha1,iRed1,iGreen1,iBlue1));
			SolidBrush oBrush2 = new System.Drawing.SolidBrush(Color.FromArgb(iAlpha2,iRed2,iGreen2,iBlue2));

			b = new Bitmap(iRadius,iRadius);
			oFile = Graphics.FromImage(b);
			oFile.SmoothingMode = SmoothingMode.AntiAlias;
			oFile.FillRectangle(oBrush2, -0.5f, -0.5f, iRadius, iRadius);
			oFile.FillPie(oBrush, new Rectangle(0, 0, iDiameter, iDiameter), 180f, 90f);
			b.Save("corner-nw.png", System.Drawing.Imaging.ImageFormat.Png);
			oFile.Dispose();
			b.Dispose();

			b = new Bitmap(iRadius,iRadius);
			oFile = Graphics.FromImage(b);
			oFile.SmoothingMode = SmoothingMode.AntiAlias;
			oFile.FillRectangle(oBrush2, -0.5f, -0.5f, iRadius, iRadius);
			oFile.FillPie(oBrush, new Rectangle(-iRadius-1, 0, iDiameter, iDiameter), 270f, 90f);
			b.Save("corner-ne.png", System.Drawing.Imaging.ImageFormat.Png);
			oFile.Dispose();
			b.Dispose();

			b = new Bitmap(iRadius,iRadius);
			oFile = Graphics.FromImage(b);
			oFile.SmoothingMode = SmoothingMode.AntiAlias;
			oFile.FillRectangle(oBrush2, -0.5f, -0.5f, iRadius, iRadius);
			oFile.FillPie(oBrush, new Rectangle(0, -iRadius-1, iDiameter,iDiameter), 90f, 90f);
			b.Save("corner-sw.png", System.Drawing.Imaging.ImageFormat.Png);
			oFile.Dispose();
			b.Dispose();

			b = new Bitmap(iRadius,iRadius);
			oFile = Graphics.FromImage(b);
			oFile.SmoothingMode = SmoothingMode.AntiAlias;
			oFile.FillRectangle(oBrush2, -0.5f, -0.5f, iRadius, iRadius);
			oFile.FillPie(oBrush, new Rectangle(-iRadius-1, -iRadius-1, iDiameter,iDiameter), 0f, 90f);
			b.Save("corner-se.png", System.Drawing.Imaging.ImageFormat.Png);
			oFile.Dispose();
			b.Dispose();
*/			
		}

		private void scrWidth_Scroll(object sender, System.Windows.Forms.ScrollEventArgs e)
		{
			txtWidth.Text = scrWidth.Value.ToString();
			pctGradient.Width = scrWidth.Value;
			RenderCorners();
		}

		private void scrHeight_Scroll(object sender, System.Windows.Forms.ScrollEventArgs e)
		{
			txtHeight.Text = scrHeight.Value.ToString();
			pctGradient.Height = scrHeight.Value;
			RenderCorners();
		}




		private void Form1_Load(object sender, System.EventArgs e)
		{
			RenderCorners();
		}


		private void txtWidth_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar==13) 
			{
				scrWidth.Value=System.Int32.Parse(txtWidth.Text);
				pctGradient.Width=scrWidth.Value;
				RenderCorners();
			}
	}

		private void txtWidth_Leave(object sender, System.EventArgs e)
		{
			scrWidth.Value=System.Int32.Parse(txtWidth.Text);	
			pctGradient.Width=scrWidth.Value;
			RenderCorners();
		}

		private void txtHeight_Leave(object sender, System.EventArgs e)
		{
			scrHeight.Value=System.Int32.Parse(txtHeight.Text);	
			pctGradient.Height=scrHeight.Value;
			RenderCorners();		
		}

		private void txtHeight_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar==13) 
			{
				scrHeight.Value=System.Int32.Parse(txtHeight.Text);
				pctGradient.Height=scrHeight.Value;
				RenderCorners();
			}		
		}







		private void chkHorizontal_Click(object sender, System.EventArgs e)
		{
			RenderCorners();
		}

		private void scrRadius_Scroll(object sender, System.Windows.Forms.ScrollEventArgs e)
		{
			txtRadius.Text = scrRadius.Value.ToString();
			RenderCorners();
		}

		private void txtRadius_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar==13) 
			{
				scrRadius.Value=System.Int32.Parse(txtRadius.Text);
				RenderCorners();
			}				
		}



		private void scrRed1_Scroll(object sender, System.Windows.Forms.ScrollEventArgs e)
		{
			txtRed1.Text = scrRed1.Value.ToString();
			miRed[miSelectedColour] = scrRed1.Value;
			RenderCorners();
		}
		private void scrGreen1_Scroll(object sender, System.Windows.Forms.ScrollEventArgs e)
		{
			txtGreen1.Text = scrGreen1.Value.ToString();
			miGreen[miSelectedColour] = scrGreen1.Value;
			RenderCorners();
		}
		private void scrBlue1_Scroll(object sender, System.Windows.Forms.ScrollEventArgs e)
		{
			txtBlue1.Text = scrBlue1.Value.ToString();
			miBlue[miSelectedColour] = scrBlue1.Value;
			RenderCorners();
		}
		private void scrAlpha1_Scroll(object sender, System.Windows.Forms.ScrollEventArgs e)
		{
			txtAlpha1.Text = scrAlpha1.Value.ToString();
			miAlpha[miSelectedColour] = scrAlpha1.Value;
			RenderCorners();		
		}

		private void txtRed1_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar==13) 
			{
				scrRed1.Value=System.Int32.Parse(txtRed1.Text);
				miRed[miSelectedColour] = scrRed1.Value;
				RenderCorners();
			}		
		}
		private void txtGreen1_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar==13) 
			{
				scrGreen1.Value=System.Int32.Parse(txtGreen1.Text);
				miGreen[miSelectedColour] = scrGreen1.Value;
				RenderCorners();
			}		
		}
		private void txtBlue1_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar==13) 
			{
				scrBlue1.Value=System.Int32.Parse(txtBlue1.Text);
				miBlue[miSelectedColour] = scrBlue1.Value;
				RenderCorners();
			}		
		}
		private void txtAlpha1_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar==13) 
			{
				scrAlpha1.Value=System.Int32.Parse(txtAlpha1.Text);
				miAlpha[miSelectedColour] = scrAlpha1.Value;
				RenderCorners();
			}				
		}

		private void txtRed1_Leave(object sender, System.EventArgs e)
		{
			scrRed1.Value=System.Int32.Parse(txtRed1.Text);
			miRed[miSelectedColour] = scrRed1.Value;
			RenderCorners();				
		}
		private void txtGreen1_Leave(object sender, System.EventArgs e)
		{
			scrGreen1.Value=System.Int32.Parse(txtGreen1.Text);
			miGreen[miSelectedColour] = scrGreen1.Value;
			RenderCorners();				
		}
		private void txtBlue1_Leave(object sender, System.EventArgs e)
		{
			scrBlue1.Value=System.Int32.Parse(txtBlue1.Text);
			miBlue[miSelectedColour] = scrBlue1.Value;
			RenderCorners();						
		}
		private void txtAlpha1_Leave(object sender, System.EventArgs e)
		{
			scrAlpha1.Value=System.Int32.Parse(txtAlpha1.Text);
			miAlpha[miSelectedColour] = scrAlpha1.Value;
			RenderCorners();				
		}

		private void cmdSelector1_Click(object sender, System.EventArgs e)
		{
			miSelectedColour = 0;
			scrRed1.Value = miRed[0];
			scrGreen1.Value = miGreen[0];
			scrBlue1.Value = miBlue[0];
			txtRed1.Text = miRed[0].ToString();
			txtGreen1.Text = miGreen[0].ToString();
			txtBlue1.Text = miBlue[0].ToString();
		}

		private void cmdSelector2_Click(object sender, System.EventArgs e)
		{
			miSelectedColour = 1;
			scrRed1.Value = miRed[1];
			scrGreen1.Value = miGreen[1];
			scrBlue1.Value = miBlue[1];
			txtRed1.Text = miRed[1].ToString();
			txtGreen1.Text = miGreen[1].ToString();
			txtBlue1.Text = miBlue[1].ToString();
		}

		private void cmdSelector3_Click(object sender, System.EventArgs e)
		{
			miSelectedColour = 2;
			scrRed1.Value = miRed[2];
			scrGreen1.Value = miGreen[2];
			scrBlue1.Value = miBlue[2];
			txtRed1.Text = miRed[2].ToString();
			txtGreen1.Text = miGreen[2].ToString();
			txtBlue1.Text = miBlue[2].ToString();
		}

		private void cmdSelector4_Click(object sender, System.EventArgs e)
		{
			miSelectedColour = 3;
			scrRed1.Value = miRed[3];
			scrGreen1.Value = miGreen[3];
			scrBlue1.Value = miBlue[3];
			txtRed1.Text = miRed[3].ToString();
			txtGreen1.Text = miGreen[3].ToString();
			txtBlue1.Text = miBlue[3].ToString();
		}

	}
}
