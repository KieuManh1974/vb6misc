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
		private System.Windows.Forms.Button cmdGenerateFile;
		private System.Windows.Forms.TextBox txtWidth;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.TextBox txtRadius;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.HScrollBar scrRadius;
		private System.Windows.Forms.HScrollBar scrRed1;
		private System.Windows.Forms.HScrollBar scrGreen1;
		private System.Windows.Forms.HScrollBar scrBlue1;
		private System.Windows.Forms.HScrollBar scrAlpha1;
		private System.Windows.Forms.HScrollBar scrAlpha2;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.HScrollBar scrBlue2;
		private System.Windows.Forms.HScrollBar scrGreen2;
		private System.Windows.Forms.HScrollBar scrRed2;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Label label10;
		private System.Windows.Forms.Label label11;
		private System.Windows.Forms.TextBox txtRed1;
		private System.Windows.Forms.TextBox txtGreen1;
		private System.Windows.Forms.TextBox txtBlue1;
		private System.Windows.Forms.TextBox txtAlpha1;
		private System.Windows.Forms.TextBox txtAlpha2;
		private System.Windows.Forms.TextBox txtBlue2;
		private System.Windows.Forms.TextBox txtGreen2;
		private System.Windows.Forms.TextBox txtRed2;
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
			this.scrRed1 = new System.Windows.Forms.HScrollBar();
			this.scrGreen1 = new System.Windows.Forms.HScrollBar();
			this.scrBlue1 = new System.Windows.Forms.HScrollBar();
			this.txtHeight = new System.Windows.Forms.TextBox();
			this.txtRed1 = new System.Windows.Forms.TextBox();
			this.txtGreen1 = new System.Windows.Forms.TextBox();
			this.txtBlue1 = new System.Windows.Forms.TextBox();
			this.cmdGenerateFile = new System.Windows.Forms.Button();
			this.txtWidth = new System.Windows.Forms.TextBox();
			this.txtAlpha1 = new System.Windows.Forms.TextBox();
			this.scrAlpha1 = new System.Windows.Forms.HScrollBar();
			this.label6 = new System.Windows.Forms.Label();
			this.txtRadius = new System.Windows.Forms.TextBox();
			this.scrRadius = new System.Windows.Forms.HScrollBar();
			this.label7 = new System.Windows.Forms.Label();
			this.txtAlpha2 = new System.Windows.Forms.TextBox();
			this.scrAlpha2 = new System.Windows.Forms.HScrollBar();
			this.label8 = new System.Windows.Forms.Label();
			this.txtBlue2 = new System.Windows.Forms.TextBox();
			this.txtGreen2 = new System.Windows.Forms.TextBox();
			this.txtRed2 = new System.Windows.Forms.TextBox();
			this.scrBlue2 = new System.Windows.Forms.HScrollBar();
			this.scrGreen2 = new System.Windows.Forms.HScrollBar();
			this.scrRed2 = new System.Windows.Forms.HScrollBar();
			this.label9 = new System.Windows.Forms.Label();
			this.label10 = new System.Windows.Forms.Label();
			this.label11 = new System.Windows.Forms.Label();
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
			this.scrRadius.Location = new System.Drawing.Point(232, 64);
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
			this.label7.Location = new System.Drawing.Point(192, 64);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(40, 16);
			this.label7.TabIndex = 30;
			this.label7.Text = "Radius";
			// 
			// txtAlpha2
			// 
			this.txtAlpha2.Location = new System.Drawing.Point(696, 80);
			this.txtAlpha2.Name = "txtAlpha2";
			this.txtAlpha2.Size = new System.Drawing.Size(32, 20);
			this.txtAlpha2.TabIndex = 44;
			this.txtAlpha2.Text = "255";
			this.txtAlpha2.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtAlpha2_KeyPress);
			this.txtAlpha2.Leave += new System.EventHandler(this.txtRed2_Leave);
			// 
			// scrAlpha2
			// 
			this.scrAlpha2.Location = new System.Drawing.Point(600, 80);
			this.scrAlpha2.Maximum = 264;
			this.scrAlpha2.Name = "scrAlpha2";
			this.scrAlpha2.Size = new System.Drawing.Size(88, 16);
			this.scrAlpha2.TabIndex = 43;
			this.scrAlpha2.Value = 255;
			this.scrAlpha2.Scroll += new System.Windows.Forms.ScrollEventHandler(this.scrAlpha2_Scroll);
			// 
			// label8
			// 
			this.label8.Location = new System.Drawing.Point(560, 80);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(40, 16);
			this.label8.TabIndex = 42;
			this.label8.Text = "Alpha";
			// 
			// txtBlue2
			// 
			this.txtBlue2.Location = new System.Drawing.Point(696, 56);
			this.txtBlue2.Name = "txtBlue2";
			this.txtBlue2.Size = new System.Drawing.Size(32, 20);
			this.txtBlue2.TabIndex = 41;
			this.txtBlue2.Text = "255";
			this.txtBlue2.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtBlue2_KeyPress);
			this.txtBlue2.Leave += new System.EventHandler(this.txtRed2_Leave);
			// 
			// txtGreen2
			// 
			this.txtGreen2.Location = new System.Drawing.Point(696, 32);
			this.txtGreen2.Name = "txtGreen2";
			this.txtGreen2.Size = new System.Drawing.Size(32, 20);
			this.txtGreen2.TabIndex = 40;
			this.txtGreen2.Text = "255";
			this.txtGreen2.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtGreen2_KeyPress);
			this.txtGreen2.Leave += new System.EventHandler(this.txtRed2_Leave);
			// 
			// txtRed2
			// 
			this.txtRed2.Location = new System.Drawing.Point(696, 8);
			this.txtRed2.Name = "txtRed2";
			this.txtRed2.Size = new System.Drawing.Size(32, 20);
			this.txtRed2.TabIndex = 39;
			this.txtRed2.Text = "255";
			this.txtRed2.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtRed2_KeyPress);
			this.txtRed2.Leave += new System.EventHandler(this.txtRed2_Leave);
			// 
			// scrBlue2
			// 
			this.scrBlue2.Location = new System.Drawing.Point(600, 56);
			this.scrBlue2.Maximum = 264;
			this.scrBlue2.Name = "scrBlue2";
			this.scrBlue2.Size = new System.Drawing.Size(88, 16);
			this.scrBlue2.TabIndex = 38;
			this.scrBlue2.Value = 255;
			this.scrBlue2.Scroll += new System.Windows.Forms.ScrollEventHandler(this.scrBlue2_Scroll);
			// 
			// scrGreen2
			// 
			this.scrGreen2.Location = new System.Drawing.Point(600, 32);
			this.scrGreen2.Maximum = 264;
			this.scrGreen2.Name = "scrGreen2";
			this.scrGreen2.Size = new System.Drawing.Size(88, 16);
			this.scrGreen2.TabIndex = 37;
			this.scrGreen2.Value = 255;
			this.scrGreen2.Scroll += new System.Windows.Forms.ScrollEventHandler(this.scrGreen2_Scroll);
			// 
			// scrRed2
			// 
			this.scrRed2.Location = new System.Drawing.Point(600, 8);
			this.scrRed2.Maximum = 264;
			this.scrRed2.Name = "scrRed2";
			this.scrRed2.Size = new System.Drawing.Size(88, 16);
			this.scrRed2.TabIndex = 36;
			this.scrRed2.Value = 255;
			this.scrRed2.Scroll += new System.Windows.Forms.ScrollEventHandler(this.scrRed2_Scroll);
			// 
			// label9
			// 
			this.label9.Location = new System.Drawing.Point(560, 56);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(40, 16);
			this.label9.TabIndex = 35;
			this.label9.Text = "Blue";
			// 
			// label10
			// 
			this.label10.Location = new System.Drawing.Point(560, 32);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(40, 16);
			this.label10.TabIndex = 34;
			this.label10.Text = "Green";
			// 
			// label11
			// 
			this.label11.Location = new System.Drawing.Point(560, 8);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(40, 16);
			this.label11.TabIndex = 33;
			this.label11.Text = "Red";
			// 
			// Form1
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(896, 614);
			this.Controls.AddRange(new System.Windows.Forms.Control[] {
																		  this.txtAlpha2,
																		  this.scrAlpha2,
																		  this.label8,
																		  this.txtBlue2,
																		  this.txtGreen2,
																		  this.txtRed2,
																		  this.scrBlue2,
																		  this.scrGreen2,
																		  this.scrRed2,
																		  this.label9,
																		  this.label10,
																		  this.label11,
																		  this.txtRadius,
																		  this.scrRadius,
																		  this.label7,
																		  this.txtAlpha1,
																		  this.scrAlpha1,
																		  this.label6,
																		  this.cmdGenerateFile,
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

			pctGradient.Refresh();
			
			int iRed1 = scrRed1.Value;
			int iGreen1 = scrGreen1.Value;
			int iBlue1 = scrBlue1.Value;
			int iAlpha1 = scrAlpha1.Value;

			int iRed2 = scrRed2.Value;
			int iGreen2 = scrGreen2.Value;
			int iBlue2 = scrBlue2.Value;
			int iAlpha2 = scrAlpha2.Value;

			int iDiameter = scrRadius.Value*2;
			int iRadius = scrRadius.Value;

			Pen oPen = new System.Drawing.Pen(Color.FromArgb(iAlpha1,iRed1,iGreen1,iBlue1));
			SolidBrush oBrush2 = new System.Drawing.SolidBrush(Color.FromArgb(iAlpha2,iRed2,iGreen2,iBlue2));
			SolidBrush oBrush = new System.Drawing.SolidBrush(Color.FromArgb(iAlpha1,iRed1,iGreen1,iBlue1));

			oCanvas.FillRectangle(oBrush2, 0, 0, pctGradient.Width, pctGradient.Height);

			oCanvas.FillPie(oBrush, new Rectangle(0, 0, iDiameter, iDiameter), 180f, 90f);
			oCanvas.FillPie(oBrush, new Rectangle(pctGradient.Width-iDiameter,0, iDiameter,iDiameter), 270f, 90f);
			oCanvas.FillPie(oBrush, new Rectangle(0, pctGradient.Height-iDiameter, iDiameter,iDiameter), 90f, 90f);
			oCanvas.FillPie(oBrush, new Rectangle(pctGradient.Width-iDiameter, pctGradient.Height-iDiameter, iDiameter,iDiameter), 0f, 90f);

			oCanvas.FillRectangle(oBrush, iRadius, 0, pctGradient.Width-iDiameter, iRadius);
			oCanvas.FillRectangle(oBrush, 0, iRadius, iRadius, pctGradient.Height-iDiameter);
			oCanvas.FillRectangle(oBrush, pctGradient.Width-iRadius, iRadius, iRadius, pctGradient.Height-iDiameter);
			oCanvas.FillRectangle(oBrush, iRadius, pctGradient.Height-iRadius, pctGradient.Width-iDiameter, iRadius);

			oCanvas.FillRectangle(oBrush, iRadius, iRadius, pctGradient.Width-iDiameter, pctGradient.Height-iDiameter);
		}

		private void cmdGenerateFile_Click(object sender, System.EventArgs e)
		{

			int iRed1 = scrRed1.Value;
			int iGreen1 = scrGreen1.Value;
			int iBlue1 = scrBlue1.Value;
			int iAlpha1 = scrAlpha1.Value;

			int iRed2 = scrRed2.Value;
			int iGreen2 = scrGreen2.Value;
			int iBlue2 = scrBlue2.Value;
			int iAlpha2 = scrAlpha2.Value;

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
			RenderCorners();
		}
		private void scrGreen1_Scroll(object sender, System.Windows.Forms.ScrollEventArgs e)
		{
			txtGreen1.Text = scrGreen1.Value.ToString();
			RenderCorners();
		}
		private void scrBlue1_Scroll(object sender, System.Windows.Forms.ScrollEventArgs e)
		{
			txtBlue1.Text = scrBlue1.Value.ToString();
			RenderCorners();
		}
		private void scrAlpha1_Scroll(object sender, System.Windows.Forms.ScrollEventArgs e)
		{
			txtAlpha1.Text = scrAlpha1.Value.ToString();
			RenderCorners();		
		}

		private void scrRed2_Scroll(object sender, System.Windows.Forms.ScrollEventArgs e)
		{
			txtRed2.Text = scrRed2.Value.ToString();
			RenderCorners();
		}
		private void scrGreen2_Scroll(object sender, System.Windows.Forms.ScrollEventArgs e)
		{
			txtGreen2.Text = scrGreen2.Value.ToString();
			RenderCorners();		
		}
		private void scrBlue2_Scroll(object sender, System.Windows.Forms.ScrollEventArgs e)
		{
			txtBlue2.Text = scrBlue2.Value.ToString();
			RenderCorners();		
		}
		private void scrAlpha2_Scroll(object sender, System.Windows.Forms.ScrollEventArgs e)
		{
			txtAlpha2.Text = scrAlpha2.Value.ToString();
			RenderCorners();		
		}

		private void txtRed1_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar==13) 
			{
				scrRed1.Value=System.Int32.Parse(txtRed1.Text);
				RenderCorners();
			}		
		}
		private void txtGreen1_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar==13) 
			{
				scrGreen1.Value=System.Int32.Parse(txtGreen1.Text);
				RenderCorners();
			}		
		}
		private void txtBlue1_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar==13) 
			{
				scrBlue1.Value=System.Int32.Parse(txtBlue1.Text);
				RenderCorners();
			}		
		}
		private void txtAlpha1_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar==13) 
			{
				scrAlpha1.Value=System.Int32.Parse(txtAlpha1.Text);
				RenderCorners();
			}				
		}

		private void txtRed2_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar==13) 
			{
				scrRed2.Value=System.Int32.Parse(txtRed2.Text);
				RenderCorners();
			}		
		}
		private void txtGreen2_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar==13) 
			{
				scrGreen2.Value=System.Int32.Parse(txtGreen2.Text);
				RenderCorners();
			}		
		}
		private void txtBlue2_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar==13) 
			{
				scrBlue2.Value=System.Int32.Parse(txtBlue2.Text);
				RenderCorners();
			}		
		}
		private void txtAlpha2_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar==13) 
			{
				scrAlpha2.Value=System.Int32.Parse(txtAlpha2.Text);
				RenderCorners();
			}				
		}

		private void txtRed1_Leave(object sender, System.EventArgs e)
		{
			scrRed1.Value=System.Int32.Parse(txtRed1.Text);
			RenderCorners();				
		}
		private void txtGreen1_Leave(object sender, System.EventArgs e)
		{
			scrGreen1.Value=System.Int32.Parse(txtGreen1.Text);
			RenderCorners();				
		}
		private void txtBlue1_Leave(object sender, System.EventArgs e)
		{
			scrBlue1.Value=System.Int32.Parse(txtBlue1.Text);
			RenderCorners();						
		}
		private void txtAlpha1_Leave(object sender, System.EventArgs e)
		{
			scrAlpha1.Value=System.Int32.Parse(txtAlpha1.Text);
			RenderCorners();				
		}

		private void txtRed2_Leave(object sender, System.EventArgs e)
		{
			scrRed2.Value=System.Int32.Parse(txtRed2.Text);
			RenderCorners();				
		}
		private void txtGreen2_Leave(object sender, System.EventArgs e)
		{
			scrGreen2.Value=System.Int32.Parse(txtGreen2.Text);
			RenderCorners();				
		}
		private void txtBlue2_Leave(object sender, System.EventArgs e)
		{
			scrBlue2.Value=System.Int32.Parse(txtBlue2.Text);
			RenderCorners();						
		}
		private void txtAlpha2_Leave(object sender, System.EventArgs e)
		{
			scrAlpha2.Value=System.Int32.Parse(txtAlpha2.Text);
			RenderCorners();				
		}

	}
}
