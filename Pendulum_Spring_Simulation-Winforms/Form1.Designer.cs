namespace WindowsFormsApplication1
{
    partial class Simulation
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
            this.components = new System.ComponentModel.Container();
            System.Windows.Forms.DataVisualization.Charting.ChartArea chartArea1 = new System.Windows.Forms.DataVisualization.Charting.ChartArea();
            System.Windows.Forms.DataVisualization.Charting.Legend legend1 = new System.Windows.Forms.DataVisualization.Charting.Legend();
            System.Windows.Forms.DataVisualization.Charting.Series series1 = new System.Windows.Forms.DataVisualization.Charting.Series();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.btnStart = new System.Windows.Forms.Button();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.lblX = new System.Windows.Forms.Label();
            this.lblY = new System.Windows.Forms.Label();
            this.chart = new System.Windows.Forms.DataVisualization.Charting.Chart();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label8 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.tbH = new System.Windows.Forms.TextBox();
            this.label29 = new System.Windows.Forms.Label();
            this.label30 = new System.Windows.Forms.Label();
            this.tbG = new System.Windows.Forms.TextBox();
            this.label20 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.tbk = new System.Windows.Forms.TextBox();
            this.label18 = new System.Windows.Forms.Label();
            this.label13 = new System.Windows.Forms.Label();
            this.label15 = new System.Windows.Forms.Label();
            this.tbm = new System.Windows.Forms.TextBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.vy0 = new System.Windows.Forms.TextBox();
            this.vx0 = new System.Windows.Forms.TextBox();
            this.label27 = new System.Windows.Forms.Label();
            this.label28 = new System.Windows.Forms.Label();
            this.tbYi = new System.Windows.Forms.TextBox();
            this.label25 = new System.Windows.Forms.Label();
            this.label26 = new System.Windows.Forms.Label();
            this.tbXi = new System.Windows.Forms.TextBox();
            this.scaleXTxt = new System.Windows.Forms.TextBox();
            this.btnScale = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.btnExcel = new System.Windows.Forms.Button();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.lblTransfer = new System.Windows.Forms.Label();
            this.btnReset = new System.Windows.Forms.Button();
            this.label6 = new System.Windows.Forms.Label();
            this.scaleYTxt = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.chart)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.SuspendLayout();
            // 
            // pictureBox1
            // 
            this.pictureBox1.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.pictureBox1.Location = new System.Drawing.Point(22, 12);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(849, 938);
            this.pictureBox1.TabIndex = 0;
            this.pictureBox1.TabStop = false;
            this.pictureBox1.Paint += new System.Windows.Forms.PaintEventHandler(this.pictureBox1_Paint);
            // 
            // btnStart
            // 
            this.btnStart.Location = new System.Drawing.Point(1160, 464);
            this.btnStart.Name = "btnStart";
            this.btnStart.Size = new System.Drawing.Size(75, 23);
            this.btnStart.TabIndex = 3;
            this.btnStart.Text = "Start";
            this.btnStart.UseVisualStyleBackColor = true;
            this.btnStart.Click += new System.EventHandler(this.btnStart_Click);
            // 
            // timer1
            // 
            this.timer1.Interval = 10;
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // lblX
            // 
            this.lblX.AutoSize = true;
            this.lblX.Location = new System.Drawing.Point(989, 472);
            this.lblX.Name = "lblX";
            this.lblX.Size = new System.Drawing.Size(17, 13);
            this.lblX.TabIndex = 4;
            this.lblX.Text = "X:";
            // 
            // lblY
            // 
            this.lblY.AutoSize = true;
            this.lblY.Location = new System.Drawing.Point(989, 495);
            this.lblY.Name = "lblY";
            this.lblY.Size = new System.Drawing.Size(17, 13);
            this.lblY.TabIndex = 5;
            this.lblY.Text = "Y:";
            // 
            // chart
            // 
            chartArea1.Name = "ChartArea1";
            this.chart.ChartAreas.Add(chartArea1);
            legend1.Name = "Legend1";
            this.chart.Legends.Add(legend1);
            this.chart.Location = new System.Drawing.Point(877, 12);
            this.chart.Name = "chart";
            series1.ChartArea = "ChartArea1";
            series1.Legend = "Legend1";
            series1.Name = "Series1";
            this.chart.Series.Add(series1);
            this.chart.Size = new System.Drawing.Size(459, 300);
            this.chart.TabIndex = 7;
            // 
            // groupBox1
            // 
            this.groupBox1.BackColor = System.Drawing.Color.Silver;
            this.groupBox1.Controls.Add(this.label8);
            this.groupBox1.Controls.Add(this.label9);
            this.groupBox1.Controls.Add(this.tbH);
            this.groupBox1.Controls.Add(this.label29);
            this.groupBox1.Controls.Add(this.label30);
            this.groupBox1.Controls.Add(this.tbG);
            this.groupBox1.Controls.Add(this.label20);
            this.groupBox1.Controls.Add(this.label10);
            this.groupBox1.Controls.Add(this.tbk);
            this.groupBox1.Controls.Add(this.label18);
            this.groupBox1.Controls.Add(this.label13);
            this.groupBox1.Controls.Add(this.label15);
            this.groupBox1.Controls.Add(this.tbm);
            this.groupBox1.Location = new System.Drawing.Point(986, 318);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(168, 143);
            this.groupBox1.TabIndex = 22;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Constants";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(114, 111);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(15, 13);
            this.label8.TabIndex = 44;
            this.label8.Text = "m";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(8, 104);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(13, 13);
            this.label9.TabIndex = 43;
            this.label9.Text = "h";
            // 
            // tbH
            // 
            this.tbH.Location = new System.Drawing.Point(47, 104);
            this.tbH.Name = "tbH";
            this.tbH.Size = new System.Drawing.Size(57, 20);
            this.tbH.TabIndex = 42;
            this.tbH.Text = "0.2";
            this.tbH.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label29
            // 
            this.label29.AutoSize = true;
            this.label29.Location = new System.Drawing.Point(111, 78);
            this.label29.Name = "label29";
            this.label29.Size = new System.Drawing.Size(49, 13);
            this.label29.TabIndex = 25;
            this.label29.Text = "m/sec^2";
            // 
            // label30
            // 
            this.label30.AutoSize = true;
            this.label30.Location = new System.Drawing.Point(8, 71);
            this.label30.Name = "label30";
            this.label30.Size = new System.Drawing.Size(13, 13);
            this.label30.TabIndex = 24;
            this.label30.Text = "g";
            // 
            // tbG
            // 
            this.tbG.Location = new System.Drawing.Point(47, 71);
            this.tbG.Name = "tbG";
            this.tbG.Size = new System.Drawing.Size(57, 20);
            this.tbG.TabIndex = 23;
            this.tbG.Text = "9.81";
            this.tbG.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label20
            // 
            this.label20.AutoSize = true;
            this.label20.Location = new System.Drawing.Point(111, 48);
            this.label20.Name = "label20";
            this.label20.Size = new System.Drawing.Size(32, 13);
            this.label20.TabIndex = 22;
            this.label20.Text = "N/kg";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(7, 48);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(13, 13);
            this.label10.TabIndex = 21;
            this.label10.Text = "k";
            // 
            // tbk
            // 
            this.tbk.Location = new System.Drawing.Point(47, 45);
            this.tbk.Name = "tbk";
            this.tbk.Size = new System.Drawing.Size(57, 20);
            this.tbk.TabIndex = 20;
            this.tbk.Text = "2";
            this.tbk.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label18
            // 
            this.label18.AutoSize = true;
            this.label18.Location = new System.Drawing.Point(110, 22);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(19, 13);
            this.label18.TabIndex = 17;
            this.label18.Text = "kg";
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(6, 74);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(0, 13);
            this.label13.TabIndex = 6;
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Location = new System.Drawing.Point(6, 22);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(38, 13);
            this.label15.TabIndex = 2;
            this.label15.Text = "Masse";
            // 
            // tbm
            // 
            this.tbm.Location = new System.Drawing.Point(47, 19);
            this.tbm.Name = "tbm";
            this.tbm.Size = new System.Drawing.Size(57, 20);
            this.tbm.TabIndex = 2;
            this.tbm.Text = "1";
            this.tbm.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // groupBox3
            // 
            this.groupBox3.BackColor = System.Drawing.Color.Silver;
            this.groupBox3.Controls.Add(this.label4);
            this.groupBox3.Controls.Add(this.label3);
            this.groupBox3.Controls.Add(this.label2);
            this.groupBox3.Controls.Add(this.label1);
            this.groupBox3.Controls.Add(this.vy0);
            this.groupBox3.Controls.Add(this.vx0);
            this.groupBox3.Controls.Add(this.label27);
            this.groupBox3.Controls.Add(this.label28);
            this.groupBox3.Controls.Add(this.tbYi);
            this.groupBox3.Controls.Add(this.label25);
            this.groupBox3.Controls.Add(this.label26);
            this.groupBox3.Controls.Add(this.tbXi);
            this.groupBox3.Location = new System.Drawing.Point(1160, 318);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(176, 143);
            this.groupBox3.TabIndex = 24;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Initial Values";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(110, 100);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(25, 13);
            this.label4.TabIndex = 38;
            this.label4.Text = "m/s";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(110, 78);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(25, 13);
            this.label3.TabIndex = 37;
            this.label3.Text = "m/s";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(11, 104);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(20, 13);
            this.label2.TabIndex = 36;
            this.label2.Text = "vyi";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(11, 78);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(20, 13);
            this.label1.TabIndex = 35;
            this.label1.Text = "vxi";
            // 
            // vy0
            // 
            this.vy0.Location = new System.Drawing.Point(47, 97);
            this.vy0.Name = "vy0";
            this.vy0.Size = new System.Drawing.Size(57, 20);
            this.vy0.TabIndex = 34;
            this.vy0.Text = "0";
            this.vy0.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // vx0
            // 
            this.vx0.Location = new System.Drawing.Point(47, 71);
            this.vx0.Name = "vx0";
            this.vx0.Size = new System.Drawing.Size(57, 20);
            this.vx0.TabIndex = 33;
            this.vx0.Text = "0";
            this.vx0.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label27
            // 
            this.label27.AutoSize = true;
            this.label27.Location = new System.Drawing.Point(111, 48);
            this.label27.Name = "label27";
            this.label27.Size = new System.Drawing.Size(15, 13);
            this.label27.TabIndex = 32;
            this.label27.Text = "m";
            // 
            // label28
            // 
            this.label28.AutoSize = true;
            this.label28.Location = new System.Drawing.Point(15, 52);
            this.label28.Name = "label28";
            this.label28.Size = new System.Drawing.Size(16, 13);
            this.label28.TabIndex = 31;
            this.label28.Text = "Yi";
            // 
            // tbYi
            // 
            this.tbYi.Location = new System.Drawing.Point(48, 45);
            this.tbYi.Name = "tbYi";
            this.tbYi.Size = new System.Drawing.Size(57, 20);
            this.tbYi.TabIndex = 30;
            this.tbYi.Text = "1";
            this.tbYi.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label25
            // 
            this.label25.AutoSize = true;
            this.label25.Location = new System.Drawing.Point(110, 22);
            this.label25.Name = "label25";
            this.label25.Size = new System.Drawing.Size(15, 13);
            this.label25.TabIndex = 29;
            this.label25.Text = "m";
            // 
            // label26
            // 
            this.label26.AutoSize = true;
            this.label26.Location = new System.Drawing.Point(15, 22);
            this.label26.Name = "label26";
            this.label26.Size = new System.Drawing.Size(16, 13);
            this.label26.TabIndex = 28;
            this.label26.Text = "Xi";
            // 
            // tbXi
            // 
            this.tbXi.Location = new System.Drawing.Point(47, 19);
            this.tbXi.Name = "tbXi";
            this.tbXi.Size = new System.Drawing.Size(57, 20);
            this.tbXi.TabIndex = 27;
            this.tbXi.Text = "1";
            this.tbXi.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // scaleXTxt
            // 
            this.scaleXTxt.Location = new System.Drawing.Point(877, 340);
            this.scaleXTxt.Name = "scaleXTxt";
            this.scaleXTxt.Size = new System.Drawing.Size(100, 20);
            this.scaleXTxt.TabIndex = 25;
            this.scaleXTxt.Text = "100";
            // 
            // btnScale
            // 
            this.btnScale.Location = new System.Drawing.Point(877, 402);
            this.btnScale.Name = "btnScale";
            this.btnScale.Size = new System.Drawing.Size(75, 23);
            this.btnScale.TabIndex = 26;
            this.btnScale.Text = "Scale";
            this.btnScale.UseVisualStyleBackColor = true;
            this.btnScale.Click += new System.EventHandler(this.btnScale_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(877, 324);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(44, 13);
            this.label5.TabIndex = 27;
            this.label5.Text = "Scale X";
            // 
            // btnExcel
            // 
            this.btnExcel.Location = new System.Drawing.Point(1161, 512);
            this.btnExcel.Name = "btnExcel";
            this.btnExcel.Size = new System.Drawing.Size(104, 23);
            this.btnExcel.TabIndex = 28;
            this.btnExcel.Text = "Record to Excel";
            this.btnExcel.UseVisualStyleBackColor = true;
            this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
            // 
            // saveFileDialog1
            // 
            this.saveFileDialog1.Filter = "Fichier Excel|*.xls";
            // 
            // lblTransfer
            // 
            this.lblTransfer.AutoSize = true;
            this.lblTransfer.Font = new System.Drawing.Font("Microsoft Sans Serif", 27.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTransfer.Location = new System.Drawing.Point(141, 495);
            this.lblTransfer.Name = "lblTransfer";
            this.lblTransfer.Size = new System.Drawing.Size(624, 42);
            this.lblTransfer.TabIndex = 29;
            this.lblTransfer.Text = "Transfering to Excel Please Wait...";
            this.lblTransfer.Visible = false;
            // 
            // btnReset
            // 
            this.btnReset.Location = new System.Drawing.Point(1161, 489);
            this.btnReset.Name = "btnReset";
            this.btnReset.Size = new System.Drawing.Size(75, 23);
            this.btnReset.TabIndex = 30;
            this.btnReset.Text = "Reset";
            this.btnReset.UseVisualStyleBackColor = true;
            this.btnReset.Click += new System.EventHandler(this.btnReset_Click);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(877, 363);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(44, 13);
            this.label6.TabIndex = 31;
            this.label6.Text = "Scale Y";
            // 
            // scaleYTxt
            // 
            this.scaleYTxt.Location = new System.Drawing.Point(877, 379);
            this.scaleYTxt.Name = "scaleYTxt";
            this.scaleYTxt.Size = new System.Drawing.Size(100, 20);
            this.scaleYTxt.TabIndex = 32;
            this.scaleYTxt.Text = "100";
            // 
            // Simulation
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1356, 962);
            this.Controls.Add(this.scaleYTxt);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.btnReset);
            this.Controls.Add(this.lblTransfer);
            this.Controls.Add(this.btnExcel);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.btnScale);
            this.Controls.Add(this.scaleXTxt);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.chart);
            this.Controls.Add(this.lblY);
            this.Controls.Add(this.lblX);
            this.Controls.Add(this.btnStart);
            this.Controls.Add(this.pictureBox1);
            this.Name = "Simulation";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.chart)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Button btnStart;
        private System.Windows.Forms.Timer timer1;
        private System.Windows.Forms.Label lblX;
        private System.Windows.Forms.Label lblY;
        private System.Windows.Forms.DataVisualization.Charting.Chart chart;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.TextBox tbH;
        private System.Windows.Forms.Label label29;
        private System.Windows.Forms.Label label30;
        private System.Windows.Forms.TextBox tbG;
        private System.Windows.Forms.Label label20;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.TextBox tbk;
        private System.Windows.Forms.Label label18;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.TextBox tbm;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox vy0;
        private System.Windows.Forms.TextBox vx0;
        private System.Windows.Forms.Label label27;
        private System.Windows.Forms.Label label28;
        private System.Windows.Forms.TextBox tbYi;
        private System.Windows.Forms.Label label25;
        private System.Windows.Forms.Label label26;
        private System.Windows.Forms.TextBox tbXi;
        private System.Windows.Forms.TextBox scaleXTxt;
        private System.Windows.Forms.Button btnScale;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button btnExcel;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.Label lblTransfer;
        private System.Windows.Forms.Button btnReset;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox scaleYTxt;
    }
}

