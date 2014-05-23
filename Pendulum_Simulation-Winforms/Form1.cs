using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Drawing.Drawing2D;
using System.Windows.Forms.DataVisualization.Charting;
using System.Diagnostics;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        public int circleDiamater = 40;
        public double angle = 0;
        public double x = 0;
        public double y = 0;
        public double direction = -.5;
        public int pendulumShaftLength = 200;
        public Stopwatch stopWatch = new Stopwatch();

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            convertPolarToCartesian(angle);
        }

        private void pictureBox1_Paint(object sender, PaintEventArgs e)
        {
            GraphicsPath myPath = new GraphicsPath();
            myPath.AddLine(pictureBox1.Width / 2, 0, (int)x , (int)y );
            Rectangle pathRect = new Rectangle((int)(x - circleDiamater / 2), (int)(y - circleDiamater / 2), circleDiamater, circleDiamater);
            SolidBrush blueBrush = new SolidBrush(Color.Blue);
            Pen myPen = new Pen(Color.Black, 2);
            e.Graphics.DrawPath(myPen, myPath);
            e.Graphics.FillEllipse(blueBrush, pathRect);

        }
        public void convertPolarToCartesian(double angle)
        {
            x = Math.Cos(angle * Math.PI / 180) * pendulumShaftLength + pictureBox1.Width / 2;       
            y = -1 * Math.Sin(angle * Math.PI / 180) * pendulumShaftLength;
           
        }
        public void setLabels()
        {
            lblX.Text = "X: " + ((int)x).ToString();
            lblY.Text = "Y: " + ((int)y).ToString();
            lblAngle.Text = "Angle: " + ((int)angle).ToString() + '°';
        }
        private void btnStart_Click(object sender, EventArgs e)
        {
            pictureBox1.Invalidate();
            timer1.Enabled = true;

            chart.Series.Clear(); //ensure that the chart is empty
            chart.Series.Add("Series0");
            chart.Series[0].ChartType = SeriesChartType.Line;
            chart.Legends.Clear();
            stopWatch.Start();

        }
        private void timer1_Tick(object sender, EventArgs e)
        {
            if (angle <= -180) 
                direction = 1;
            else if (angle >= 0)
                direction = -1;
            angle += direction;
            convertPolarToCartesian(angle);
            setLabels();
            pictureBox1.Invalidate();
            chart.Series[0].Points.AddXY(stopWatch.ElapsedMilliseconds, angle); 
        }
    }
}
