using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Drawing.Drawing2D;
using System.Windows.Forms.DataVisualization.Charting;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;

namespace WindowsFormsApplication1
{
    public partial class Simulation : Form
    {
        Excel.Application xlApp; 
        Excel.Workbook xlWorkBook;
        Excel.Worksheet xlWorkSheetData; 
        Excel.Range xlRng; 
        Excel.WorksheetFunction xlFunction;
        Stopwatch stopwatch = new Stopwatch();
        List<calculationsPackage> excelData = new List<calculationsPackage>();
        DataTable dataTable = new DataTable("Pendulum");

        public bool start = true;
        public bool cont = false;
        public int circleDiamater = 40;
        public double L = 0.5;
        public double x = 0;
        public double y = 0;
        public Stopwatch stopWatch = new Stopwatch();
        double ay;
        double ax;
        double vx;
        double vy;
        double dt ;
        double g;
        double m;
        double k;
        double h;
        double Xi;
        double Yi;
        double vxi;
        double vyi;
        int lastRowExcel = 3;
        double x1;
        double y1 ;
        double scaleValueX = 100;
        double scaleValueY = 100;
        double drawX;
        double drawY;

        struct calculationsPackage
        {
            public double time;
            public double x;
            public double y;
            public double vx;
            public double vy;
            public calculationsPackage(double time, double x, double y, double vx, double vy)
            {
                this.time = time;
                this.x = x;
                this.y = y;
                this.vx = vx;
                this.vy = vy;
            }
        }
        public Simulation()
        {
            InitializeComponent();
        }
        private void prepareInitValeurs()
        {
            g = Double.Parse(tbG.Text);
            m = Double.Parse(tbm.Text);
            k = Double.Parse(tbk.Text);
          
            h = Double.Parse(tbH.Text);
            
            Xi = Double.Parse(tbXi.Text);
            Yi = Double.Parse(tbYi.Text);
            x1 = Xi;
            y1 = Yi;
            vxi = double.Parse(vx0.Text);
            vyi = double.Parse(vx0.Text);
            

        }

        private void pictureBox1_Paint(object sender, PaintEventArgs e)
        {
            shiftCartesianCordinates();
            GraphicsPath myPath = new GraphicsPath();
            myPath.AddLine(pictureBox1.Width/2 , 0, (int)drawX, (int)drawY);
            Rectangle pathRect = new Rectangle((int)(drawX - circleDiamater / 4), (int)(drawY - circleDiamater / 4), circleDiamater/2, circleDiamater/2);
            SolidBrush blueBrush = new SolidBrush(Color.Red);
            Pen myPen = new Pen(Color.Black, 1);
            e.Graphics.DrawPath(myPen, myPath);
            e.Graphics.FillEllipse(blueBrush, pathRect);

        }
        public void shiftCartesianCordinates()
        {
            drawY = y1 * scaleValueY;
            if (x1 == 0)
                drawX =  pictureBox1.Width / 2;
            else
                drawX = x1 * scaleValueX + pictureBox1.Width / 2;
        }
        public void setLabels()
        {
            lblX.Text = "X: " + ((double)x1).ToString();
            lblY.Text = "Y: " + ((double)y1).ToString();
          
        }
        private void btnStart_Click(object sender, EventArgs e)
        {
            if (start && !cont)
            {
                timer1.Start();
                stopwatch.Start();
                timer1.Enabled = true;

                chart.Series.Clear(); //ensure that the chart is empty
                chart.Series.Add("Series0");
                chart.Series[0].ChartType = SeriesChartType.Line;
                chart.Legends.Clear();
                stopWatch.Start();
                prepareInitValeurs();
                start = false;
                btnStart.Text = "Stop";
                cont = true;
            }
            else if (!start)
            {
                stopwatch.Stop();
                start = true; ;
                timer1.Stop();
                btnStart.Text = "Start";

            }
            else
            {
                timer1.Start();
                stopwatch.Start();
                timer1.Enabled = true;
                start = false;
                btnStart.Text = "Stop";
                cont = true;
            }
        }
        
           
        
        private void timer1_Tick(object sender, EventArgs e)
        {
            dt = 1 / ((k+0.1)/ m) * 0.05;
            double ax0 = ax = (-k / m) * (Xi - h * Xi / (Math.Sqrt(Math.Pow(Xi, 2) + Math.Pow(Yi, 2))));
            double ay0 = (-k / m) * (Yi - h * Yi / (Math.Sqrt(Math.Pow(Yi, 2) + Math.Pow(Xi, 2))))+g;
            vx= vxi + dt* ax0;
            vy = vyi + dt * ay0;
            x1 = Xi + vx * dt + ax0 * Math.Pow(dt, 2)/2;
            y1 = Yi + vy * dt + ay0 * Math.Pow(dt, 2)/ 2;
            ax=(-k / m) * (x1 - h * x1 / (Math.Sqrt(Math.Pow(x1, 2) + Math.Pow(y1, 2))));
            ay=(-k / m) * (y1 - h * y1 / (Math.Sqrt(Math.Pow(y1, 2) + Math.Pow(x1, 2))))+g;
           
             ax0=ax;
             ay0=ay;
             vxi=vx;
             vyi=vy;
             Xi = x1;
             Yi = y1;


            pictureBox1.Invalidate();
            chart.Series[0].Points.AddXY(y1, x1);
            excelData.Add(new calculationsPackage(stopwatch.ElapsedMilliseconds,x1, y1, vx, vy));
            setLabels();

        }

        private void button1_Click(object sender, EventArgs e)
        {
            timer1.Stop();
            this.chart.SaveImage("C:\\Users\\Sahar\\Desktop.png", ChartImageFormat.Png);
            
        }

        private void btnScale_Click(object sender, EventArgs e)
        {
            scaleValueX = Convert.ToDouble(scaleXTxt.Text);
            scaleValueY = Convert.ToDouble(scaleYTxt.Text);
        }

        private void btnExcel_Click(object sender, EventArgs e)
        {
           if ( !timer1.Enabled )
                btnStart_Click(sender, e);
               object misValue = System.Reflection.Missing.Value;
                xlApp = new Excel.Application();
                xlApp.Visible = false; 
                xlWorkBook = xlApp.Workbooks.Add(misValue);
                xlFunction = xlApp.WorksheetFunction;
               

                xlWorkSheetData = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                xlWorkSheetData.Name = "Pendulum";
                xlWorkSheetData.Activate();

                xlWorkSheetData.Cells[1, 1] = "Constants used, mass: " + tbm.Text + ", gravity: " + tbG.Text + ", Spring Constant: " + tbk.Text + ", Length: " + tbH.Text;
                xlRng = xlWorkSheetData.get_Range("A1", "N1");
                xlRng.Select();
                xlRng.Merge();

                xlWorkSheetData.Cells[2, 1] = "Initial Values used, Intial X: " + tbXi.Text + ", Initial Y: " + tbYi.Text + ", Initial X Velocity: " + vx0.Text + ", Initial Y Velocity: " + vy0.Text;
                xlRng = xlWorkSheetData.get_Range("A2", "N2");
                xlRng.Select();
                xlRng.Merge();


                xlWorkSheetData.Cells[lastRowExcel, 1] = "t"; // changes these to whatever you want
                xlWorkSheetData.Cells[lastRowExcel, 2] = "X";
                xlWorkSheetData.Cells[lastRowExcel, 3] = "Y";
                xlWorkSheetData.Cells[lastRowExcel, 4] = "Vx";
                xlWorkSheetData.Cells[lastRowExcel, 5] = "Vy";
                lblTransfer.Visible = true;
                for (int i = 0; i < excelData.Count; i++)
                {
                    xlWorkSheetData.Cells[i + 4, 1] = (excelData[i].time / 1000.00).ToString();
                    xlWorkSheetData.Cells[i + 4, 2] = excelData[i].x.ToString();
                    xlWorkSheetData.Cells[i + 4, 3] = excelData[i].y.ToString();
                    xlWorkSheetData.Cells[i + 4, 4] = excelData[i].vx.ToString();
                    xlWorkSheetData.Cells[i + 4, 5] = excelData[i].vy.ToString();
                }
                lblTransfer.Visible = false;
                try //essaye le sauvegarde
                {

                    if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                    {
                        //sauvegarde le classeur courant
                        xlWorkBook.SaveAs(saveFileDialog1.FileName,
                            Excel.XlFileFormat.xlWorkbookDefault, misValue,
                            misValue, misValue, misValue,
                            Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue,
                            misValue, misValue, misValue);
                        xlWorkBook.Close();
                    }

                }
                catch //en cas d'erreur affiche le message
                {
                    MessageBox.Show("Impossible de sauvegarder le fichier.", "Erreur de sauvegarde de fichier Excel", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

        private void btnReset_Click(object sender, EventArgs e)
        {
            timer1.Enabled = false;
            stopwatch.Reset();
            excelData.Clear();
            start = true;
            prepareInitValeurs();
            chart.Series.Clear();
            ax = 0;
            ay = 0;
            x1 = Xi;
            y1 = Yi;
            pictureBox1.Invalidate();
            btnStart.Text = "Start";
         

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            prepareInitValeurs();
        }

        
    }
}
