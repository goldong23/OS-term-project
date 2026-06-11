using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace Memory_Policy_Simulator
{
    public partial class Form1 : Form
    {
        Graphics g;
        PictureBox pbPlaceHolder;
        Bitmap bResultImage;

        public Form1()
        {
            InitializeComponent();
            if (this.comboBox1.SelectedIndex < 0)
            {
                this.comboBox1.SelectedIndex = 0;
            }
            this.pbPlaceHolder = new PictureBox();
            this.bResultImage = new Bitmap(2048, 2048);
            this.pbPlaceHolder.Size = new Size(2048, 2048);
            g = Graphics.FromImage(this.bResultImage);
            pbPlaceHolder.Image = this.bResultImage;
            this.pImage.Controls.Add(this.pbPlaceHolder);
            this.tbConsole.Multiline = true;
            this.tbConsole.ScrollBars = ScrollBars.Vertical;
            this.comboBox1.SelectedIndexChanged += new EventHandler(this.comboBox1_SelectedIndexChanged);
            this.comboBox1.TextChanged += new EventHandler(this.comboBox1_SelectedIndexChanged);
            UpdatePolicyOptions();
        }

        private void DrawBase(Core core, int windowSize, int dataLength)
        {
            g.Clear(Color.Black);

            for ( int i = 0; i < dataLength; i++ ) // length
            {
                int cursor = core.pageHistory[i].loc;
                char data = core.pageHistory[i].data;
                Page.STATUS status = core.pageHistory[i].status;

                for ( int j = 0; j <= windowSize; j++) // height - STEP
                {
                    if (j == 0)
                    {
                        DrawGridText(i, j, data);
                    }
                    else
                    {
                        DrawGrid(i, j);
                    }
                }

                DrawGridHighlight(i, cursor, status);

                string[] frames = core.pageHistory[i].frameSnapshot.Split(' ');
                for (int depth = 1; depth <= windowSize; depth++)
                {
                    if (depth - 1 < frames.Length && frames[depth - 1] != "-")
                    {
                        DrawGridText(i, depth, frames[depth - 1][0]);
                    }
                }
            }
        }


        private void DrawGrid(int x, int y)
        {
            int gridSize = 30;
            int gridSpace = 5;
            int gridBaseX = x * gridSize;
            int gridBaseY = y * gridSize;

            g.DrawRectangle(new Pen(Color.White), new Rectangle(
                gridBaseX + (x * gridSpace),
                gridBaseY,
                gridSize,
                gridSize
                ));
        }

        private void DrawGridHighlight(int x, int y, Page.STATUS status)
        {
            int gridSize = 30;
            int gridSpace = 5;
            int gridBaseX = x * gridSize;
            int gridBaseY = y * gridSize;

            SolidBrush highlighter = new SolidBrush(Color.LimeGreen);

            switch (status)
            {
                case Page.STATUS.HIT:
                    break;
                case Page.STATUS.MIGRATION:
                    highlighter.Color = Color.Purple;
                    break;
                case Page.STATUS.PAGEFAULT:
                    highlighter.Color = Color.Red;
                    break;
            }

            g.FillRectangle(highlighter, new Rectangle(
                gridBaseX + (x * gridSpace),
                gridBaseY,
                gridSize,
                gridSize
                ));
        }

        private void DrawGridText(int x, int y, char value)
        {
            int gridSize = 30;
            int gridSpace = 5;
            int gridBaseX = x * gridSize;
            int gridBaseY = y * gridSize;

            g.DrawString(
                value.ToString(), 
                new Font(FontFamily.GenericMonospace, 8), 
                new SolidBrush(Color.White), 
                new PointF(
                    gridBaseX + (x * gridSpace) + gridSize / 3,
                    gridBaseY + gridSize / 4));
        }

        private void btnOperate_Click(object sender, EventArgs e)
        {
            this.tbConsole.Clear();

            if (this.tbQueryString.Text != "" && this.tbWindowSize.Text != "")
            {
                string data = this.tbQueryString.Text;
                int windowSize;

                if (!int.TryParse(this.tbWindowSize.Text, out windowSize) || windowSize <= 0)
                {
                    MessageBox.Show("Frame size must be a positive number.");
                    return;
                }

                /* initalize */
                Core.ReplacementPolicy policy = Core.ParsePolicy(this.comboBox1.Text);
                int clock = 0;
                string modifiedPages = "";

                if (UsesClock(policy))
                {
                    if (!int.TryParse(this.tbResetInterval.Text, out clock) || clock < 0)
                    {
                        MessageBox.Show("Clock must be zero or a positive number.");
                        return;
                    }
                }

                if (UsesModifiedPages(policy))
                {
                    modifiedPages = this.tbModifiedPages.Text;
                }

                var window = new Core(windowSize, policy, clock, modifiedPages);
                Stopwatch stopwatch = Stopwatch.StartNew();

                for (int i = 0; i < data.Length; i++)
                {
                    char element = data[i];
                    var status = window.Operate(element, i, data);
                    Page current = window.pageHistory[i];
                    string victimText = current.hasVictim ? ", victim=" + current.victim : "";
                    this.tbConsole.Text += "[" + (i + 1) + "] DATA " + element + " is " +
                        ((status == Page.STATUS.PAGEFAULT) ? "Page Fault" : status == Page.STATUS.MIGRATION ? "Migrated" : "Hit")
                        + victimText + ", frames=[" + current.frameSnapshot + "]"
                        + ", state={" + current.algorithmState + "}\r\n";
                }

                stopwatch.Stop();
                DrawBase(window, windowSize, data.Length);
                this.pbPlaceHolder.Refresh();

                int total = window.hit + window.fault;
                float faultRate = total == 0 ? 0 : (float)window.fault / total;

                this.tbConsole.Text += "\r\nPolicy: " + this.comboBox1.Text + "\r\n";
                this.tbConsole.Text += "Policy Detail: " + window.GetPolicyDescription() + "\r\n";
                if (UsesClock(policy))
                {
                    this.tbConsole.Text += "Clock: " + clock + "\r\n";
                }
                if (UsesModifiedPages(policy))
                {
                    this.tbConsole.Text += "Modified Pages Input: " + (modifiedPages == "" ? "(none)" : modifiedPages) + "\r\n";
                }
                this.tbConsole.Text += "Hit Count: " + window.hit + "\r\n";
                this.tbConsole.Text += "Page Fault Count: " + window.fault + "\r\n";
                this.tbConsole.Text += "Migration Count: " + window.migration + "\r\n";
                this.tbConsole.Text += "Page Fault Rate: " + Math.Round(faultRate * 100, 2) + "%\r\n";
                this.tbConsole.Text += "Execution Time: " + stopwatch.Elapsed.TotalMilliseconds.ToString("0.###") + " ms\r\n";
                this.tbConsole.Text += "Estimated Delay by Page Fault: " + window.GetEstimatedPageFaultDelay() + " ms\r\n";

                /* 차트 생성 */
                chart1.Series.Clear();
                Series resultChartContent = chart1.Series.Add("Statics");
                resultChartContent.ChartType = SeriesChartType.Pie;
                resultChartContent.IsVisibleInLegend = true;
                resultChartContent.Points.AddXY("Hit", window.hit);
                resultChartContent.Points.AddXY("Fault", window.fault);
                resultChartContent.Points[0].IsValueShownAsLabel = true;
                resultChartContent.Points[0].LegendText = "Hit " + window.hit;
                resultChartContent.Points[1].IsValueShownAsLabel = true;
                resultChartContent.Points[1].LegendText = "Fault " + window.fault + " (Migrated " + window.migration + ")";

                this.lbPageFaultRatio.Text = Math.Round(faultRate * 100, 2) + "%";
            }
            else
            {
                MessageBox.Show("Please enter both a reference string and frame size.");
            }

        }

        private void pbPlaceHolder_Paint(object sender, PaintEventArgs e)
        {
        }

        private void chart1_Click(object sender, EventArgs e)
        {

        }

        private void tbWindowSize_KeyDown(object sender, KeyEventArgs e)
        {

        }

        private void tbWindowSize_KeyPress(object sender, KeyPressEventArgs e)
        {
                if (!(Char.IsDigit(e.KeyChar)) && e.KeyChar != 8)
                {
                    e.Handled = true;
                }
        }

        private void tbNumeric_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(Char.IsDigit(e.KeyChar)) && e.KeyChar != 8)
            {
                e.Handled = true;
            }
        }

        private void btnRand_Click(object sender, EventArgs e)
        {
            Random rd = new Random();

            int count = rd.Next(5, 50);
            StringBuilder sb = new StringBuilder();


            for ( int i = 0; i < count; i++ )
            {
                sb.Append((char)rd.Next(65, 90));
            }

            this.tbQueryString.Text = sb.ToString();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            bResultImage.Save("./result.jpg");
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdatePolicyOptions();
        }

        private void UpdatePolicyOptions()
        {
            Core.ReplacementPolicy policy = Core.ParsePolicy(this.comboBox1.Text);
            bool usesClock = UsesClock(policy);
            bool usesModifiedPages = UsesModifiedPages(policy);

            this.label6.Visible = usesClock;
            this.tbResetInterval.Visible = usesClock;
            this.label7.Visible = usesModifiedPages;
            this.tbModifiedPages.Visible = usesModifiedPages;

            this.label6.Text = "Clock";
            this.label7.Text = "Modified";
            this.label6.Location = new Point(650, 1);
            this.tbResetInterval.Location = new Point(644, 33);
            this.tbResetInterval.Width = 70;
            this.label7.Location = new Point(725, 1);
            this.tbModifiedPages.Location = new Point(724, 33);
            this.tbModifiedPages.Width = 152;
        }

        private static bool UsesClock(Core.ReplacementPolicy policy)
        {
            return policy == Core.ReplacementPolicy.NUR_01_First
                || policy == Core.ReplacementPolicy.NUR_10_First
                || policy == Core.ReplacementPolicy.SecondChance
                || policy == Core.ReplacementPolicy.WSClockLite;
        }

        private static bool UsesModifiedPages(Core.ReplacementPolicy policy)
        {
            return policy == Core.ReplacementPolicy.NUR_01_First
                || policy == Core.ReplacementPolicy.NUR_10_First
                || policy == Core.ReplacementPolicy.WSClockLite;
        }
    }
}
