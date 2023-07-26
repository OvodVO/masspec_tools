using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WashU.BatemanLab.MassSpec.TrackIN
{
    public partial class NAmsrunsTabControl : TabControl
    {
        public NAmsrunsTabControl()
        {
            InitializeComponent();
            
            Multiline = false;
            Alignment = TabAlignment.Bottom;
            //SizeMode = TabSizeMode.Fixed;
            //_tabMsRuns.ItemSize = new System.Drawing.Size(30, 120);
            //_tabMsRuns.Width = 30;
            //_tabMsRuns.Height = 120;
            DrawMode = TabDrawMode.OwnerDrawFixed;
            Dock = DockStyle.Fill;
        }

        protected override void OnPaint(PaintEventArgs pe)
        {
            base.OnPaint(pe);
        }

        protected override void OnDrawItem(DrawItemEventArgs e)
        {
            using (var _textBrush = new SolidBrush(this.ForeColor))
            {
                TabPage _tabPage = this.TabPages[e.Index];
                Rectangle _tabBounds = this.GetTabRect(e.Index);

                if (e.State != DrawItemState.Selected) e.DrawBackground();
                else
                {
                    using (var brush = new System.Drawing.Drawing2D.LinearGradientBrush(e.Bounds, Color.White, Color.Yellow, 90f))
                    {
                        e.Graphics.FillRectangle(brush, e.Bounds);
                    }
                }

                Font _tabFont = new Font("Arial", (float)10.0, FontStyle.Bold, GraphicsUnit.Pixel);

                StringFormat _stringFlags = new StringFormat();
                _stringFlags.Alignment = StringAlignment.Center;
                _stringFlags.LineAlignment = StringAlignment.Center;
                e.Graphics.DrawString(_tabPage.Text, _tabFont, _textBrush, _tabBounds, new StringFormat(_stringFlags));
            }
        }
    }

    
}
