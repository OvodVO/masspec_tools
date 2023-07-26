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
    public partial class NApeptidesTabControl : TabControl
    {
        public NApeptidesTabControl()
        {
            InitializeComponent();
            Multiline = false;
            SizeMode = TabSizeMode.Fixed;
            ItemSize = new System.Drawing.Size(30, 120);
            Width = 30;
            Height = 120;
            DrawMode = TabDrawMode.OwnerDrawFixed;
            Alignment = TabAlignment.Left;
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

                if (e.State != DrawItemState.Selected)
                {
                    e.DrawBackground();
                    Font _tabFont = new Font("Arial", (float)10.0, FontStyle.Bold, GraphicsUnit.Pixel);
                    StringFormat _stringFlags = new StringFormat();
                    _stringFlags.Alignment = StringAlignment.Center;
                    _stringFlags.LineAlignment = StringAlignment.Center;
                    e.Graphics.DrawString(_tabPage.Text, _tabFont, _textBrush, _tabBounds, new StringFormat(_stringFlags));
                }
                else
                {
                    using (var brush = new System.Drawing.Drawing2D.LinearGradientBrush(e.Bounds, Color.White, Color.Blue, 90f))
                    {
                        e.Graphics.FillRectangle(brush, e.Bounds);
                        StringFormat _stringFlagsSel = new StringFormat();
                        _stringFlagsSel.Alignment = StringAlignment.Center;
                        _stringFlagsSel.LineAlignment = StringAlignment.Center;
                        SolidBrush _textBrushSel = new SolidBrush(Color.Yellow);
                        Font _tabFontSel = new Font("Arial", (float)16.0, FontStyle.Bold, GraphicsUnit.Pixel);
                        e.Graphics.DrawString(_tabPage.Text, _tabFontSel, _textBrushSel, _tabBounds, new StringFormat(_stringFlagsSel));
                    }
                }
            }
        }
    }
}
