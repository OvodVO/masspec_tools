using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SkylineTool;
using ZedGraph;

namespace WashU.BatemanLab.MassSpec.TrackIN
{
    public class Graph
    {
        private ZedGraphControl _graphControl;
        public Graph(ZedGraphControl graphControl, string xLabel, string yLabel)
        {
            _graphControl = graphControl;
            var pane = _graphControl.GraphPane;
            pane.Title.IsVisible = false;
            pane.Border.IsVisible = false;
            pane.Chart.Border.IsVisible = false;

            pane.XAxis.Title.Text = xLabel;
            pane.XAxis.MinorTic.IsOpposite = false;
            pane.XAxis.MajorTic.IsOpposite = false;
            pane.XAxis.MajorTic.IsAllTics = false;
            pane.XAxis.Scale.FontSpec.Angle = 65;

            //test
            pane.XAxis.Scale.FontSpec.Size = 10;

            pane.XAxis.Scale.Align = AlignP.Inside;

            pane.YAxis.Title.Text = yLabel;
            pane.YAxis.MinorTic.IsOpposite = false;
            pane.YAxis.MajorTic.IsOpposite = false;
        }
    }

    public class RatioGraph : Graph
    {
        public RatioGraph(ZedGraphControl graphControl, string xLabel, string yLabel)
            : base(graphControl, xLabel, yLabel)
        {

        }
    }
}
