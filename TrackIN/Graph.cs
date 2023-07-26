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
        public ZedGraphControl GraphControl
        {
            get { return _graphControl; }
        }
    
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

            //pane.XAxis.MajorTic.IsAllTics = false;
            //pane.XAxis.Scale.FontSpec.Angle = 65;

            //test
            //pane.XAxis.Scale.FontSpec.Size = 10;

            //pane.XAxis.Scale.Align = AlignP.Inside;

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

    public class NoiseGraph : Graph
    {
        public NoiseGraph(ZedGraphControl graphControl, string xLabel, string yLabel)
            : base(graphControl, xLabel, yLabel)
        {

        }
    }

    public class MasspecDataGraph
    {
        private ZedGraphControl _graphControl;
        public ZedGraphControl GraphControl
        {
            get { return _graphControl; }
        }
        public MasspecDataGraph(string xLabel, string yLabel)
        {
            _graphControl = new ZedGraphControl();
            var pane = _graphControl.GraphPane;

            pane.Title.IsVisible = false;
            pane.Border.IsVisible = false;
            pane.Chart.Border.IsVisible = false;

            pane.XAxis.Title.Text = xLabel;
            pane.XAxis.MinorTic.IsOpposite = false;
            pane.XAxis.MajorTic.IsOpposite = false;

            pane.YAxis.Title.Text = yLabel;
            pane.YAxis.MinorTic.IsOpposite = false;
            pane.YAxis.MajorTic.IsOpposite = false;

            pane.YAxis.Scale.Min = 0;
            pane.Y2Axis.Scale.Min = 0;

            //_graphControl.IsEnableVZoom = false;
            _graphControl.IsEnableVPan = false;

        }
    }

    public class NoiseMasspecDataGraph : MasspecDataGraph
    {
        public NoiseMasspecDataGraph(string xLabel, string yLabel, string y2Label)  : base(xLabel, yLabel)
        {
            var pane = GraphControl.GraphPane;
            pane.Y2Axis.Title.Text = y2Label;
            pane.Y2Axis.MinorTic.IsOpposite = false;
            pane.Y2Axis.MajorTic.IsOpposite = false;

            GraphPane istdPane = new GraphPane();

            istdPane.Title.Text = "ISTD";

            istdPane.Title.IsVisible = false;
            istdPane.Border.IsVisible = false;
            istdPane.Chart.Border.IsVisible = false;

            istdPane.XAxis.Title.Text = xLabel;
            istdPane.XAxis.MinorTic.IsOpposite = false;
            istdPane.XAxis.MajorTic.IsOpposite = false;

            istdPane.YAxis.Title.Text = yLabel;
            istdPane.YAxis.MinorTic.IsOpposite = false;
            istdPane.YAxis.MajorTic.IsOpposite = false;


            istdPane.Y2Axis.Title.Text = y2Label;
            istdPane.Y2Axis.MinorTic.IsOpposite = false;
            istdPane.Y2Axis.MajorTic.IsOpposite = false;


            GraphControl.MasterPane.Add(istdPane);



            GraphControl.IsSynchronizeYAxes = true;
            GraphControl.IsSynchronizeXAxes = true;

        }
    }
}
