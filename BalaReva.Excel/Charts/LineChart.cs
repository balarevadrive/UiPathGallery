namespace BalaReva.Excel.Charts
{
    using BalaReva.Excel.Design;
    using BalaReva.Excel.Interfaces;
    using Microsoft.Office.Interop.Excel;
    using System;
    using System.Activities;
    using System.ComponentModel;

    [DisplayName("Line Chart")]
    [Designer(typeof(LineChartDesigner))]
    public class LineChart : BaseChart, ILineChart
    {
        [Category("Input")]
        [DisplayName("Chart Type")]
        public LineChartEnum ChartType { get; set; } = LineChartEnum.Line;

        protected override void Execute(CodeActivityContext context)
        {
            try
            {
                // Initalize the values to base class variables
                base.InitValue(context);

                base.DrawChart((XlChartType)ChartType);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
