namespace BalaReva.Excel.Charts
{
    using BalaReva.Excel.Design;
    using BalaReva.Excel.Interfaces;
    using Microsoft.Office.Interop.Excel;
    using System;
    using System.Activities;
    using System.ComponentModel;

    [DisplayName("Bar Chart")]
    [Designer(typeof(BarChartDesigner))]
    public class BarChart : BaseChart, IBarChart
    {
        [Category("Input")]
        [DisplayName("Chart Type")]
        public BarChartEnum ChartType { get; set; } = BarChartEnum.BarClustered;

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
