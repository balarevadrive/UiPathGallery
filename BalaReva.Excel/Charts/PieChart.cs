namespace BalaReva.Excel.Charts
{
    using BalaReva.Excel.Design;
    using BalaReva.Excel.Interfaces;
    using Microsoft.Office.Interop.Excel;
    using System;
    using System.Activities;
    using System.ComponentModel;

    [DisplayName("Pie Chart")]
    [Designer(typeof(PieChartDesigner))]
    public class PieChart : BaseChart, IPieChart
    {
        [Category("Input")]
        [DisplayName("Chart Type")]
        public PieChartEnum ChartType { get; set; } = PieChartEnum.Pie;

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
