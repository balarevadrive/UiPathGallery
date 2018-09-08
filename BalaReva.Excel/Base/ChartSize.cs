namespace BalaReva.Excel.Charts
{
    using BalaReva.Excel.Interfaces;
    using System.Activities;
    using System.ComponentModel;

    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class ChartSize : IChartSize
    {
        [RequiredArgument]
        public double Left { get; set; } = 50;

        [RequiredArgument]
        public double Top { get; set; } = 50;

        [RequiredArgument]
        public double Width { get; set; } = 250;

        [RequiredArgument]
        public double Height { get; set; } = 250;
    }
}
