namespace BalaReva.Excel.Charts
{
    using BalaReva.Excel.Interfaces;
    using System.Activities;
    using System.ComponentModel;

    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class ShowOptions : IShowOptions
    {
        [Category("Input")]
        public bool AutoText { get; set; } = true;

        [Category("Input")]
        [DisplayName("Show Ledgend")]
        [RequiredArgument]
        public bool ShowLedgend { get; set; } = false;

        [Category("Input")]
        [Description("Show CategoryName")]
        [DisplayName("Show CategoryName")]
        public bool ShowCategoryName { get; set; } = false;

        [Category("Input")]
        [Description("Show Percentage")]
        [DisplayName("Show Percentage")]
        public bool ShowPercentage { get; set; } = false;

        [Category("Input")]
        [RequiredArgument]
        [DisplayName("Data Labels Type")]
        public DataLabelsEnum DataLabelsType { get; set; } = DataLabelsEnum.ShowPercent;

        [Category("Input")]
        [RequiredArgument]
        [DisplayName("Ledgend Key")]
        public DataLabelsEnum LedgendKey { get; set; } = DataLabelsEnum.ShowLabel;

        [Category("Input")]
        [DisplayName("Has Leader Lines")]
        public bool HasLeaderLines { get; set; } = false;

        [Category("Input")]
        [DisplayName("Show Series Name")]
        public bool ShowSeriesName { get; set; } = false;

        [Category("Input")]
        [DisplayName("Show Value")]
        public bool ShowValue { get; set; } = false;

        [Category("Input")]
        [DisplayName("Show Bubble Size")]
        public bool ShowBubbleSize { get; set; } = false;


        [Category("Input")]
        public string Separator { get; set; } = "";
    }
}
