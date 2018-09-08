namespace Formater
{
    using System;
    using System.Activities;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Globalization;
    using System.Linq;

    public class DoubleFormater : CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        [Description("Enter the double Value")]
        public InArgument<double> InputValue { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [Description("Format separated by comma(###,##)")]
        public InArgument<string> Format { get; set; }

        [Category("Input")]
        [Description("Default dot(.)")]
        [DefaultValue(".")]
        public InArgument<string> DecimalSepartor { get; set; } = ".";

        [Category("Input")]
        [Description("Default comma(,)")]
        [DefaultValue(",")]
        public InArgument<string> GroupSepartor { get; set; } = ",";

        [Category("Input")]
        [Description("Sample Format ##")]
        public InArgument<string> AfterDecimalFormat { get; set; } = "##";

        [Category("OutPut")]
        public OutArgument<string> Result { get; set; }

        public DoubleFormater()
        {
            this.DecimalSepartor = ".";
            this.GroupSepartor = ",";
            this.AfterDecimalFormat = "##";
        }

        private double InternalInputValue { get; set; }
        private string InternalFormat { get; set; }
        private string InternalDecimalSepartor { get; set; }
        private string InternalGroupSepartor { get; set; }
        private string InternalAfterDecimalCount { get; set; }
        private string InternalResult { get; set; }

        private void SetValues(CodeActivityContext context)
        {
            this.InternalInputValue = InputValue.Get(context);
            this.InternalFormat = Format.Get(context);

            this.InternalDecimalSepartor = DecimalSepartor.Get(context);

            if (this.InternalDecimalSepartor == null || string.IsNullOrEmpty(this.InternalDecimalSepartor))
            {
                this.InternalDecimalSepartor = ".";
            }

            this.InternalGroupSepartor = GroupSepartor.Get(context);

            if (this.InternalGroupSepartor == null || string.IsNullOrEmpty(this.InternalGroupSepartor))
            {
                this.InternalGroupSepartor = ",";
            }

            this.InternalAfterDecimalCount = AfterDecimalFormat.Get(context);

            if (this.InternalAfterDecimalCount == null || string.IsNullOrEmpty(this.InternalAfterDecimalCount))
            {
                this.InternalAfterDecimalCount = string.Empty;
            }
        }


        protected override void Execute(CodeActivityContext context)
        {
            this.SetValues(context);

            if (InternalInputValue != 0)
            {
                Result.Set(context, DoubleOuptPut());
            }
            else
            {
                Result.Set(context, InternalInputValue.ToString());
            }
        }

        private string DoubleOuptPut()
        {
            string result = string.Empty;

            Int32 decLength = 0;
            CultureInfo cultureInfo = new CultureInfo(CultureInfo.CurrentCulture.Name);

            NumberFormatInfo numberFormat = cultureInfo.NumberFormat;
            numberFormat.NumberDecimalSeparator = InternalDecimalSepartor;
            numberFormat.NumberGroupSeparator = InternalGroupSepartor;

            List<string> strformat = InternalFormat.Split(',').ToList();
            List<Int32> NumberArr = strformat.Select(a => a.Length).ToList();
            NumberArr = NumberArr.AsEnumerable().Reverse().ToList();
            NumberArr.RemoveAt(NumberArr.Count - 1);
            NumberArr.Add(0);
            numberFormat.NumberGroupSizes = NumberArr.ToArray();

            decLength = InternalAfterDecimalCount.Length;

            if (decLength == 0)
            {
                InternalInputValue = Math.Round(InternalInputValue,0);
            }

            result = InternalInputValue.ToString("N" + decLength, numberFormat);

            return result;
        }
    }
}
