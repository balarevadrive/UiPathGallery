namespace Formater
{
    using System;
    using System.Activities;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Globalization;
    using System.Linq;

    public class NumberFormater : CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        [Description("Enter the integer Value")]
        public InArgument<Int64> InputValue { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [Description("Format separated by comma(###,##)")]
        public InArgument<string> Format { get; set; }

        [Category("Input")]
        [Description("Default comma(,)")]
        [DefaultValue(",")]
        public InArgument<string> GroupSepartor { get; set; } = ",";

        [Category("OutPut")]
        public OutArgument<string> Result { get; set; }

        public NumberFormater()
        {
            this.GroupSepartor = ",";
        }

        private Int64 InternalInputValue { get; set; }
        private string InternalFormat { get; set; }
        private string InternalGroupSepartor { get; set; }
        private string InternalResult { get; set; }

        private void SetValues(CodeActivityContext context)
        {
            this.InternalInputValue = InputValue.Get(context);
            this.InternalFormat = Format.Get(context);

            this.InternalGroupSepartor = GroupSepartor.Get(context);

            if (this.InternalGroupSepartor == null || string.IsNullOrEmpty(this.InternalGroupSepartor))
            {
                this.InternalGroupSepartor = ",";
            }
        }

        protected override void Execute(CodeActivityContext context)
        {
            this.SetValues(context);

            if (InternalInputValue != 0)
            {
                Result.Set(context, NumberOuptPut());
            }
            else
            {
                Result.Set(context, InternalInputValue.ToString());
            }
        }

        private string NumberOuptPut()
        {
            CultureInfo cultureInfo = new CultureInfo("en-US");
            NumberFormatInfo numberFormat = cultureInfo.NumberFormat;

            numberFormat.NumberGroupSeparator = InternalGroupSepartor;

            List<string> strformat = InternalFormat.Split(',').ToList();
            List<Int32> NumberArr = strformat.Select(a => a.Length).ToList();
            NumberArr = NumberArr.AsEnumerable().Reverse().ToList();
            NumberArr.RemoveAt(NumberArr.Count - 1);
            NumberArr.Add(0);
            numberFormat.NumberGroupSizes = NumberArr.ToArray();

            return InternalInputValue.ToString("N0", numberFormat);
        }
    }
}
