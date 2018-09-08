namespace BalaReva.Externals.Others
{
    using Newtonsoft.Json.Linq;
    using System;
    using System.Activities;
    using System.ComponentModel;
    using System.Net;

    [DisplayName("Exchange Rate")]
    public class ExchangeRate : CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        [Description("From Currency")]
        [DisplayName("From Currency")]
        public InArgument<string> FromCurrency { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [Description("To Currency")]
        [DisplayName("To Currency")]
        public InArgument<string> ToCurrency { get; set; }

        [Category("Output")]
        [Description("Exchange Rage")]
        public OutArgument<decimal> Rate { get; set; }

        private string weblink = "http://free.currencyconverterapi.com/api/v5/convert?q={0}_{1}&compact=ultra";

        private string strFCur { get; set; }
        private string strTCur { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            strFCur = FromCurrency.Get(context);
            strTCur = ToCurrency.Get(context);
            Rate.Set(context, GetExchangeRate());
        }

        private decimal GetExchangeRate()
        {
            decimal result = 0;

            try
            {
                weblink = string.Format(weblink, strFCur, strTCur);

                using (WebClient web = new WebClient())
                {
                    string response = web.DownloadString(weblink);
                    JObject results = JObject.Parse(response);

                    var rate = results.First.Last;
                    result = Convert.ToDecimal(rate);
                }
            }
            catch (Exception)
            {
                result = 0;
            }

            return result;
        }
    }
}
