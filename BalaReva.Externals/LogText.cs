namespace BalaReva.Externals.Others
{
    using System.Activities;
    using System.ComponentModel;
    using System.IO;
    using System.Reflection;

    [DisplayName("Log Text")]
    public class LogText : CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        [Description("File Name")]
        [DisplayName("Log File Name")]
        public InArgument<string> LogFileName { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [Description("Message")]
        public InArgument<string> Message { get; set; }

        private string strFilePath { get; set; }
        private string strMessage { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            strFilePath = LogFileName.Get(context);
            strMessage = Message.Get(context);

            string strFullPath = Path.Combine(Directory.GetCurrentDirectory(), strFilePath);

            using (StreamWriter file = new StreamWriter(strFullPath, true))
            {
                file.WriteLine(strMessage + "    :" + System.DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt"));
            }
        }
    }
}
