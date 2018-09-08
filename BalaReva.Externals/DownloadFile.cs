namespace BalaReva.Externals
{
    using System;
    using System.Activities;
    using System.ComponentModel;
    using System.Net;

    [DisplayName("Download File")]
    public class DownloadFile : CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        [Description("Enter the Network Address")]
        public InArgument<string> URL { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [Description("Local file path with file Name ")]
        public InArgument<string> FilePath { get; set; }

        private string internalNetworkAdd { get; set; }
        private string internalFilePath { get; set; }


        protected override void Execute(CodeActivityContext context)
        {
            try
            {
                internalNetworkAdd = URL.Get(context);
                internalFilePath = FilePath.Get(context);

                this.Download_File();

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void Download_File()
        {
            using (WebClient webClient = new WebClient())
            {
                webClient.DownloadFile(internalNetworkAdd, internalFilePath);
            }
        }
    }
}
