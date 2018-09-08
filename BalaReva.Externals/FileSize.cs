namespace BalaReva.Externals.Systems
{
    using System;
    using System.Activities;
    using System.ComponentModel;
    using System.IO;

    [DisplayName("File Size")]
    public class FileSize : CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        [Description("Local file full path")]
        public InArgument<string> FilePath { get; set; }

        [Category("OutPut")]
        [Description("Size in Bytes")]
        public OutArgument<long> SizeBytes { get; set; }

        [Category("OutPut")]
        [Description("Size in MB")]
        public OutArgument<double> SizeMB { get; set; }

        private string localPath { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            long resultBytes = 0;
            double resultMB = 0;

            try
            {
                localPath = FilePath.Get(context);

                this.Validate();

                resultBytes = FindFileSize(localPath);

                if (resultBytes > 0)
                {
                    resultMB = resultBytes / 1024;
                    resultMB = resultMB / 1024;
                    resultMB = System.Math.Round(resultMB, 2);
                }

                SizeBytes.Set(context, resultBytes);
                SizeMB.Set(context, resultMB);
            }

            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void Validate()
        {
            if (!File.Exists(localPath))
            {
                throw new Exception("Invalid File");
            }
        }

        private long FindFileSize(string filePath)
        {
            long result = new FileInfo(filePath).Length;

            return result;
        }
    }
}
