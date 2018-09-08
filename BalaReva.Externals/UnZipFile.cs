namespace BalaReva.Externals
{
    using Design;
    using Ionic.Zip;
    using System;
    using System.Activities;
    using System.ComponentModel;
    using System.IO;

    [DisplayName("UnZip")]
    [Designer(typeof(UnZipFileDesign))]
    public class UnZipFile : CodeActivity
    {
        [Category("Input"), RequiredArgument]
        [Description("Enter the Zip File with path")]
        [DisplayName("Zip File")]
        public InArgument<string> ZipFile { get; set; }

        [Category("Input"), RequiredArgument]
        [Description("Enter the Extract Folder Path")]
        [DisplayName("Extract Folder Path")]
        public InArgument<string> ExtractFolderPath { get; set; }

        [Category("Input")]
        [Description("Encoding Code Page No")]
        [DisplayName("Code Page")]
        public InArgument<int> CodePage { get; set; } = 1252;

        [Category("Input")]
        [Description("Enter the Password")]
        public InArgument<string> Password { get; set; }

        private string filePath = string.Empty;
        private string folderPath = string.Empty;
        private string pwd = string.Empty;
        private int intCodePage;

        protected override void Execute(CodeActivityContext context)
        {
            try
            {
                filePath = ZipFile.Get(context);
                folderPath = ExtractFolderPath.Get(context);
                pwd = Password.Get(context);
                intCodePage = CodePage.Get(context);

                this.DoUnZip();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void DoUnZip()
        {
            this.Verify();

            ZipFile zip = new ZipFile();

            ReadOptions options = new ReadOptions
            {
                StatusMessageWriter = System.Console.Out,
                Encoding = System.Text.Encoding.GetEncoding(intCodePage)  // "932" for Shift-JIS Japanse
            };

            using (zip = Ionic.Zip.ZipFile.Read(filePath, options))
            {
                if (!string.IsNullOrEmpty(pwd))
                {
                    zip.Password = pwd;
                }

                if (!Directory.Exists(folderPath))
                {
                    Directory.CreateDirectory(folderPath);
                }

                foreach (ZipEntry e in zip)
                {
                    // check if you want to extract e or not
                    e.Extract(folderPath, ExtractExistingFileAction.OverwriteSilently);
                }
            }
        }

        private void Verify()
        {
            if (this.filePath == null || string.IsNullOrEmpty(this.filePath))
            {
                throw new Exception("Zip File path is empty");
            }

            if (this.folderPath == null || string.IsNullOrEmpty(this.folderPath))
            {
                throw new Exception("Folder path is empty");
            }
        }
    }
}
