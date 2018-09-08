namespace BalaReva.Externals
{
    using System;
    using System.Activities;
    using System.ComponentModel;
    using System.IO;

    [DisplayName("Zip")]
    public class ZipFiles : CodeActivity
    {
        [Category("Input"),RequiredArgument]
        [Description("Enter the Zip File with Path")]
        public InArgument<string> ZipFile { get; set; }

        [Category("Input"),RequiredArgument]
        [Description("Select folder option")]
        [DefaultValue(ZipEnum.Single)]
        public ZipEnum FolderOption { get; set; } = ZipEnum.Single;

        [Category("Input")]
        [Description("Enter the Zip single Folder")]
        [DisplayName("SingleFolder")]
        public InArgument<string> SourceFolder { get; set; }

        [Category("Input")]
        [Description("Enter the Zip multiple Folders")]
        public InArgument<string[]> MultipleFolders { get; set; }

        [Category("Input")]
        [Description("Enter the Password")]
        public InArgument<string> Password { get; set; }

        private string filePath = string.Empty;
        private string folderPath = string.Empty;
        private string[] folderArray = null;
        private string pwd = string.Empty;

        protected override void Execute(CodeActivityContext context)
        {
            try
            {
                filePath = ZipFile.Get(context);
                folderPath = SourceFolder.Get(context);
                pwd = Password.Get(context);
                folderArray = MultipleFolders.Get(context);

                this.DoZip();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void DoZip()
        {
            try
            {
                this.Verify();

                if (FolderOption == ZipEnum.Single)
                {
                    folderArray = new string[] { folderPath };
                }

                using (Ionic.Zip.ZipFile zip = new Ionic.Zip.ZipFile())
                {
                    if (!string.IsNullOrEmpty(pwd))
                    {
                        zip.Password = pwd;
                    }

                    foreach (string item in folderArray)
                    {
                        DirectoryInfo directoryInfo = new DirectoryInfo(item);

                        zip.AddDirectory(item, directoryInfo.Name);
                    }

                    zip.Save(filePath);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void Verify()
        {
            if (FolderOption == ZipEnum.Single && string.IsNullOrEmpty(this.folderPath))
            {
                throw new Exception("Folder path is empty");
            }

            else if (FolderOption == ZipEnum.Multiple && MultipleFolders == null)
            {
                throw new Exception("Multi Folder is empty");
            }
        }
    }
}
