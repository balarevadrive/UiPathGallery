namespace BalaReva.Externals.Systems
{
    using System;
    using System.Activities;
    using System.ComponentModel;
    using System.IO;
    using System.Linq;

    [DisplayName("Directory Size")]
    public class DirectorySize : CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        [Description("Local directory path")]
        public InArgument<string> DirectoryPath { get; set; }

        [Category("Input")]
        [Description("Directory name to search")]
        public InArgument<string> FindDirectoryName { get; set; }

        [Category("Input")]
        [Description("Include the sub directory ")]
        [DefaultValue(true)]
        public InArgument<bool> IncludeSubDir { get; set; } = true;

        [Category("OutPut")]
        [Description("Size in Bytes")]
        public OutArgument<long> SizeBytes { get; set; }

        [Category("OutPut")]
        [Description("Size in MB")]
        public OutArgument<double> SizeMB { get; set; }

        [Category("OutPut")]
        [Description("Result of find directory's full path")]
        public OutArgument<string> OutPutPath { get; set; }

        private string localPath { get; set; }
        private bool localSubDir { get; set; }
        private string localSearch { get; set; }
        private string localOutPutPath { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            try
            {
                long resultBytes = 0;
                double resultMB = 0;
                localOutPutPath = string.Empty;

                localPath = DirectoryPath.Get(context);
                localSubDir = IncludeSubDir.Get(context);
                localSearch = FindDirectoryName.Get(context);

                this.Validate();

                resultBytes = FindDirectorySize(new DirectoryInfo(localPath), localSubDir);
                if (resultBytes > 0)
                {
                    resultMB = resultBytes / 1024;
                    resultMB = resultMB / 1024;
                    resultMB = System.Math.Round(resultMB, 2);
                }

                SizeBytes.Set(context, resultBytes);
                SizeMB.Set(context, resultMB);
                if (!string.IsNullOrEmpty(localOutPutPath))
                {
                    OutPutPath.Set(context, localOutPutPath);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private long FindDirectorySize(DirectoryInfo dInfo, bool includeSubDir)
        {
            if (!string.IsNullOrEmpty(localSearch)
                && string.IsNullOrEmpty(localOutPutPath)
                && localSearch.ToUpper() == dInfo.Name.ToUpper())
            {
                localOutPutPath = dInfo.FullName;
            }

            // Enumerate all the files
            long totalSize = dInfo.EnumerateFiles()
                         .Sum(file => file.Length);

            // If Subdirectories are to be included
            if (includeSubDir)
            {
                // Enumerate all sub-directories
                totalSize += dInfo.EnumerateDirectories()
                         .Sum(dir => FindDirectorySize(dir, true));
            }

            return totalSize;
        }

        private void Validate()
        {
            if (!Directory.Exists(localPath))
            {
                throw new Exception("Invalid Directory");
            }
        }
    }
}

