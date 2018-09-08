namespace DataTableExtensions.XML
{
    using System;
    using System.Activities;
    using System.ComponentModel;
    using System.Data;
    using System.IO;

    [DisplayName("DataTable to XML File")]
    public class DataTableToXMLFile : CodeActivity
    {
        private DataTable tempTable { get; set; }
        private string tempPath { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [Description("Enter the Input datatable name")]
        public InArgument<DataTable> InputTable { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [Description("Select XmlWriteMode")]
        public XmlWriteMode WriteMode { get; set; }

        [Category("OutPut")]
        [RequiredArgument]
        [Description("Enter the Input datatable name")]
        public InArgument<string> FilePath { get; set; }

        public DataTableToXMLFile()
        {
            this.WriteMode = XmlWriteMode.IgnoreSchema;
        }

        protected override void Execute(CodeActivityContext context)
        {
            try
            {
                tempTable = InputTable.Get(context);
                tempPath = FilePath.Get(context);

                string[] words = tempPath.Split('/');

                if (words.Length <= 1)
                {
                    tempPath = Path.Combine(Environment.CurrentDirectory, tempPath);
                }

                if (string.IsNullOrEmpty(tempTable.TableName))
                {
                    tempTable.TableName = "Table1";
                }

                tempTable.WriteXml(tempPath, WriteMode);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
