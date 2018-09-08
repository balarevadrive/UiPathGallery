namespace DataTableExtensions.XML
{
    using System;
    using System.Activities;
    using System.ComponentModel;
    using System.Data;
    using System.IO;
    using System.Xml;

    [DisplayName("DataTable to XML Document")]
    public class DataTableToXMLDocument : CodeActivity
    {
        private DataTable tempTable { get; set; }

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
        [Description("Enter the XmlDocument")]
        public OutArgument<XmlDocument> Document { get; set; }

        public DataTableToXMLDocument()
        {
            this.WriteMode = XmlWriteMode.IgnoreSchema;
        }

        protected override void Execute(CodeActivityContext context)
        {
            try
            {
                this.tempTable = InputTable.Get(context);

                XmlDocument xmldoc = this.GetXmlDocument();

                Document.Set (context, xmldoc);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public XmlDocument GetXmlDocument()
        {
            string output = string.Empty;
            XmlDocument result = new XmlDocument();

            if (tempTable != null)
            {
                if (string.IsNullOrEmpty(tempTable.TableName))
                {
                    this.tempTable.TableName = "Table1";
                }

                using (StringWriter stringWriter = new StringWriter())
                {
                    this.tempTable.WriteXml(stringWriter, WriteMode);
                    output = stringWriter.ToString();
                }

                result.LoadXml(output);
            }

            return result;
        }
    }
}
