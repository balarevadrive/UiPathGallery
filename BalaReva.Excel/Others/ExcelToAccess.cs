namespace BalaReva.Excel.Others
{
    using Design;
    using System;
    using System.Activities;
    using System.ComponentModel;
    using System.Data;
    using System.Data.OleDb;
    using System.IO;
    using Access = Microsoft.Office.Interop.Access;

    [DisplayName("Excel To Access")]
    [Designer(typeof(ExcelSelection))]
    public class ExcelToAccess :BaseExcel
    {
        [Category("Input- Excel"), RequiredArgument]
        [Description("Excel File path"), DisplayName("Excel File Path")]
        public InArgument<string> FilePath { get; set; }

        [Category("Input- Excel"), RequiredArgument]
        [Description("Sheet Name"), DisplayName("Sheet Name")]
        public InArgument<string> SheetName { get; set; }

        [Category("Input- Excel")]
        [Description("Cell Range"), DisplayName("Cell Range")]
        public InArgument<string> CellRange { get; set; }


        [Category("Input- Access"), RequiredArgument]
        [Description("Access File path"), DisplayName("Access File Path")]
        public InArgument<string> AccessFilePath { get; set; }

        [Category("Input- Access"), RequiredArgument]
        [Description("Access TableName"), DisplayName("Access TableName")]
        public InArgument<string> AccessTableName { get; set; }

        private string strFilePath { get; set; }
        private string strSheetName { get; set; }
        private string strAccessFilePath { get; set; }
        private string strAccessTableName { get; set; }
        private string strCellRange { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            this.strFilePath = FilePath.Get(context);
            this.strSheetName = SheetName.Get(context);
            this.strAccessFilePath = AccessFilePath.Get(context);
            this.strAccessTableName = AccessTableName.Get(context);
            this.strCellRange = CellRange.Get(context);

            this.ExportToAccess();
        }


        private void ExportToAccess()
        {
            try
            {
                this.FileCreation();

                string _conn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + strFilePath + ";Extended Properties=Excel 12.0;";

                using (OleDbConnection _connection = new OleDbConnection(_conn))
                {
                    using (OleDbCommand _command = new OleDbCommand())
                    {
                        string strSql = string.Empty;

                        if (!string.IsNullOrEmpty(this.strCellRange))
                        {
                            strSql = @"SELECT * INTO [MS Access;Database=" + this.strAccessFilePath + "].[" + this.strAccessTableName + "] FROM [" + this.strSheetName + "$" + this.strCellRange + "]";
                        }
                        else
                        {
                            strSql = @"SELECT * INTO [MS Access;Database=" + this.strAccessFilePath + "].[" + this.strAccessTableName + "] FROM [" + this.strSheetName + "$]";
                        }

                        _command.Connection = _connection;
                        _command.CommandText = strSql;

                        _connection.Open();
                        _command.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void FileCreation()
        {
            if (!File.Exists(strAccessFilePath))
            {
                if (!new DirectoryInfo(strAccessFilePath).Exists)
                {
                    new DirectoryInfo(new FileInfo(strAccessFilePath).DirectoryName).Create();
                }

                Access.Application _accessApp = new Access.Application();

                _accessApp.Visible = false;

                _accessApp.NewCurrentDatabase(strAccessFilePath);

                _accessApp.CloseCurrentDatabase();

                _accessApp.Quit(Access.AcQuitOption.acQuitSaveAll);

                base.releaseObject(_accessApp);

                base.ClearnGarbage();
            }
            else
            {
                string _conn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + this.strAccessFilePath + "'";
                using (OleDbConnection _connection = new OleDbConnection(_conn))
                {
                    string[] restrictions = new string[4];
                    restrictions[2] = this.strAccessTableName;
                    _connection.Open();
                    DataTable dbTbl = _connection.GetSchema("Tables", restrictions);

                    if (dbTbl != null)
                    {
                        if (dbTbl.Rows.Count > 0)
                        {
                            dbTbl.Dispose();
                            dbTbl = null;
                            throw new Exception("table is already exists");
                        }
                        else
                        {
                            dbTbl.Dispose();
                            dbTbl = null;
                        }
                    }
                }
            }
        }
    }
}

