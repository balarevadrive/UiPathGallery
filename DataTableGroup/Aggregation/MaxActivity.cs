namespace DataTableExtensions.Aggregation
{
    using System;
    using System.Activities;
    using System.ComponentModel;
    using System.Data;

    [DisplayName("Max")]
    public  class MaxActivity : CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        [Description("Enter the Input datatable name")]
        public InArgument<DataTable> InputTable { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [Description("Column name to Max")]
        public InArgument<string> ColumnName { get; set; }

        [Category("Input")]
        [Description("Where Condition Optional")]
        public InArgument<string> Select { get; set; }

        [Category("Input")]
        [Description("Convert InputColumn to [Optional]")]
        [DefaultValue(EnumMax.Select)]
        public EnumMax ConvertInputColumn { get; set; } = EnumMax.Select;

        [Category("Input")]
        [Description("Default value if it has null")]
        public InArgument<string> NullValue { get; set; }

        [Category("OutPut is Object")]
        public OutArgument<Object> Result { get; set; }

        private Type tempColumnType { get; set; }
        private DataTable localDataTable { get; set; }
        private string localColumnName { get; set; }
        private string localSelect { get; set; }
        private object localOutPut { get; set; }
        private string tempField { get; set; }
        private string localNullValue { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            try
            {
                this.localDataTable = InputTable.Get(context);
                this.localColumnName = ColumnName.Get(context);
                this.localSelect = Select.Get(context);
                this.localNullValue = NullValue.Get(context);

                this.DoMax();

                Result.Set(context, this.localOutPut);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void DoMax()
        {
            this.VerifyColumns();

            string strColumn;
            if (ConvertInputColumn != EnumMax.Select)
            {
                this.CopayData();

                strColumn = this.tempField;
            }
            else
            {
                strColumn = this.localColumnName;
            }

            if (string.IsNullOrEmpty(this.localSelect))
            {
                this.localSelect = string.Empty;
            }

            localOutPut = localDataTable.Compute("Max(" + strColumn + ")", this.localSelect);

            if (!string.IsNullOrEmpty(this.tempField))
            {
                localDataTable.Columns.Remove(this.tempField);
            }
        }

        private bool VerifyColumns()
        {
            bool blnResult = true;

            if (!string.IsNullOrEmpty(this.localColumnName))
            {

                if (!localDataTable.Columns.Contains(this.localColumnName))
                {
                    blnResult = false;

                    throw new Exception(this.localColumnName + " is invalid Column");
                }
            }

            return blnResult;
        }

        private void CopayData()
        {
            if (this.ConvertInputColumn == EnumMax.Integer)
            {
                this.tempColumnType = System.Type.GetType("System.Int64");
            }
            else if (this.ConvertInputColumn == EnumMax.Double)
            {
                this.tempColumnType = System.Type.GetType("System.Double");
            }
            else if (this.ConvertInputColumn == EnumMax.DateTime)
            {
                this.tempColumnType = System.Type.GetType("System.DateTime");
            }

            this.tempField = this.localColumnName + "temp";

            this.localDataTable.Columns.Add(this.tempField, this.tempColumnType);

            for (int intRow = 0; intRow < this.localDataTable.Rows.Count; intRow++)
            {
                try
                {
                    if (this.localDataTable.Rows[intRow][localColumnName] == DBNull.Value || this.localDataTable.Rows[intRow][localColumnName] == null
                        || string.IsNullOrEmpty(this.localDataTable.Rows[intRow][localColumnName].ToString()))
                    {

                        if (!string.IsNullOrEmpty(this.localNullValue))
                        {
                            this.localDataTable.Rows[intRow][tempField] = Convert.ChangeType(this.localNullValue, this.tempColumnType);
                        }
                        else
                        {
                            this.localDataTable.Rows[intRow][tempField] = DBNull.Value;
                        }
                    }
                    else
                    {
                        this.localDataTable.Rows[intRow][tempField] = Convert.ChangeType(this.localDataTable.Rows[intRow][localColumnName], this.tempColumnType);
                    }
                }
                catch (Exception ex)
                {
                    string msg = "Invalid data at row#" + (intRow + 1).ToString() + " Value :" + this.localDataTable.Rows[intRow][this.localColumnName].ToString();
                    throw new Exception(ex.Message + Environment.NewLine + msg);
                }
            }
        }
    }
}
