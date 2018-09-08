namespace DataTableExtensions.Aggregation
{
    using System;
    using System.Activities;
    using System.ComponentModel;
    using System.Data;

    [DisplayName("Count")]
    public class CountActivity : CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        [Description("Enter the Input datatable name")]
        public InArgument<DataTable> InputTable { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [Description("Column name to Count")]
        public InArgument<string> ColumnName { get; set; }

        [Category("Input")]
        [Description("Where Condition Optional")]
        public InArgument<string> Select { get; set; }


        [Category("OutPut is Object")]
        public OutArgument<Object> Result { get; set; }

        private DataTable localDataTable { get; set; }
        private string localColumnName { get; set; }
        private string localSelect { get; set; }
        private object localOutPut { get; set; }
        private string tempField { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            try
            {
                this.localDataTable = InputTable.Get(context);
                this.localColumnName = ColumnName.Get(context);
                this.localSelect = Select.Get(context);

                this.DoCount();

                Result.Set(context, this.localOutPut);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void DoCount()
        {
            try
            {
                this.VerifyColumns();

                this.tempField = this.localColumnName;

                this.CopayData();

                if (string.IsNullOrEmpty(this.localSelect))
                {
                    this.localSelect = string.Empty;
                }

                localOutPut = localDataTable.Compute("Count(" + this.tempField + ")", this.localSelect);

                if (!string.IsNullOrEmpty(this.tempField))
                {
                    localDataTable.Columns.Remove(this.tempField);
                }
            }
            catch (Exception ex)
            {
                throw ex;
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
            this.tempField = this.localColumnName + "temp";
            this.localDataTable.Columns.Add(this.tempField, System.Type.GetType("System.Int64"));

            for (int intRow = 0; intRow < this.localDataTable.Rows.Count; intRow++)
            {
                try
                {
                    if (string.IsNullOrEmpty(this.localDataTable.Rows[intRow][localColumnName].ToString()))
                    {
                        this.localDataTable.Rows[intRow][tempField] = DBNull.Value;
                    }
                    else
                    {
                        this.localDataTable.Rows[intRow][tempField] = 1;
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
