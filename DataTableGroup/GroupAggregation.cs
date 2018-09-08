namespace DataTableExtensions
{
    using System;
    using System.Activities;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Data;
    using System.Linq;

    [DisplayName("Group By Aggregation")]
    public class GroupAggregation : CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        [Description("Enter the Input datatable name")]
        [DisplayName("Input Table")]
        public InArgument<DataTable> InputTable { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [Description("Column names are separated by comma")]
        [DisplayName("Group by Columns")]
        public InArgument<string> GroupbyColumns { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [Description("Should be one columnName")]
        [DisplayName("Aggregate Column")]
        public InArgument<string> AggregateColumn { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [DefaultValue(EnumType.Integer)]
        [DisplayName("Aggregate Type")]
        public EnumType AggregateType { get; set; } = EnumType.Integer;

        [Category("Input")]
        [RequiredArgument]
        [DefaultValue(EnumAggregation.Sum)]
        [DisplayName("Aggregate By")]
        public EnumAggregation AggregateBy { get; set; } = EnumAggregation.Sum;

        [Category("Input")]
        [Description("Default value if it has null")]
        [DisplayName("Null Value")]
        public InArgument<string> NullValue { get; set; }

        [Category("OutPut")]
        public OutArgument<DataTable> Result { get; set; }

        private DataTable localSourceTable { get; set; }
        private string localGroupbyColumn { get; set; }
        private string localAggregateColumn { get; set; }
        private string localAggregateBy { get; set; }
        private DataTable localResult { get; set; }
        private string tempField { get; set; }
        public string localNullValue { get; set; }

        private Type localAggregateColumnType { get; set; }

        private void SetValue(CodeActivityContext context)
        {
            this.localSourceTable = InputTable.Get(context);

            this.localNullValue = NullValue.Get(context);

            if (this.localSourceTable == null)
            {
                throw new Exception("Input table is null");
            }

            this.localGroupbyColumn = GroupbyColumns.Get(context);

            this.localAggregateColumn = AggregateColumn.Get(context);

            this.localAggregateBy = AggregateBy.ToString();

            if (AggregateType.ToString() == "Integer")
            {
                this.localAggregateColumnType = System.Type.GetType("System.Int64");
            }
            else
            {
                this.localAggregateColumnType = System.Type.GetType("System." + AggregateType.ToString());
            }
        }

        protected override void Execute(CodeActivityContext context)
        {
            try
            {
                this.SetValue(context);

                Result.Set(context, this.GroupBy());
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private DataTable GroupBy()
        {
            DataTable dtGroup = new DataTable();

            string strCondition = string.Empty;
            string strCol = string.Empty;

            if (this.VerifyColumns())
            {
                this.CopayData();

                strCol = this.localAggregateColumn;

                if (!string.IsNullOrEmpty(tempField))
                {
                    this.localAggregateColumn = this.tempField;
                }

                DataView dv = new DataView(localSourceTable);

                var FilterArr = localGroupbyColumn.Split(new char[] { ',' }, StringSplitOptions.None);

                //getting distinct values for group column
                dtGroup = dv.ToTable(true, FilterArr);

                //adding column for the row count
                dtGroup.Columns.Add(strCol, localAggregateColumnType);

                //looping thru distinct values for the group, counting
                foreach (DataRow dr in dtGroup.Rows)
                {
                    strCondition = GetWhereCondition(FilterArr, dr);

                    dr[strCol] = localSourceTable.Compute(localAggregateBy + "(" + localAggregateColumn + ")", strCondition);
                }

                if (!string.IsNullOrEmpty(tempField) && localSourceTable.Columns.Contains(tempField))
                {
                    localSourceTable.Columns.Remove(tempField);
                }
            }

            return dtGroup;
        }

        private string GetWhereCondition(string[] columns, DataRow dataRow)
        {
            string Result = string.Empty;
            List<string> lstColumns = columns.ToList();
            List<string> lstCondition = new List<string>();

            foreach (string col in lstColumns)
            {
                if (localSourceTable.Columns[col].DataType == typeof(string))
                {
                    lstCondition.Add("[" + col + "]= '" + dataRow[col] + "'");
                }
                else if (localSourceTable.Columns[col].DataType == typeof(object))
                {
                    lstCondition.Add("[" + col + "]= '" + dataRow[col].ToString() + "'");
                }
                else if (localSourceTable.Columns[col].DataType == typeof(Int16)
                    || localSourceTable.Columns[col].DataType == typeof(Int32)
                    || localSourceTable.Columns[col].DataType == typeof(Int64)
                    || localSourceTable.Columns[col].DataType == typeof(double)
                    || localSourceTable.Columns[col].DataType == typeof(decimal)
                    || localSourceTable.Columns[col].DataType == typeof(bool))
                {
                    lstCondition.Add("[" + col + "]= " + dataRow[col] + "");
                }
                else if (localSourceTable.Columns[col].DataType == typeof(DateTime))
                {
                    if (dataRow[col].ToString().Length > 18)
                    {
                        if (dataRow[col].ToString().Substring(10).Trim() == "00:00:00")
                        {
                            lstCondition.Add("[" + col + "]= #" + Convert.ToDateTime(dataRow[col]).Date.ToString("MM/dd/yyy") + "#");
                        }
                        else
                        {
                            lstCondition.Add("[" + col + "]= #" + Convert.ToDateTime(dataRow[col]).Date.ToString("MM/dd/yyy hh:mm:ss tt") + "#");
                        }
                    }
                    else
                    {
                        lstCondition.Add("[" + col + "]= #" + Convert.ToDateTime(dataRow[col]).Date.ToString("MM/dd/yyy") + "#");
                    }
                }
            }

            if (lstCondition.Count == 1)
            {
                Result = lstCondition.FirstOrDefault();
            }
            else
            {
                Result = string.Join(" And ", lstCondition.ToArray());
            }

            return Result;
        }

        private bool VerifyColumns()
        {
            bool blnResult = true;
            if ((this.localAggregateColumnType == typeof(DateTime)) 
                && (this.localAggregateBy == EnumAggregation.Avg.ToString()
                        || this.localAggregateBy == EnumAggregation.Count.ToString()
                        || this.localAggregateBy == EnumAggregation.Sum.ToString()))
            {
                blnResult = false;
                throw new Exception(this.localAggregateBy.ToUpper() + " function will not support the " + this.localAggregateColumnType.ToString());
            }


            if (!string.IsNullOrEmpty(localGroupbyColumn))
            {
                List<string> LstColumn = localGroupbyColumn.Split(new char[] { ',' }, StringSplitOptions.None).ToList();
                LstColumn.Add(localAggregateColumn);

                if (LstColumn.Count() > 0)
                {
                    foreach (string str in LstColumn)
                    {
                        if (!localSourceTable.Columns.Contains(str))
                        {
                            blnResult = false;
                            throw new Exception(str + " is invalid Column");
                        }
                    }
                }
            }

            return blnResult;
        }

        private void CopayData()
        {
            if (this.localSourceTable.Columns[localAggregateColumn].DataType == this.localAggregateColumnType)
            {
                this.tempField = string.Empty;
            }
            else
            {
                this.tempField = localAggregateColumn + "temp";
                this.localSourceTable.Columns.Add(tempField, this.localAggregateColumnType);

                for (int intRow = 0; intRow < localSourceTable.Rows.Count; intRow++)
                {
                    try
                    {
                        if (this.localSourceTable.Rows[intRow][localAggregateColumn] == DBNull.Value || this.localSourceTable.Rows[intRow][localAggregateColumn] == null
                        || string.IsNullOrEmpty(this.localSourceTable.Rows[intRow][localAggregateColumn].ToString()))
                        {
                            if (!string.IsNullOrEmpty(this.localNullValue))
                            {
                                this.localSourceTable.Rows[intRow][tempField] = Convert.ChangeType(this.localNullValue, this.localAggregateColumnType);
                            }
                            else
                            {
                                this.localSourceTable.Rows[intRow][tempField] = DBNull.Value;
                            }

                        }
                        else
                        { 
                            this.localSourceTable.Rows[intRow][tempField] = Convert.ChangeType(this.localSourceTable.Rows[intRow][localAggregateColumn], this.localAggregateColumnType);
                        }
                    }
                    catch (Exception ex)
                    {
                        string msg = "Invalid data at row#" + (intRow + 1).ToString() + " Value :" + this.localSourceTable.Rows[intRow][localAggregateColumn].ToString();
                        throw new Exception(ex.Message + Environment.NewLine + msg);
                    }
                }
            }
        }
    }
}
