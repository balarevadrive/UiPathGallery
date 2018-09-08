namespace DataTableExtensions
{
    using System;
    using System.Activities;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Data;
    using System.Linq;

    [DisplayName("Remove Data Row Select")]
    public class RemoveDataRow : CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        [Description("Enter the Input datatable name")]
        public InArgument<DataTable> InputTable { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [Description("Where Condition")]
        public InArgument<string> Select { get; set; }

        [Category("OutPut")]
        public OutArgument<DataTable> Result { get; set; }

        private DataTable table { get; set; }
        private string strCondition { get; set; }
        private List<DataRow> filterRows { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            try
            {
                table = InputTable.Get(context);
                strCondition = Select.Get(context);

                if (table != null && table.Rows.Count > 0 && !string.IsNullOrEmpty(strCondition))
                {
                    filterRows = table.Select(strCondition).ToList();

                    if (filterRows != null && filterRows.Count > 0)
                    {
                        for (int intRow = 0; intRow < filterRows.Count; intRow++)
                        {
                            table.Rows.Remove(filterRows[intRow]);
                        }
                    }
                }

                Result.Set(context, table);
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }
    }
}
