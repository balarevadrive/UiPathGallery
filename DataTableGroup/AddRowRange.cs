namespace DataTableExtensions
{
    using System;
    using System.Activities;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Data;

    [DisplayName("Add Data Row Range")]
    public class AddRowRange : CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        [Description("Enter the Input datatable name")]
        public InArgument<DataTable> InputTable { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [Description("List of Datarow")]
        public InArgument<List<DataRow>> RowCollection { get; set; }

        [Category("OutPut")]
        public OutArgument<DataTable> Result { get; set; }

        private DataTable table { get; set; }
        private List<DataRow> localRowList { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            try
            {
                table = InputTable.Get(context);
                localRowList = RowCollection.Get(context);

                if (table != null && localRowList != null && localRowList.Count > 0)
                {
                    for (int introw = 0; introw < localRowList.Count; introw++)
                    {
                        table.Rows.Add(localRowList[introw].ItemArray);
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
