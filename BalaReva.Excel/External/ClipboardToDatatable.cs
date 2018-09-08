namespace BalaReva.Excel.External
{
    using System;
    using System.Activities;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Data;
    using System.Linq;
    using System.Windows.Forms;

    [DisplayName("Clipboard To Datatable")]
    public class ClipboardToDatatable : BaseExcel
    {
        [Category("Input")]
        [RequiredArgument]
        [Description("Want to add Header")]
        [DisplayName("Add Header")]
        public InArgument<bool> HasHeader { get; set; } = false;

        [Category("Output")]
        [RequiredArgument]
        [Description("Data table")]
        [DisplayName("Datatable")]
        public OutArgument<DataTable> Datatable { get; set; }

        private bool IsHeader { get; set; }
        private DataTable dtClipboard;

        protected override void Execute(CodeActivityContext context)
        {
            try
            {
                this.IsHeader = HasHeader.Get(context);
                this.ClipboarToDatatable();

                Datatable.Set(context, dtClipboard);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void ClipboarToDatatable()
        {
            dtClipboard = new DataTable();
            char[] rowSeparator = { '\r', '\n' };
            char[] colSparator = { '\t' };

            try
            {
                IDataObject dataInClipboard = Clipboard.GetDataObject();

                string ClipboardText = (string)dataInClipboard.GetData(DataFormats.Text);

                if (!string.IsNullOrEmpty(ClipboardText))
                {
                    List<string> DataRows = ClipboardText.Split(rowSeparator).ToList();
                    List<string> ExcelRows = new List<string>();

                    foreach (string ItemRow in DataRows)
                    {
                        if (!string.IsNullOrEmpty(ItemRow))
                        {
                            ExcelRows.Add(ItemRow);
                        }
                    }

                    foreach (string ItemRow in ExcelRows)
                    {
                        List<string> ColData = ItemRow.Split(colSparator).ToList();
                        this.AddDataDatatable(ColData);
                    }
                }
                else
                {
                    MessageBox.Show("Invalid data format");
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Invalid data format");
            }
        }


        private void AddDataDatatable(List<string> lstData)
        {
            if (dtClipboard.Columns.Count == 0)
            {
                if (this.IsHeader)
                {
                    foreach (string colItem in lstData)
                    {
                        dtClipboard.Columns.Add(colItem);
                    }
                }
                else
                {
                    for (int intCol = 0; intCol <= lstData.Count - 1; intCol++)
                    {
                        dtClipboard.Columns.Add("Column" + (intCol + 1));
                    }

                    dtClipboard.Rows.Add(lstData.ToArray());
                }
            }
            else
            {
                dtClipboard.Rows.Add(lstData.ToArray());
            }
        }
    }
}
