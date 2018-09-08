namespace BalaReva.Excel
{
    using Enums;
    using Interfaces;
    using System;
    using System.Activities;
    using System.ComponentModel;
    using ExcelObj = Microsoft.Office.Interop.Excel;

    public abstract class BaseCommnet : BaseExcelNew, IBaseCommnet
    {
        [Category("Input")]
        [RequiredArgument]
        [Description("Comment cell")]
        [DisplayName("Cell")]
        public InArgument<string> Cell { get; set; }

        protected string CommentValue { get; set; }

        private string strCell { get; set; }

        protected void InitValue(CodeActivityContext context)
        {
            try
            {
                base.LoadVariables(context);

                this.strCell = Cell.Get(context);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        protected void DoComment(CommentEnums CommnetOption, string strComment = "", bool blnShow = false)
        {
            try
            {
                base.InitWorkSheet();

                ExcelObj.Range xlRange = base.GetSheetRange(strCell);

                if (xlRange != null)
                {
                    if (CommnetOption == CommentEnums.AddComment)
                    {
                        xlRange.AddComment(strComment);
                    }

                    else if (CommnetOption == CommentEnums.DeleteComment)
                    {
                        xlRange.Comment.Delete();
                    }
                    else if (CommnetOption == CommentEnums.ShowHide)
                    {
                        xlRange.Comment.Visible = blnShow;
                    }
                    else if (CommnetOption == CommentEnums.GetComments)
                    {
                        this.CommentValue = xlRange.Comment.Text();
                    }

                    base.SaveWorkBook(true);
                }
                else
                {
                    base.ClearObject();
                    throw new Exception("Invalid Range");
                }
            }
            catch (Exception ex)
            {
                base.ClearObject();
                throw ex;
            }
        }

    }
}
