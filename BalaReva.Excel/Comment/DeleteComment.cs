namespace BalaReva.Excel.Comment
{
    using Design;
    using System;
    using System.Activities;
    using System.ComponentModel;

    [DisplayName("Delete Comment")]
    [Designer(typeof(ExcelSelection))]
    public class DeleteComment : BaseCommnet
    {
        protected override void Execute(CodeActivityContext context)
        {
            try
            {
                base.InitValue(context);

                base.DoComment(Enums.CommentEnums.DeleteComment);

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
