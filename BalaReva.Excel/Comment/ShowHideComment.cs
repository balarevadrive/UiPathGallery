namespace BalaReva.Excel.Comment
{
    using Design;
    using System;
    using System.Activities;
    using System.ComponentModel;

    [DisplayName("Show/Hide Comment")]
    [Designer(typeof(ExcelSelection))]
    public class ShowHideComment : BaseCommnet
    {
        [Category("Input")]
        [RequiredArgument]
        [Description("Show / Hide the comment")]
        [DisplayName("Show Comment"),DefaultValue(true)]
        public InArgument<bool> ShowComment { get; set; } = true;

        protected override void Execute(CodeActivityContext context)
        {
            try
            {
                base.InitValue(context);

                base.DoComment(Enums.CommentEnums.ShowHide, string.Empty, ShowComment.Get(context));
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
