namespace BalaReva.Excel.Comment
{
    using Design;
    using System;
    using System.Activities;
    using System.ComponentModel;

    [DisplayName("Add Comment")]
    [Designer(typeof(ExcelSelection))]
    public class AddComment : BaseCommnet
    {
        [Category("Input")]
        [RequiredArgument]
        [Description("Comment String")]
        [DisplayName("Comment")]
        public InArgument<string> Comment { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            try
            {
                base.InitValue(context);
                base.DoComment(Enums.CommentEnums.AddComment, Comment.Get(context));
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
