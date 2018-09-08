namespace BalaReva.Excel.Comment
{
    using Design;
    using System;
    using System.Activities;
    using System.ComponentModel;

    [DisplayName("Get Comment")]
    [Designer(typeof(ExcelSelection))]
    public class GetComment : BaseCommnet
    {
        [Category("Output")]
        [RequiredArgument]
        [Description("Return the comment text")]
        [DisplayName("Result")]
        public OutArgument<string> Result { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            try
            {
                base.InitValue(context);

                base.DoComment(Enums.CommentEnums.GetComments);

                Result.Set(context, base.CommentValue);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
