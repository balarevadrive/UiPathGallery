namespace BalaReva.Access
{
    using Design;
    using System;
    using System.Activities;
    using System.ComponentModel;

    [DisplayName("Execute Macro")]
    [Designer(typeof(FileSelection))]
    public class ExecuteMacro : BaseAccess
    {
        [Category("Input")]
        [RequiredArgument]
        [Description("Enter the Macro Name")]
        public InArgument<string> MacroName { get; set; }

        private string internalMacro;

        protected override void Execute(CodeActivityContext context)
        {
            base.LoadVariables(context);

            internalMacro = MacroName.Get(context);

            this.Execute_Macro();
        }

        private void Execute_Macro()
        {
            try
            {
                base.InitApp();

                if (base.AccApp != null)
                {
                    base.AccApp.Visible = true;
                    base.AccApp.DoCmd.RunMacro((internalMacro as Object), (1 as Object), (true as Object));
                    base.AccApp.DoCmd.Quit(Microsoft.Office.Interop.Access.AcQuitOption.acQuitSaveNone);
                    base.Save();
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
