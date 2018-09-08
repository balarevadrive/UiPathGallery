namespace BalaReva.Externals.WindowsService
{
    using System;
    using System.Activities;
    using System.ComponentModel;
    using System.ServiceProcess;

    [DisplayName("Service Status")]
    public class ServiceStatus : CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        [Description("Enter Service Name")]
        public InArgument<string> ServiceName { get; set; }


        [Category("OutPut")]
        [Description("Service Status")]
        public OutArgument<string> Status { get; set; }

        private string localService { get; set; }
        private string localStatus { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            try
            {
                localService = ServiceName.Get(context);

                this.GetServiceStatus();

                if (!string.IsNullOrEmpty(localStatus))
                {
                    Status.Set(context, localStatus);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void GetServiceStatus()
        {
            using (ServiceController controller = new ServiceController(localService))
            {
                ServiceControllerStatus serviceStatus = controller.Status;

                localStatus= serviceStatus.ToString();
            }
        }
    }
}
