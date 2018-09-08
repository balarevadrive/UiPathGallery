namespace BalaReva.Externals.WindowsService
{
    using System;
    using System.Activities;
    using System.ComponentModel;
    using System.ServiceProcess;

    [DisplayName("Service Execute")]
    public class ServiceExecute : CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        [Description("Enter Service Name")]
        public InArgument<string> ServiceName { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [DefaultValue(ServiceEnum.Start)]
        public ServiceEnum ExecuteType { get; set; } = ServiceEnum.Start;

        private string localService { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            try
            {
                localService = ServiceName.Get(context);
                this.ExecuteService();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void ExecuteService()
        {
            using (ServiceController service = new ServiceController(localService))
            {
                ServiceControllerStatus serviceStatus = service.Status;

                if (ExecuteType == ServiceEnum.Start)
                {
                    if (serviceStatus == ServiceControllerStatus.Running)
                    {
                        throw new Exception("Service already is running");
                    }
                    else
                    {
                        service.Start();
                    }

                }
                else if (ExecuteType == ServiceEnum.Stop)
                {
                    if (serviceStatus == ServiceControllerStatus.Stopped)
                    {
                        throw new Exception("Service already stopped");
                    }
                    else
                    {
                        service.Stop();
                    }
                }
                else if (ExecuteType == ServiceEnum.Pause)
                {
                    if (serviceStatus == ServiceControllerStatus.Paused)
                    {
                        throw new Exception("Service already paused");
                    }
                    else
                    {
                        service.Pause();
                    }
                }
                else if (ExecuteType == ServiceEnum.Continue)
                {
                    service.Continue();
                }
            }
        }
    }
}
