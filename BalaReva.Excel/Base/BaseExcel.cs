namespace BalaReva.Excel
{
    using System;
    using System.Activities;

    public abstract class BaseExcel : CodeActivity
    {
        //release com objects to fully kill excel process from running in the background 
        protected void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception)
            {
                obj = null;
            }
        }

        //cleanup the Garbage.
        protected void ClearnGarbage()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}
