using System;
using System.Runtime.InteropServices;
using SolidEdgeFramework;


namespace Blueprint
{
    //background tools for Solid Edge
    //keeps the ugly code hidden
    public class SETools
    {
        private Application connect_to_SE()
        {
            Application app = null;
            try
            {
                OleMessageFilter.Register();
                app = (Application)Marshal.GetActiveObject("SolidEdge.Application");  //connect to a running instance of Solid Edge
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
            finally { OleMessageFilter.Revoke(); }
            return app;
        }


        private Application start_SE()
        {
            Type type = null;
            Application app = null;
            try
            {
                OleMessageFilter.Register();
                type = Type.GetTypeFromProgID("SolidEdge.Application");
                app = (Application)Activator.CreateInstance(type);  //starts Solid Edge
                app.Visible = true;
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
            finally { OleMessageFilter.Revoke(); }
            return app;
        }


        //connect to Solid Edge, start SE if not yet running
        public Application startup_SE()
        {
            Application app = null;
            app = connect_to_SE();
            if (app == null) app = start_SE();
            return app;
        }
    }
}
