namespace ExtensionRegistrar
{
    using System;
    using System.Diagnostics;
    using System.Windows.Forms;

    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string libraryToRegister = args[0];
                string action = args[1];

                if (action == "/register")
                {
                    Process.Start("regasm", string.Format("\"{0}\" /codebase", libraryToRegister));
                }
                else
                {
                    Process.Start("regasm", string.Format("\"{0}\" /unregister", libraryToRegister));
                    
                }
            }
            catch(Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }
    }
}
