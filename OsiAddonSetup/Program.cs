using System;
using System.Threading;
using System.Diagnostics;
using System.Windows.Forms;
using System.Security.Principal;


namespace OsiAddonSetup
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(String[] args)
        {
            // Cria o handler para exceções não tratadas
            AppDomain.CurrentDomain.UnhandledException += new UnhandledExceptionEventHandler(NotifyUnhandledException);
            Application.ThreadException += new ThreadExceptionEventHandler(NotifyThreadException);
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);

            if (args.Length != 1)
            {
                MessageBox.Show("This installer must be run from Sap Business One", "OsiAddon setup", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            // Verifica se o instalador está sendo executado com permissões administrativas
            WindowsIdentity windowsIdentity = WindowsIdentity.GetCurrent();
            WindowsPrincipal windowsPrincipal = new WindowsPrincipal(windowsIdentity);
            Boolean executingAsAdmin = windowsPrincipal.IsInRole(WindowsBuiltInRole.Administrator);

            // Verifica se a caixa de dialogo do UAC (User Account Control) é necessária
            if (!executingAsAdmin)
            {
                // Pede elevação de privilégios (executa como administrador se o usuário concordar), o programa
                // atual é encerrado e uma nova instancia é executada com os privilégios concedidos
                ProcessStartInfo processInfo = new ProcessStartInfo();
                processInfo.Verb = "runas";
                processInfo.FileName = Application.ExecutablePath;
                processInfo.Arguments = '"' + Environment.GetCommandLineArgs()[1] + '"';
                try { Process.Start(processInfo); }
                catch { }
                return;
            }

            if (args[0] == "/U")
            {
                frmInstall.Uninstall();
                return;
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new frmInstall());
        }

        private static void NotifyUnhandledException(Object sender, UnhandledExceptionEventArgs e)
        {
            Exception unhandledException = (Exception)e.ExceptionObject;
            MessageBox.Show(unhandledException.Message);
        }

        private static void NotifyThreadException(Object sender, ThreadExceptionEventArgs e)
        {
            MessageBox.Show(e.Exception.Message);
        }
    }

}
