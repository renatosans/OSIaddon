//
// 1) Addon installer program should be able to accept a command line parameter from SBO.
//    This parameter is a string built from 2 strings devided by "|".
//    The first string is the path recommended by SBO for installation folder.
//    The second string is the location of "AddOnInstallAPI.dll".
//    For example, a command line parameter that looks like this:
//    "C:\MyAddon|C:\Program Files\SAP Manage\SAP Business One\AddOnInstallAPI.dll"
//    Means that the recommended installation folder for this addon is "C:\MyAddon" and the
//    location of "AddOnInstallAPI.dll" is "C:\Program Files\SAP Manage\SAP Business One\AddOnInstallAPI.dll"
//
// 2) When the installation is complete the installer must call the function 
//    "EndInstall" from "AddOnInstallAPI.dll" to inform SBO the installation is complete.
//    This dll contains 3 functions that can be used during the installation.
//    The functions are: 
//         1) EndInstall - Signals SBO that the installation is complete.
//         2) SetAddOnFolder - Use it if you want to change the installation folder.
//         3) RestartNeeded - Use it if your installation requires a restart, it will cause
//            the SBO application to close itself after the installation is complete.
//    All 3 functions return a 32 bit integer. There are 2 possible values for this integer.
//    0 - Success, 1 - Failure.
//
// 3) After your installer is ready you need to create an add-on registration file.
//    In order to create it you have a utility - "Add-On Registration Data Creator"
//    "..\SAP Manage\SAP Business One SDK\Tools\AddOnRegDataGen\AddOnRegDataGen.exe"
//    This utility creates a file with the extention 'ard', you will be asked to
//    point to this file when you register your addon.

using System;
using System.IO;
using System.Reflection;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Microsoft.Win32;


namespace OsiAddonSetup
{
    public partial class frmInstall : Form
    {
        // EndInstall - Signals SBO that the installation is complete.
        [DllImport("AddOnInstallAPI.dll")]
        public static extern Int32 EndInstall();

        // EndUninstall - Signals SBO that the addon removal is is complete.
        [DllImport("AddOnInstallAPI.dll")]
        public static extern Int32 EndUninstall(String path, Boolean succeed);


        private String destinationFolder;

        private String addonInstallDllFolder;

        private Boolean fileCreated;


        public frmInstall()
        {
            InitializeComponent();
            this.Icon = new Icon(GetEmbeddedResource("Setup.ico"));
            this.imgLogo.Image = new Bitmap(GetEmbeddedResource("Logo.png"));
            FileWatcher.EnableRaisingEvents = false;
        }

        private void frmInstall_Shown(Object sender, EventArgs e)
        {
            String commandLine = Environment.GetCommandLineArgs()[1];
            String[] commandLineElements = commandLine.Split(char.Parse("|"));
            if (commandLineElements.Length != 2)
            {
                btnInstall.Enabled = false;
                return;
            }
            destinationFolder = commandLineElements[0];
            addonInstallDllFolder = Path.GetDirectoryName(commandLineElements[1]);
            String info = "Setup Commandline:" + Environment.NewLine + "     " + Environment.CommandLine + Environment.NewLine +
                          "Destination Folder:" + Environment.NewLine + "     " +  destinationFolder  + Environment.NewLine +
                          "AddOnInstallAPI.dll Location:" + Environment.NewLine + "     " + addonInstallDllFolder;
            EventLog.WriteEntry("Ged Addon Setup", info);
        }

        //  This function extracts the given file into the folder specified
        private void ExtractFile(String destinationFolder, String filename)
        {
            try
            {
                Assembly thisExe = Assembly.GetExecutingAssembly();
                String finalFilename = destinationFolder + @"\" + filename;
                String tempFilename = destinationFolder + @"\" + Path.ChangeExtension(filename, ".tmp");

                Stream file = thisExe.GetManifestResourceStream(filename);

                //  Create a tmp file first, after file is extracted change to exe
                if (File.Exists(tempFilename)) File.Delete(tempFilename);

                Byte[] buffer = new Byte[file.Length];
                file.Read(buffer, 0, (int)file.Length);

                FileStream addonTempFile = File.Create(tempFilename);
                addonTempFile.Write(buffer, 0, (int)file.Length);
                addonTempFile.Close();

                if (File.Exists(finalFilename)) File.Delete(finalFilename);

                //  Change file extension to exe
                File.Move(tempFilename, finalFilename);
            }
            catch (Exception exc)
            {
                ShowError("Falha extraindo arquivos do addon: " + exc.Message);
            }
        }

        /// <summary>
        /// Busca um recurso embarcado na aplicação caso esteja disponível, alternativamente busca em disco
        /// Recebe o nome do recurso embarcado (ignora namespaces no nome)
        /// </summary>
        public static Stream GetEmbeddedResource(String name)
        {
            if (String.IsNullOrEmpty(name)) return null;

            String[] nameParts = name.Split(new Char[] { '.' });
            int length = nameParts.Length;
            if (length < 2) return null;
            String rawName = nameParts[length - 2] + "." + nameParts[length - 1];
            String qualifiedName = rawName;

            Assembly runningExe = Assembly.GetEntryAssembly();
            if (runningExe == null)
            {
                if (!File.Exists(rawName)) return null;
                return new FileStream(rawName, FileMode.Open);
            }

            Stream resourceStream = runningExe.GetManifestResourceStream(rawName);
            if (resourceStream == null)
            {
                qualifiedName = runningExe.GetName().Name + "." + rawName;
                resourceStream = runningExe.GetManifestResourceStream(qualifiedName);
            }
            if (resourceStream == null) return null;

            return resourceStream;
        }

        public Boolean AddToWindowsRegistry(String installationFolder, String addonInstallDllFolder)
        {
            RegistryKey parentKey = Registry.LocalMachine.OpenSubKey("SOFTWARE", true);
            if (parentKey == null)
            {
                ShowError("Erro ao registrar addon. Não foi encontrada a chave SOFTWARE.");
                return false;
            }

            RegistryKey addonKey = parentKey.OpenSubKey("Osi Addon", true);
            try
            {
                if (addonKey == null) addonKey = parentKey.CreateSubKey("Osi Addon");
                addonKey.SetValue("InstallationFolder", installationFolder);
                addonKey.SetValue("AddonInstallDllFolder", addonInstallDllFolder);
            }
            catch (Exception exc)
            {
                ShowError("Erro ao registrar addon." + exc.Message);
                return false;
            }

            addonKey.Close();
            parentKey.Close();

            return true;
        }

        private void btnInstall_Click(Object sender, EventArgs e)
        {
            this.btnInstall.Enabled = false;

            //  Create installation folder
            if (!Directory.Exists(destinationFolder)) Directory.CreateDirectory(destinationFolder);

            // Create sub folders
            String xmlFolder = Path.Combine(destinationFolder, "Xml");
            if (!Directory.Exists(xmlFolder)) Directory.CreateDirectory(xmlFolder);

            // Extrai todos os arquivos para o diretório
            ExtractAllFiles();

            //  Inform SBO the installation ended
            AddonInstallFinished(addonInstallDllFolder);

            // Write installation details to registry
            AddToWindowsRegistry(destinationFolder, addonInstallDllFolder);

            // Encerra o instalador
            this.Close();
        }

        private void ExtractAllFiles()
        {
            FileWatcher.Path = destinationFolder;
            FileWatcher.Renamed += new RenamedEventHandler(FileRenamed);
            FileWatcher.EnableRaisingEvents = true;

            // Begin EXE copy
            fileCreated = false;
            ExtractFile(destinationFolder, "OsiAddon.exe");
            // Don't continue running until the file is copied...
            while (fileCreated == false) Application.DoEvents();

            // Begin XML copy
            String xmlFolder = Path.Combine(destinationFolder, "Xml");
            fileCreated = false;
            ExtractFile(xmlFolder, "DataAccess.xml");
            // Don't continue running until the file is copied...
            while (fileCreated == false) Application.DoEvents();


            FileWatcher.EnableRaisingEvents = false;
        }

        //  This event happens when the addon file is renamed to exe extention
        private void FileRenamed(Object sender, RenamedEventArgs e)
        {
            fileCreated = true;
        }

        public static Boolean Uninstall()
        {
            RegistryKey parentKey = Registry.LocalMachine.OpenSubKey("SOFTWARE", true);
            if (parentKey == null) return false;

            RegistryKey addonKey = parentKey.OpenSubKey("Osi Addon", true);
            if (addonKey == null) return false;

            // Obtem informações de instalação
            String installationFolder = (String)addonKey.GetValue("InstallationFolder");
            String addonInstallDllFolder = (String)addonKey.GetValue("AddonInstallDllFolder");
            addonKey.Close();
            if (String.IsNullOrEmpty(installationFolder)) return false;
            if (String.IsNullOrEmpty(addonInstallDllFolder)) return false;

            try
            {
                // Remove o diretório de instalação
                if (Directory.Exists(installationFolder))
                    Directory.Delete(installationFolder, true);
                // Remove o registro do Addon
                parentKey.DeleteSubKey("Osi Addon");
                parentKey.Close();
            }
            catch (Exception exc)
            {
                ShowError("Falha na remoção: " + exc.Message);
                return false;
            }

            // Avisa o SAP sobre o termino da operação, trata exceções
            AddonUninstallFinished(addonInstallDllFolder);

            // Avisa o usuário sobre o termino da operação
            ShowInfo("Osi Addon removido com sucesso");
            return true;
        }

        public static void AddonInstallFinished(String addonInstallDllFolder)
        {
            // Muda o diretório corrente para o lugar onde está a DLL de instalação de Addons do SAP para
            // que seja possível utilizar as funções EndInstall e EndUninstall
            Environment.CurrentDirectory = addonInstallDllFolder;
            try
            {
                // Informa o SAP Business One que a instalação foi concluída
                EndInstall();
            }
            catch (Exception exc)
            {
                ShowError("Falha chamando EndInstall(). " + exc.Message);
            }
        }

        public static void AddonUninstallFinished(String addonInstallDllFolder)
        {
            // Muda o diretório corrente para o lugar onde está a DLL de instalação de Addons do SAP para
            // que seja possível utilizar as funções EndInstall e EndUninstall
            Environment.CurrentDirectory = addonInstallDllFolder;
            try
            {
                // Informa o SAP Business One que a remoção do addon foi concluída
                EndUninstall(null, true);
            }
            catch (Exception exc)
            {
                ShowError("Falha chamando EndUninstall(). " + exc.Message);
            }
        }

        private static void ShowError(String errorMessage)
        {
            MessageBox.Show(errorMessage, "Falha na instalação do Addon", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private static void ShowInfo(String message)
        {
            MessageBox.Show(message, "Informação", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

}
