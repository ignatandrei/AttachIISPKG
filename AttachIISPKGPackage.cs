using System;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.ComponentModel.Design;
using Microsoft.Win32;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using EnvDTE80;
using EnvDTE;
using System.IO;

namespace Company.AttachIISPKG
{
    enum AttachResult
    {
        None = 0,
        Ok ,
        NotFound,
        FailedElevation,
        Failed

    }
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    ///
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the 
    /// IVsPackage interface and uses the registration attributes defined in the framework to 
    /// register itself and its components with the shell.
    /// </summary>
    // This attribute tells the PkgDef creation utility (CreatePkgDef.exe) that this class is
    // a package.
    [PackageRegistration(UseManagedResourcesOnly = true)]
    // This attribute is used to register the information needed to show this package
    // in the Help/About dialog of Visual Studio.
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    // This attribute is needed to let the shell know that this package exposes some menus.
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(GuidList.guidAttachIISPKGPkgString)]
    public sealed class AttachIISPKGPackage : Package
    {
        /// <summary>
        /// Default constructor of the package.
        /// Inside this method you can place any initialization code that does not require 
        /// any Visual Studio service because at this point the package object is created but 
        /// not sited yet inside Visual Studio environment. The place to do all the other 
        /// initialization is the Initialize method.
        /// </summary>
        public AttachIISPKGPackage()
        {
            Debug.WriteLine(string.Format(CultureInfo.CurrentCulture, "Entering constructor for: {0}", this.ToString()));
        }



        /////////////////////////////////////////////////////////////////////////////
        // Overridden Package Implementation
        #region Package Members

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override void Initialize()
        {
            Debug.WriteLine (string.Format(CultureInfo.CurrentCulture, "Entering Initialize() of: {0}", this.ToString()));
            base.Initialize();
            _applicationObject = (DTE2)this.GetService(typeof(DTE));
            // Add our command handlers for menu (commands must exist in the .vsct file)
            OleMenuCommandService mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if ( null != mcs )
            {
                // Create the command for the menu item.
                CommandID menuCommandID = new CommandID(GuidList.guidAttachIISPKGCmdSet, (int)PkgCmdIDList.cmdIdAttachIIS);
                MenuCommand menuItem = new MenuCommand(MenuItemCallback, menuCommandID );
                mcs.AddCommand( menuItem );
            }
        }
        #endregion

        /// <summary>
        /// This function is the callback used to execute a command when the a menu item is clicked.
        /// See the Initialize method to see how the menu item is associated to this function using
        /// the OleMenuCommandService service and the MenuCommand class.
        /// </summary>
        private void MenuItemCallback(object sender, EventArgs e)
        {
            // Show a Message Box to prove we were here
            //IVsUIShell uiShell = (IVsUIShell)GetService(typeof(SVsUIShell));
            //Guid clsid = Guid.Empty;
            //int result;
            //Microsoft.VisualStudio.ErrorHandler.ThrowOnFailure(uiShell.ShowMessageBox(
            //           0,
            //           ref clsid,
            //           "AttachIISPKG",
            //           string.Format(CultureInfo.CurrentCulture, "Inside {0}.MenuItemCallback()", this.ToString()),
            //           string.Empty,
            //           0,
            //           OLEMSGBUTTON.OLEMSGBUTTON_OK,
            //           OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST,
            //           OLEMSGICON.OLEMSGICON_INFO,
            //           0,        // false
            //           out result));

            LogToOutput("message from AttachIIS(CTRL+SHIFT+F1) - http://msprogrammer.serviciipeweb.ro/");
            var res = AttachTo("w3wp.exe");
            switch(res)
            {
                case AttachResult.Ok:
                    LogToOutput("attached");
                    return;
                case AttachResult.NotFound:
                    //TODO: intercept this case also
                    AttachTo("aspnet_wp.exe");
                    return;
                case AttachResult.Failed:
                    ShowMessage("Failed to attach. See Outout tool window for more details");
                    return;
                case AttachResult.FailedElevation:
                    ShowMessage("please run Visual Studio as administrator");
                    return;
                default:
                    ShowMessage("unknown error "+ res.ToString()) ;
                    return;

            }
                
            
        }
        
        private DTE2 _applicationObject;        
        void ShowMessage(string message)
        {
            IVsUIShell uiShell = (IVsUIShell)GetService(typeof(SVsUIShell));
            Guid clsid = Guid.Empty;
            int result;
            Microsoft.VisualStudio.ErrorHandler.ThrowOnFailure(uiShell.ShowMessageBox(
                       0,
                       ref clsid,
                       "AttachIISPKG",
                       message,
                       string.Empty,
                       0,
                       OLEMSGBUTTON.OLEMSGBUTTON_OK,
                       OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST,
                       OLEMSGICON.OLEMSGICON_INFO,
                       0,        // false
                       out result));
        }
        void LogToOutput(string message)
        {
            _applicationObject.ToolWindows.OutputWindow.ActivePane.OutputString(Environment.NewLine + message);
        }
        AttachResult AttachTo(string Name)
        {
            Name = Name.ToLower();
            var procs = _applicationObject.Debugger.LocalProcesses;
            foreach (var p in procs)
            {
                var proc = p as EnvDTE.Process;

                if (proc == null)
                    continue;

                string NameProc = Path.GetFileName(proc.Name).ToLower();
                if (NameProc != Name)
                    continue;

                try
                {
                    proc.Attach();
                    
                    try
                    {
                        string s = "attached to " + proc.Name + "->" + proc.ProcessID;
                        Console.WriteLine(s);
                        LogToOutput(s);
                    }
                    catch (Exception)
                    {
                    }
                    return AttachResult.Ok;
                }
                catch (COMException ex)
                {
                    if (ex.ErrorCode == -2147221447)
                    {
                        try
                        {
                            string s =  "please run in elevation mode(as administrator)!!!!!!!!!!" ;
                            Console.WriteLine(s);
                            LogToOutput(s);


                        }
                        catch (Exception)
                        {
                        }
                        return AttachResult.FailedElevation;
                    }
                    else
                    {
                        try
                        {
                            Console.WriteLine(ex.Message);
                            LogToOutput(ex.Message);
                        }
                        catch (Exception)
                        {
                        }
                        return AttachResult.Failed;
                    }
                }


            }
            return AttachResult.NotFound;

        }
    }
}
