﻿namespace tmyt.CopyAssemblyNameCommand
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.Globalization;
    using System.Runtime.InteropServices;
    using System.ComponentModel.Design;
    using System.Linq;
    using System.Text;
    using System.Windows.Forms;

    using EnvDTE;
    using EnvDTE80;
    using Microsoft.VisualStudio.Shell;
    using Microsoft.VisualStudio.Shell.Interop;

    using VSLangProj;
    
    using VSConstants = Microsoft.VisualStudio.Shell.Interop.Constants;

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
    [Guid(GuidList.guidCopyAssemblyNameCommandPkgString)]
    public sealed class CopyAssemblyNameCommandPackage : Package
    {
        /// <summary>
        /// Default constructor of the package.
        /// Inside this method you can place any initialization code that does not require 
        /// any Visual Studio service because at this point the package object is created but 
        /// not sited yet inside Visual Studio environment. The place to do all the other 
        /// initialization is the Initialize method.
        /// </summary>
        public CopyAssemblyNameCommandPackage()
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

            // Add our command handlers for menu (commands must exist in the .vsct file)
            OleMenuCommandService mcs = this.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (null != mcs)
            {
                // Create the command for the menu item.
                var cmdidCopyAssemblyNameOnSolutionExplorer = new CommandID(GuidList.guidCopyAssemblyNameCommandSet, (int)PkgCmdIDList.cmdidCopyAssemblyNameOnSolutionExplorer);
                var menuItemCopyAssemblyNameOnSolutionExplorer = new MenuCommand(this.OnCopyAssemblyNameMenuCommandOnSolutionExplorer, cmdidCopyAssemblyNameOnSolutionExplorer);
                mcs.AddCommand(menuItemCopyAssemblyNameOnSolutionExplorer);

                var cmdidCopyAssemblyNameOnObjectBrowser = new CommandID(GuidList.guidCopyAssemblyNameCommandSet, (int)PkgCmdIDList.cmdidCopyAssemblyNameOnObjectBrowser);
                var menuItemCopyAssemblyNameOnObjectBrowser = new MenuCommand(this.OnCopyAssemblyNameMenuCommandOnObjectBrowser, cmdidCopyAssemblyNameOnObjectBrowser);
                mcs.AddCommand(menuItemCopyAssemblyNameOnObjectBrowser);

                var cmdidCopyFullTypeNameOnObjectrowser = new CommandID(GuidList.guidCopyAssemblyNameCommandSet, (int)PkgCmdIDList.cmdidCopyFullTypeNameOnObjectrowser);
                var menuItemCopyFullTypeNameOnObjectrowser = new MenuCommand(this.OnCopyFullTypeNameMenuCommandOnObjectBrowser, cmdidCopyFullTypeNameOnObjectrowser);
                mcs.AddCommand(menuItemCopyFullTypeNameOnObjectrowser);
            }
        }
        #endregion

        /// <summary>
        /// This function is the callback used to execute a command when the a menu item is clicked.
        /// See the Initialize method to see how the menu item is associated to this function using
        /// the OleMenuCommandService service and the MenuCommand class.
        /// </summary>
        private void OnCopyAssemblyNameMenuCommandOnSolutionExplorer(object sender, EventArgs e)
        {
            // Copy assembly name
            var dte = this.GetService(typeof(DTE)) as DTE2;
            var refs = dte.ToolWindows.SolutionExplorer.SelectedItems as IEnumerable;

            if (refs == null)
            {
                return;
            }

            var assemblies = refs.OfType<UIHierarchyItem>().Select(i => i.Object).OfType<Reference>()
                .Select(d =>
                {
                    var s = string.Format("{0}, Version={1}, Culture={2}", d.Name, d.Version, string.IsNullOrWhiteSpace(d.Culture) ? "neutral" : d.Culture);
                    if(!string.IsNullOrWhiteSpace(d.PublicKeyToken)) s += string.Format(", PublicKeyToken={0}", d.PublicKeyToken.ToLowerInvariant());
                    return s;
                })
                .ToArray();

            Clipboard.SetText(string.Join("\r\n", assemblies));
        }

        private void OnCopyAssemblyNameMenuCommandOnObjectBrowser(object sender, EventArgs e)
        {
            var browser = this.GetService(typeof(SVsObjBrowser)) as IVsNavigationTool;

            IVsSelectedSymbols symbols;
            browser.GetSelectedSymbols(out symbols);
        }

        private void OnCopyFullTypeNameMenuCommandOnObjectBrowser(object sender, EventArgs e)
        {
        }
    }
}
