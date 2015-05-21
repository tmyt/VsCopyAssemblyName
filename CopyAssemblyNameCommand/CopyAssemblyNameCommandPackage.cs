namespace tmyt.CopyAssemblyNameCommand
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.ComponentModel.Design;
    using System.Diagnostics;
    using System.Globalization;
    using System.Linq;
    using System.Reflection;
    using System.Runtime.InteropServices;
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
    [InstalledProductRegistration("#110", "#112", "1.0")]
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
            Debug.WriteLine(string.Format(CultureInfo.CurrentCulture, "Entering Initialize() of: {0}", this.ToString()));
            base.Initialize();

            // Add our command handlers for menu (commands must exist in the .vsct file)
            OleMenuCommandService mcs = this.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (null != mcs)
            {
                // Create the command for the menu item.
                var cmdidCopyAssemblyNameOnSolutionExplorer = new CommandID(
                    GuidList.guidCopyAssemblyNameCommandSet,
                    (int)PkgCmdIDList.cmdidCopyAssemblyNameOnSolutionExplorer);
                var menuItemCopyAssemblyNameOnSolutionExplorer =
                    new MenuCommand(
                        this.OnCopyAssemblyNameOnSolutionExplorer,
                        cmdidCopyAssemblyNameOnSolutionExplorer);
                mcs.AddCommand(menuItemCopyAssemblyNameOnSolutionExplorer);

                var cmdidCopyNameOnObjectBrowser = new CommandID(
                    GuidList.guidCopyAssemblyNameCommandSet,
                    (int)PkgCmdIDList.cmdidCopyAssemblyNameOnObjectBrowser);
                var menuItemCopyNameOnObjectBrowser = new OleMenuCommand(
                    this.OnCopyNameOnObjectBrowser,
                    cmdidCopyNameOnObjectBrowser);
                menuItemCopyNameOnObjectBrowser.BeforeQueryStatus += this.OnBeforeQueryStatusCopyNameOnObjectBrowser;
                mcs.AddCommand(menuItemCopyNameOnObjectBrowser);
            }
        }

        #endregion

        /// <summary>
        /// This function is the callback used to execute a command when the a menu item is clicked.
        /// See the Initialize method to see how the menu item is associated to this function using
        /// the OleMenuCommandService service and the MenuCommand class.
        /// </summary>
        private void OnCopyAssemblyNameOnSolutionExplorer(object sender, EventArgs e)
        {
            // Copy assembly name
            var dte = this.GetService(typeof(DTE)) as DTE2;

            var refs = dte.ToolWindows.SolutionExplorer.SelectedItems as IEnumerable;
            if (refs == null)
            {
                return;
            }

            var assemblies = refs
                .OfType<UIHierarchyItem>()
                .Select(i => i.Object)
                .OfType<Reference>()
                .Select(d =>
                    {
                        string culture = d.Culture;
                        if (string.IsNullOrWhiteSpace(culture))
                        {
                            culture = "neutral";
                        }

                        var s = string.Format("{0}, Version={1}, Culture={2}", d.Name, d.Version, culture);
                        if (!string.IsNullOrWhiteSpace(d.PublicKeyToken))
                        {
                            s += string.Format(", PublicKeyToken={0}", d.PublicKeyToken.ToLowerInvariant());
                        }

                        return s;
                    });

            Clipboard.SetText(string.Join(Environment.NewLine, assemblies));
        }

        private void OnCopyNameOnObjectBrowser(object sender, EventArgs e)
        {
            var selections = this.GetCurrentSelections();
            string names = string.Join(Environment.NewLine, selections.Select(x => x.Name));

            Clipboard.SetText(names);
        }

        private void OnBeforeQueryStatusCopyNameOnObjectBrowser(object sender, EventArgs eventArgs)
        {
            var menu = sender as OleMenuCommand;
            if (menu == null)
            {
                return;
            }

            var selections = this.GetCurrentSelections();
            bool visible = selections.Any(x => x.Type == ItemType.Assembly || x.Type == ItemType.Type);

            menu.Visible = visible;
        }

        private IEnumerable<SelectedItem> GetCurrentSelections()
        {
            var navTools = this.GetService(typeof(SVsObjBrowser)) as IVsNavigationTool;
            if (navTools == null)
            {
                yield break;
            }

            IVsSelectedSymbols symbols;
            int hresult = navTools.GetSelectedSymbols(out symbols);
            Marshal.ThrowExceptionForHR(hresult);

            uint count;
            hresult = symbols.GetCount(out count);
            Marshal.ThrowExceptionForHR(hresult);

            for (uint i = 0; i < count; ++i)
            {
                IVsSelectedSymbol symbol;
                hresult = symbols.GetItem(i, out symbol);
                Marshal.ThrowExceptionForHR(hresult);

                string name;
                hresult = symbol.GetName(out name);
                Marshal.ThrowExceptionForHR(hresult);

                IVsNavInfo navInfo;
                hresult = symbol.GetNavInfo(out navInfo);
                Marshal.ThrowExceptionForHR(hresult);

                uint symType;
                hresult = navInfo.GetSymbolType(out symType);
                Marshal.ThrowExceptionForHR(hresult);

                switch ((_LIB_LISTTYPE)symType)
                {
                    case _LIB_LISTTYPE.LLT_CLASSES: // type
                        IVsEnumNavInfoNodes enin;
                        hresult = navInfo.EnumCanonicalNodes(out enin);
                        Marshal.ThrowExceptionForHR(hresult);

                        IVsNavInfoNode[] nodes = new IVsNavInfoNode[1];
                        uint fetched;

                        while (enin.Next(1, nodes, out fetched) >= 0 && fetched == 1)
                        {
                            var node = nodes[0];

                            uint type2;

                            hresult = node.get_Type(out type2);
                            Marshal.ThrowExceptionForHR(hresult);

                            if (type2 != (uint)_LIB_LISTTYPE.LLT_PHYSICALCONTAINERS)
                            {
                                continue;
                            }

                            string name2;
                            hresult = node.get_Name(out name2);
                            Marshal.ThrowExceptionForHR(hresult);

                            var asmName = AssemblyName.GetAssemblyName(name2);
                            string typeName = Assembly.CreateQualifiedName(asmName.FullName, name);

                            yield return new SelectedItem
                                {
                                    Name = typeName,
                                    Type = ItemType.Type
                                };
                        }

                        break;

                    case _LIB_LISTTYPE.LLT_PHYSICALCONTAINERS: // assembly
                        var asmName2 = AssemblyName.GetAssemblyName(name);

                        yield return new SelectedItem
                            {
                                Name = asmName2.FullName,
                                Type = ItemType.Type
                            };

                        break;

                    default:
                        break;
                }
            }
        }
    }
}
