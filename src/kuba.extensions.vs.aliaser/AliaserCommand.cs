//------------------------------------------------------------------------------
// <copyright file="AliaserCommand.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.ComponentModel.Design;
using System.Globalization;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace kuba.extensions.vs.aliaser
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class AliaserCommand
    {
        /// <summary>
        /// AliaserToAliasCommandId
        /// </summary>
        public const Int32 ALIASER_TO_ALIAS_COMMAND_ID = 0x0100;

        /// <summary>
        /// AliaserToInstanceCommandId
        /// </summary>
        public const Int32 ALIASER_TO_INSTANCE_COMMAND_ID = 0x0101;

        /// <summary>
        /// cmdAliaserToAlias
        /// </summary>
        public const Int32 CMD_ALIASER_TO_ALIAS = 0x0102;

        /// <summary>
        /// cmdAliaserToInstance
        /// </summary>
        public const Int32 CMD_ALIASER_TO_INSTANCE = 0x0103;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid MenuGroup = new Guid("ede25ded-2992-49c2-b248-d34aee8aa674");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;

        private readonly IAliaserService MainService;

        private EnvDTE.DTE DTEService
        {
            get
            {
                return (EnvDTE.DTE)this.ServiceProvider.GetService(typeof(SDTE)); 
            }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="AliaserCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private AliaserCommand(Package package)
        {
            if (package == null)
            {
                throw new ArgumentNullException("package");
            }

            this.package = package;
            this.MainService = new AliaserService();

            OleMenuCommandService commandService = this.ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                CommandID menuToAliasCommandID = new CommandID(MenuGroup, ALIASER_TO_ALIAS_COMMAND_ID);
                CommandID menuToInstanceCommandID = new CommandID(MenuGroup, ALIASER_TO_INSTANCE_COMMAND_ID);

                EventHandler transformAliasToInstace = this.TransformAliasToInstace;
                EventHandler transformInstaceToAlias = this.TransformInstaceToAlias;

                MenuCommand menuItemToInstace = new MenuCommand(transformAliasToInstace, menuToInstanceCommandID);
                MenuCommand menuItemToAlias = new MenuCommand(transformInstaceToAlias, menuToAliasCommandID);

                commandService.AddCommand(menuItemToInstace);
                commandService.AddCommand(menuItemToAlias);
            }
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static AliaserCommand Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private IServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static void Initialize(Package package)
        {
            Instance = new AliaserCommand(package);
        }

        /// <summary>
        /// Shows a message box when the menu item is clicked.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void TransformAliasToInstace(object sender, EventArgs e)
        {
            String selectedText = this.GetSelectedText();
            Boolean wasChangesMade = false;
            String newText = this.MainService.GetTransformedText(selectedText, TransformationFlow.TO_INSTANCE, out wasChangesMade);

            if (wasChangesMade)
                this.ReplaceSelectedText(newText);
        }        

        /// <summary>
        /// Shows a message box when the menu item is clicked.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void TransformInstaceToAlias(object sender, EventArgs e)
        {
            String selectedText = this.GetSelectedText();
            Boolean wasChangesMade = false;
            String newText = this.MainService.GetTransformedText(selectedText, TransformationFlow.TO_ALIAS, out wasChangesMade);

            if(wasChangesMade)
                this.ReplaceSelectedText(newText);
        }

        private String GetSelectedText()
        {
            String result = String.Empty;
            EnvDTE.DTE app = this.DTEService;
            if (app.ActiveDocument != null && app.ActiveDocument.Type == "Text")
            {
                Object objTtext = app.ActiveDocument.Object(String.Empty);
                if (objTtext != null && objTtext is EnvDTE.TextDocument)
                {
                    EnvDTE.TextDocument text = (EnvDTE.TextDocument)objTtext;
                    if (!text.Selection.IsEmpty)
                    {
                        result = text.Selection.Text;
                    }
                }
            }
            return result;
        }

        private void ReplaceSelectedText(String newText)
        {
            EnvDTE.DTE app = this.DTEService;
            if (app.ActiveDocument != null && app.ActiveDocument.Type == "Text")
            {
                Object objTtext = app.ActiveDocument.Object(String.Empty);

                if (objTtext != null && objTtext is EnvDTE.TextDocument)
                {
                    EnvDTE.TextDocument text = (EnvDTE.TextDocument)objTtext;
                    if (!text.Selection.IsEmpty)
                    {
                        text.Selection.Text = newText;
                    }
                }
            }
        }
    }
}
