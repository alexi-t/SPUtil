using System;
using System.ComponentModel.Design;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using EnvDTE;
using Microsoft.VisualStudio.SharePoint;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using SPUtil.Core.Service;
using Task = System.Threading.Tasks.Task;

namespace SPUtil.CopyToRoot
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class CopyToRoot
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandItemId = 0x0101;
        public const int CommandFldId = 0x0102;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("b1adb76f-4720-451d-a3d5-1fb8e21615d2");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="CopyToRoot"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private CopyToRoot(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var copyItemCommandID = new CommandID(CommandSet, CommandItemId);
            var copyProjectItemCommand = new OleMenuCommand(this.CopyItemHandler, copyItemCommandID);
            copyProjectItemCommand.BeforeQueryStatus += CopyItemStatusCheck;
            commandService.AddCommand(copyProjectItemCommand);

            var copyFldCommandID = new CommandID(CommandSet, CommandFldId);
            var copyFldCommand = new OleMenuCommand(this.CopyItemHandler, copyFldCommandID);
            copyFldCommand.BeforeQueryStatus += CopyItemStatusCheck;
            commandService.AddCommand(copyFldCommand);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static CopyToRoot Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
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
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in CopyToRoot's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new CopyToRoot(package, commandService);
        }

        private void CopyItemStatusCheck(object sender, EventArgs e)
        {
            var command = sender as OleMenuCommand;
            if (command == null)
                return;

            ThreadHelper.JoinableTaskFactory.Run(async () =>
            {
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
                var dte = await this.ServiceProvider.GetServiceAsync(typeof(DTE)) as DTE;
                var sharePointProjectService =
                    await this.ServiceProvider.GetServiceAsync(typeof(ISharePointProjectService)) as ISharePointProjectService;
                if (dte == null || sharePointProjectService == null || !sharePointProjectService.IsSharePointInstalled)
                    return;

                var service = new CopyToRootService(sharePointProjectService);


                foreach (SelectedItem selectedItem in dte.SelectedItems)
                {
                    var artifacts = service.FlattenToArtifacts(selectedItem.ProjectItem);
                    if (artifacts.Count == 0)
                    {
                        command.Enabled = false;
                        command.Visible = false;
                    }
                    else
                    {
                        command.Enabled = true;
                        command.Visible = true;
                    }
                }
            });
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void CopyItemHandler(object sender, EventArgs e)
        {
            ThreadHelper.JoinableTaskFactory.Run(async () =>
            {
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

                var dte = await this.ServiceProvider.GetServiceAsync(typeof(DTE)) as DTE;
                var sharePointProjectService =
                    await this.ServiceProvider.GetServiceAsync(typeof(ISharePointProjectService)) as ISharePointProjectService;
                if (dte == null || sharePointProjectService == null || !sharePointProjectService.IsSharePointInstalled)
                    return;

                var service = new CopyToRootService(sharePointProjectService);

                foreach (SelectedItem selectedItem in dte.SelectedItems)
                {
                    service.CopyProjectItem(selectedItem.ProjectItem);
                }
            });

        }
    }
}
