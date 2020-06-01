using System;
using System.IO;
using System.ComponentModel.Design;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using EnvDTE;
using Microsoft.VisualStudio.SharePoint;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using SPExt.Core.Services;
using Task = System.Threading.Tasks.Task;

namespace SPUtil.DeployToGAC
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class DeployCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("27469595-e5c9-4fad-9520-69a54b0545d6");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="DeployCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private DeployCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static DeployCommand Instance
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
            // Switch to the main thread - the call to AddCommand in DeployCommand's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new DeployCommand(package, commandService);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private async void Execute(object sender, EventArgs e)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            DTE dte = await this.ServiceProvider.GetServiceAsync(typeof(DTE)) as DTE;
            ISharePointProjectService sharePointProjectService =
                await this.ServiceProvider.GetServiceAsync(typeof(ISharePointProjectService)) as ISharePointProjectService;
            var selectedProjects = dte.ActiveSolutionProjects as Array;
            if (selectedProjects != null && selectedProjects.Length > 0)
            {
                var currentProject = sharePointProjectService.Convert<Project, ISharePointProject>((Project)selectedProjects.GetValue(0)) as ISharePointProject;
                var extService = new SPProjectExtService(currentProject);
                var assembliesToDeploy = extService.GetDeployableAssemblies(out string[] missing);
                if (assembliesToDeploy.Any())
                {
                    var gacService = new System.EnterpriseServices.Internal.Publish();
                    currentProject.ProjectService.Logger.ActivateOutputWindow();
                    currentProject.ProjectService.Logger.WriteLine("========= Copy to GAC ===========", LogCategory.Status);
                    foreach (var assembly in assembliesToDeploy)
                    {
                        currentProject.ProjectService.Logger.WriteLine(Path.GetFileName(assembly) + "...", LogCategory.Status);
                        gacService.GacInstall(assembly);
                    }
                    if (missing.Any())
                    {
                        currentProject.ProjectService.Logger.WriteLine("Missing:", LogCategory.Status);
                        foreach (var missingAssembly in missing)
                        {
                            currentProject.ProjectService.Logger.WriteLine(Path.GetFileName(missingAssembly), LogCategory.Status);
                        }
                    }
                    currentProject.ProjectService.Logger.WriteLine($"===== Copy to GAC done at {DateTime.Now:hh:mm:ss} =====", LogCategory.Status);

                    extService.RecyclePools();
                }
            }
        }
    }
}
