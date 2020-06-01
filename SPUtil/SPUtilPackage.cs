using System;
using System.Runtime.InteropServices;
using System.Threading;
using EnvDTE;
using Microsoft.VisualStudio.SharePoint;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;

namespace SPUtil
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the
    /// IVsPackage interface and uses the registration attributes defined in the framework to
    /// register itself and its components with the shell. These attributes tell the pkgdef creation
    /// utility what data to put into .pkgdef file.
    /// </para>
    /// <para>
    /// To get loaded into VS, the package must be referred by &lt;Asset Type="Microsoft.VisualStudio.VsPackage" ...&gt; in .vsixmanifest file.
    /// </para>
    /// </remarks>
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [Guid(SPUtilPackage.PackageGuidString)]
    [ProvideAutoLoad(UIContextGuids.SolutionExists, PackageAutoLoadFlags.BackgroundLoad)]
    public sealed class SPUtilPackage : AsyncPackage
    {
        /// <summary>
        /// SPUtilPackage GUID string.
        /// </summary>
        public const string PackageGuidString = "8793cbd1-d58c-47f6-94b5-c4906b71a91e";

        #region Package Members

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        /// <param name="cancellationToken">A cancellation token to monitor for initialization cancellation, which can occur when VS is shutting down.</param>
        /// <param name="progress">A provider for progress updates.</param>
        /// <returns>A task representing the async work of package initialization, or an already completed task if there is none. Do not return null from this method.</returns>
        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);
            progress.Report(new ServiceProgressData("SPUtil loading...", "Get Sharepoint Service", 1, 4));
            var sharePointProjectService = await GetServiceAsync(typeof(ISharePointProjectService)) as ISharePointProjectService;
            if (sharePointProjectService != null && sharePointProjectService.IsSharePointInstalled)
            {
                progress.Report(new ServiceProgressData("SPUtil loading...", "Setup events", 2, 4));

                var dte2 = await GetServiceAsync(typeof(SDTE)) as EnvDTE80.DTE2;
                var rdt = await GetServiceAsync(typeof(SVsRunningDocumentTable)) as IVsRunningDocumentTable;

                var eventHandler = new SPUtil.Core.EventHandler(sharePointProjectService, rdt);
                if (dte2 != null)
                {
                    progress.Report(new ServiceProgressData("SPUtil loading...", "Bind DTE", 3, 4));
                    eventHandler.RegisterDteDependentHandlers(dte2);
                }
                if (rdt != null)
                {
                    rdt.AdviseRunningDocTableEvents(eventHandler, out uint cookie);
                }
            }
            progress.Report(new ServiceProgressData("SPUtil loading...", "Done", 4, 4));
        }

        #endregion
    }
}
