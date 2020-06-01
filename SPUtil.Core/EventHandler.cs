using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.SharePoint;
using Microsoft.VisualStudio.Shell.Interop;
using SPUtil.Core.Service;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SPUtil.Core
{
    public class EventHandler: IVsRunningDocTableEvents
    {
        private readonly ISharePointProjectService _spProjService;
        private readonly IVsRunningDocumentTable _rdt;

        private DTE2 _dte;

        public EventHandler(
            ISharePointProjectService spProjectService,
            IVsRunningDocumentTable rdt
            )
        {
            _spProjService = spProjectService;
            _rdt = rdt;
        }

        public void RegisterDteDependentHandlers(DTE2 dte2)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();

            _dte = dte2;

            var events = (Events)dte2.Events;
            var buildEvent = (BuildEvents)events.BuildEvents;

            buildEvent.OnBuildDone += BuildEvents_OnBuildDone;
        }

        public void BuildEvents_OnBuildDone(vsBuildScope Scope, vsBuildAction Action)
        {
            // We will never auto copy on deploy - just build or rebuild.
            if (Action == vsBuildAction.vsBuildActionBuild || Action == vsBuildAction.vsBuildActionRebuildAll)
            {
                //// Get all farm SP projects where the auto copy flag is set, and where the project was built succesfully.
                //IEnumerable<ISharePointProject> spProjects = ProjectUtilities.GetSharePointProjects(Scope == vsBuildScope.vsBuildScopeSolution, true)
                //    .Where(proj => AutoCopyAssembliesProperty.GetFromProject(proj)
                //    && successfulBuiltProjects.Any(sbp => proj.FullPath.EndsWith(sbp)));
                //if (spProjects.Count() > 0)
                //{
                //    // We don't clear the log since we want this to show straight after the build details.
                //    DTEManager.SetStatus("Auto Copying to GAC/BIN...");
                //    DTEManager.ProjectService.Logger.WriteLine("========== Auto Copying to GAC/BIN ==========", LogCategory.Status);

                //    List<string> appPools = new List<string>();

                //    foreach (ISharePointProject spProject in spProjects)
                //    {
                //        new SharePointPackageArtefact(spProject).QuickCopyBinaries(true);

                //        // Get the app pool for recycling while we are here.
                //        string appPool = new ProcessUtilities(this.DTE).GetApplicationPoolName(CKSDEVPackageSharePointProjectService, spProject.SiteUrl.ToString());
                //        if (!String.IsNullOrEmpty(appPool))
                //        {
                //            appPools.Add(appPool);
                //        }
                //    }

                //    DTEManager.ProjectService.Logger.WriteLine("========== Auto Copy to GAC/BIN succeeded ==========", LogCategory.Status);

                //    DTEManager.SetStatus("Auto Copying to GAC/BIN... All Done!");

                //    if (appPools.Count > 0)
                //    {
                //        this.RecycleAppPools(appPools.Distinct().ToArray());
                //    }
                //    else
                //    {
                //        this.RestartIIS();
                //    }
                //}
            }
        }

        public void OnItemSave(ProjectItem projectItem)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();
            
            try
            {
                EnvDTE.Project dteProject = projectItem.ContainingProject;
                ISharePointProject spProject = _spProjService.Convert<EnvDTE.Project, ISharePointProject>(dteProject);

                if (spProject != null)
                {
                    bool isAutoCopyToRoot = new ProjectProperties.PropertySource(spProject).AutoCopy;
                    if (isAutoCopyToRoot)
                    {
                        var service = new CopyToRootService(_spProjService);
                        service.CopyProjectItem(projectItem);
                    }
                }
            }
            catch (Exception ex)
            {
                _spProjService.Logger.WriteLine($"Error copy file {ex}", LogCategory.Warning);
            }
        }

        public int OnAfterFirstDocumentLock(uint docCookie, uint dwRDTLockType, uint dwReadLocksRemaining, uint dwEditLocksRemaining)
        {
            return VSConstants.S_OK;
        }

        public int OnBeforeLastDocumentUnlock(uint docCookie, uint dwRDTLockType, uint dwReadLocksRemaining, uint dwEditLocksRemaining)
        {
            return VSConstants.S_OK;
        }

        public int OnAfterSave(uint docCookie)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();

            _rdt.GetDocumentInfo(docCookie, out uint flags, out uint rLocks, out uint eLocks, out string path, out IVsHierarchy hierarchy, out uint projectItemId, out IntPtr unknow);

            try
            {
                if (VSConstants.VSITEMID_NIL != projectItemId &&
                        VSConstants.VSITEMID_ROOT != projectItemId &&
                        VSConstants.VSITEMID_SELECTION != projectItemId &&
                        VSConstants.S_OK == hierarchy.GetProperty(projectItemId, (int)__VSHPROPID.VSHPROPID_ExtObject, out object objProj))
                {
                    var projectItem = objProj as EnvDTE.ProjectItem;
                    OnItemSave(projectItem);
                }
            }
            catch
            {

            }


            return VSConstants.S_OK;
        }

        public int OnAfterAttributeChange(uint docCookie, uint grfAttribs)
        {
            return VSConstants.S_OK;
        }

        public int OnBeforeDocumentWindowShow(uint docCookie, int fFirstShow, IVsWindowFrame pFrame)
        {
            return VSConstants.S_OK;
        }

        public int OnAfterDocumentWindowHide(uint docCookie, IVsWindowFrame pFrame)
        {
            return VSConstants.S_OK;
        }
    }
}
