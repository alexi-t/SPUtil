using EnvDTE;
using Microsoft.VisualStudio.SharePoint;
using SPUtil.Core.Service.Model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SPUtil.Core.Service
{
    public class CopyToRootService
    {
        private readonly ISharePointProjectService _spProjService;

        public CopyToRootService(ISharePointProjectService spProjectService)
        {
            _spProjService = spProjectService;
        }

        private List<IArtifact> FlattenToArtifacts(ProjectItem projectItem)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();

            var list = new List<IArtifact>();
            try
            {
                var spFileProjItem = _spProjService.Convert<ProjectItem, ISharePointProjectItemFile>(projectItem);
                if (spFileProjItem != null)
                {
                    list.Add(new SPFileArtifact(spFileProjItem));
                }
            }
            catch { }

            try
            {
                if (projectItem.ProjectItems.Count > 0)
                {
                    for (int i = 1; i <= projectItem.ProjectItems.Count; i++)
                    {
                        ProjectItem childItem = projectItem.ProjectItems.Item(i);
                        list.AddRange(FlattenToArtifacts(childItem));
                    }
                }

            }
            catch { }

            return list;
        }
        
        public void CopyProjectItem(ProjectItem item)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();

            var log = _spProjService.Logger;

            log.ActivateOutputWindow();

            log.WriteLine("======= Start copy artifacts =======", LogCategory.Status);

            var artifacts = FlattenToArtifacts(item);

            log.WriteLine($"Total artifacts count: {artifacts.Count}", LogCategory.Status);

            foreach (IArtifact artifact in artifacts)
            {
                try
                {
                    artifact.QuickDeploy();
                }
                catch (Exception ex)
                {
                    log.WriteLine($"Unhandled error copy artifact {ex}", LogCategory.Status);
                }
            }

            log.WriteLine("======= End copy artifacts =======", LogCategory.Status);

        }
    }
}
