using EnvDTE;
using Microsoft.VisualStudio.SharePoint;
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

        
        public void CopyProjectItem(ProjectItem item)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();

            var logger = _spProjService.Logger;


            logger.ActivateOutputWindow();

            try
            {
                var spFile = _spProjService.Convert<ProjectItem, ISharePointProjectItemFile>(item);
                if (spFile != null)
                {
                    if (spFile.DeploymentType == DeploymentType.TemplateFile ||
                        spFile.DeploymentType == DeploymentType.RootFile)
                    {
                        var targetPath =
                            Path.Combine(
                                Path.Combine(spFile.DeploymentRoot, spFile.DeploymentPath).Replace("{SharePointRoot}", _spProjService.SharePointInstallPath),
                                Path.GetFileName(spFile.FullPath)
                            );
                        var sourcePath = spFile.FullPath;
                        File.Copy(sourcePath, targetPath, true);
                        logger.WriteLine($"Copy {sourcePath} -> {targetPath}", LogCategory.Status);

                    }
                }
            }
            catch { }

            // See if this item is an SPI.
            try
            {
                var spItem = _spProjService.Convert<ProjectItem, ISharePointProjectItem>(item);
                if (spItem != null)
                {

                }
            }
            catch { }

            // See if this item is a Feature.
            try
            {
                var spFeature = _spProjService.Convert<ProjectItem, ISharePointProjectFeature>(item);
                if (spFeature != null)
                {
                }
            }
            catch { }

            try
            {
                if (item.ProjectItems.Count > 0)
                {
                    for (int i = 1; i <= item.ProjectItems.Count; i++)
                    {
                        ProjectItem childItem = item.ProjectItems.Item(i);
                        CopyProjectItem(childItem);
                    }
                }

            }
            catch { }
        }
    }
}
