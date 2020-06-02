using EnvDTE;
using Microsoft.Build.Tasks;
using Microsoft.VisualStudio.SharePoint;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SPUtil.Core.Service.Model
{
    public class SPFileArtifact : IArtifact
    {
        private readonly ISharePointProjectService _service;
        private readonly ISharePointProjectItem _projItem;
        private readonly ISharePointProjectItemFile _item;
        private readonly ISharePointProjectLogger _log;

        public SPFileArtifact(ISharePointProjectItemFile sharePointProjectItemFile)
        {
            _service = sharePointProjectItemFile.Project.ProjectService;
            _projItem = sharePointProjectItemFile.ProjectItem;
            _item = sharePointProjectItemFile;
            _log = sharePointProjectItemFile.Project.ProjectService.Logger;
        }

        private void CopyFile(string targetPath, string sourcePath)
        {
            try
            {
                File.Copy(sourcePath, targetPath, true);
                _log.WriteLine($"Copy {sourcePath} -> {targetPath}", LogCategory.Status);
            }
            catch (Exception ex)
            {
                _log.WriteLine($"Error copy {Path.GetFileName(_item.FullPath)}: {ex.Message}", LogCategory.Status);
            }
        }

        private void DeployRootFile()
        {
            var targetPath =
                Path.Combine(
                    Path.Combine(_item.DeploymentRoot.Replace("{SharePointRoot}", _service.SharePointInstallPath), _item.DeploymentPath),
                    Path.GetFileName(_item.FullPath)
                );
            var sourcePath = _item.FullPath;

            CopyFile(targetPath, sourcePath);
        }

        private void DeployFeatureFile()
        {
            var containingFeatures = new List<ISharePointProjectFeature>();
            _service.Projects.ToList().ForEach(p =>
            {
                containingFeatures.AddRange(p.Features.Where(f => f.ProjectItems.Any(i => i.Id == _projItem.Id)));
            });
            foreach (var feature in containingFeatures)
            {
                var featureName = $"{feature.Project.Name}_{feature.Name}";

                var targetPath =
                    Path.Combine(
                        Path.Combine(_item.DeploymentRoot, _item.DeploymentPath)
                            .Replace("{SharePointRoot}", _service.SharePointInstallPath)
                            .Replace("{FeatureName}", featureName),
                        Path.GetFileName(_item.FullPath)
                    );
                var sourcePath = _item.FullPath;

                CopyFile(targetPath, sourcePath);
            }
            if (!containingFeatures.Any())
                _log.WriteLine($"File {_item.Name} not a path of any feature, skip", LogCategory.Status);
        }

        public void QuickDeploy()
        {
            if (_item.DeploymentType == DeploymentType.TemplateFile ||
                _item.DeploymentType == DeploymentType.RootFile)
                DeployRootFile();

            if (_item.DeploymentType == DeploymentType.ElementFile)
                DeployFeatureFile();
        }
    }
}
