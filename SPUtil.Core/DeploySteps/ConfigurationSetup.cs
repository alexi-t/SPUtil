using Microsoft.VisualStudio.SharePoint;
using Microsoft.VisualStudio.SharePoint.Deployment;
using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SPUtil.Core.DeploySteps
{
    [Export(typeof(ISharePointProjectExtension))]
    class ConfigurationSetup : ISharePointProjectExtension
    {
        // Implements ISharePointProjectExtension.Initialize.
        public void Initialize(ISharePointProjectService projectService)
        {
            projectService.ProjectInitialized += ProjectInitialized;
        }

        // Creates the new deployment configuration.
        private void ProjectInitialized(object sender, SharePointProjectEventArgs e)
        {
            AddUpgradeSolution(e.Project);
        }

        private void AddUpgradeSolution(ISharePointProject project)
        {
            string[] deploymentSteps = new string[]
            {
                DeploymentStepIds.PreDeploymentCommand,
                DeploymentStepIds.RecycleApplicationPool,
                Consts.Steps.UpgradeSolution,
                DeploymentStepIds.PostDeploymentCommand
            };

            string[] retractionSteps = new string[]
            {
                DeploymentStepIds.RecycleApplicationPool,
                DeploymentStepIds.RetractSolution
            };

            IDeploymentConfiguration configuration = project.DeploymentConfigurations.Add(
                "Upgrade", deploymentSteps, retractionSteps);
            configuration.Description = "This is the upgrade deployment configuration";
        }


    }
}
