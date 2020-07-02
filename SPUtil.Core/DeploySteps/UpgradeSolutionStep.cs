using Microsoft.VisualStudio.SharePoint;
using Microsoft.VisualStudio.SharePoint.Deployment;
using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CommandConsts = SPUtil.Commands.Consts;

namespace SPUtil.Core.DeploySteps
{
    // from https://docs.microsoft.com/en-us/visualstudio/sharepoint/walkthrough-creating-a-custom-deployment-step-for-sharepoint-projects?view=vs-2019
    [Export(typeof(IDeploymentStep))]
    [DeploymentStep(Consts.Steps.UpgradeSolution)]

    internal class UpgradeStep : IDeploymentStep
    {
        private string solutionName;
        private string solutionFullPath;

        public void Initialize(IDeploymentStepInfo stepInfo)
        {
            stepInfo.Name = "Upgrade solution";
            stepInfo.StatusBarMessage = "Upgrading solution...";
            stepInfo.Description = "Upgrades the solution on the local machine.";
        }

        public bool CanExecute(IDeploymentContext context)
        {
            solutionName = (context.Project.Package.Model.Name + ".wsp").ToLower();
            solutionFullPath = context.Project.Package.OutputPath;
            bool solutionExists = context.Project.SharePointConnection.ExecuteCommand<string, bool>(
                CommandConsts.IsSolutionDeployed, solutionName);

            if (context.Project.IsSandboxedSolution)
            {
                string sandboxMessage = "Cannot upgrade the solution. The upgrade deployment configuration " +
                    "does not support Sandboxed solutions.";
                context.Logger.WriteLine(sandboxMessage, LogCategory.Error);
                throw new InvalidOperationException(sandboxMessage);
            }
            else if (!solutionExists)
            {
                string notDeployedMessage = string.Format("Cannot upgrade the solution. The IsSolutionDeployed " +
                    "command cannot find the following solution: {0}.", solutionName);
                context.Logger.WriteLine(notDeployedMessage, LogCategory.Error);
                throw new InvalidOperationException(notDeployedMessage);
            }

            // Execute step and continue with deployment.
            return true;
        }

        // Implements IDeploymentStep.Execute.
        public void Execute(IDeploymentContext context)
        {
            context.Logger.WriteLine("Upgrading solution: " + solutionName, LogCategory.Status);
            context.Project.SharePointConnection
                .ExecuteCommand(
                    CommandConsts.UpgradeSolution,
                    solutionFullPath);

            context.Logger.WriteLine("Solution updated, scan for new features", LogCategory.Status);
            context.Project.SharePointConnection
                .ExecuteCommand(
                    CommandConsts.EnsureFeatures,
                    context.Project.Package.Model.SolutionId);

            context.Logger.WriteLine("Upgrade completed", LogCategory.Status);
        }
    }
}
