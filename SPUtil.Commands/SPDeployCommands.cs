using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Microsoft.VisualStudio.SharePoint.Commands;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace SPUtil.Commands
{
    internal class SPDeployCommands
    {
        // Determines whether the specified solution has been deployed to the local SharePoint server.
        [SharePointCommand(Consts.IsSolutionDeployed)]
        private bool IsSolutionDeployed(ISharePointCommandContext context, string solutionName)
        {
            SPSolution solution = SPFarm.Local.Solutions[solutionName];
            return solution != null;
        }

        // Upgrades the specified solution to the local SharePoint server.
        [SharePointCommand(Consts.UpgradeSolution)]
        private void UpgradeSolution(ISharePointCommandContext context, string fullWspPath)
        {
            var solution = SPFarm.Local.Solutions[Path.GetFileName(fullWspPath)];

            if (solution == null)
            {
                throw new InvalidOperationException("The solution has not been deployed.");
            }

            solution.Upgrade(fullWspPath);
        }

        [SharePointCommand(Consts.EnsureFeatures)]
        private void EnsureFeatures(ISharePointCommandContext context, Guid solutionId)
        {
            SPFarm.Local.FeatureDefinitions.ScanForFeatures(solutionId, false, false);
        }

        [SharePointCommand(Consts.UpdateFeatures)]
        private void UpdateFeatures(ISharePointCommandContext context, Guid[] featureIds)
        {
            context.Logger.WriteLine($"Update features CMD {string.Join(",", featureIds)}", LogCategory.Verbose);

            using (var site = new SPSite(context.Site.ID))
            using (var web = site.OpenWeb(context.Web.ID))
            {
                var farm = site.WebApplication.Farm;
                foreach (var featureId in featureIds)
                {
                    context.Logger.WriteLine($"Looking for feature with Id {featureId}", LogCategory.Verbose);
                    try
                    {
                        var featureDefinition = farm.FeatureDefinitions[featureId];

                        context.Logger.WriteLine(
                            $"Found feature definition {featureId}/{featureDefinition.DisplayName}/{featureDefinition.Scope}@{featureDefinition.Version}",
                            LogCategory.Status);
                        if (featureDefinition != null)
                        {

                            SPFeature feature = null;
                            switch (featureDefinition.Scope)
                            {
                                case SPFeatureScope.WebApplication:
                                    feature = site.WebApplication.Features[featureId];
                                    break;
                                case SPFeatureScope.Farm:
                                case SPFeatureScope.Site:
                                    feature = site.Features[featureId];
                                    break;
                                case SPFeatureScope.Web:
                                    feature = web.Features[featureId];
                                    break;
                            }
                            if (feature == null)
                            {
                                context.Logger.WriteLine($"No local feature found", LogCategory.Status);
                                continue;
                            }
                            if (featureDefinition.Version != feature.Version)
                            {
                                context.Logger.WriteLine($"Found local feature with version {feature.Version}, upgrading", LogCategory.Status);
                                feature.Upgrade(false);
                                context.Logger.WriteLine($"Successfully upgraded to {feature.Version}", LogCategory.Status);
                            }
                            else
                            {
                                context.Logger.WriteLine($"Feature up to date at {feature.Version}", LogCategory.Status);
                            }

                        }
                        else
                        {
                            context.Logger.WriteLine($"No feature definition found for {featureId}", LogCategory.Status);
                        }
                    }
                    catch (Exception ex)
                    {
                        context.Logger.WriteLine($"Error update feature {featureId}: {ex.Message}", LogCategory.Status);
                    }
                }
            }
        }
    }
}
