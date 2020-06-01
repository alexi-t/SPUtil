using Microsoft.VisualStudio.SharePoint;
using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SPUtil.ProjectProperties
{
    [Export(typeof(ISharePointProjectExtension))]
    public class DeploymentProjectExtension : ISharePointProjectExtension
    {
        public void Initialize(ISharePointProjectService projectService)
        {
            if (projectService.IsSharePointInstalled)
                projectService.ProjectPropertiesRequested += new EventHandler<SharePointProjectPropertiesRequestedEventArgs>(projectService_ProjectPropertiesRequested);
        }

        void projectService_ProjectPropertiesRequested(object sender, SharePointProjectPropertiesRequestedEventArgs e)
        {
            if (!e.Project.IsSandboxedSolution)
            {
                e.PropertySources.Add(new PropertySource(e.Project));
            }
        }

    }
}
