using Microsoft.VisualStudio;
using Microsoft.VisualStudio.ProjectSystem;
using Microsoft.VisualStudio.SharePoint;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Threading;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SPUtil.ProjectProperties
{
    public class PropertySource
    {
        private ISharePointProject sharePointProject;
        private IVsBuildPropertyStorage projectStorage;

        public PropertySource(ISharePointProject myProject)
        {
            sharePointProject = myProject;
            projectStorage = sharePointProject.ProjectService.Convert<ISharePointProject, IVsBuildPropertyStorage>(sharePointProject);
        }

        private const string AutoCopyToRootPropertyId = "SPUtilAutoCopyToRoot";
        private const string AutoCopyToRootDefaultValue = "false";

        [DisplayName("Auto copy to SP root")]
        [DescriptionAttribute("Copy files to SP Hive folder on save")]
        [DefaultValue(false)]
        [Category("SP Util")]
        public bool AutoCopy
        {
            get
            {
                Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();
                string propertyValue;
                
                int hr = projectStorage.GetPropertyValue(AutoCopyToRootPropertyId, string.Empty,
                    (uint)_PersistStorageType.PST_PROJECT_FILE, out propertyValue);

                // Try to get the current value from the project file; if it does not yet exist, return a default value.
                if (!ErrorHandler.Succeeded(hr) || String.IsNullOrEmpty(propertyValue))
                {
                    propertyValue = AutoCopyToRootDefaultValue;
                }

                return bool.Parse(propertyValue);
            }

            set
            {
                Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();

                // Do not save the default value.
                if (value.ToString().ToLower() != AutoCopyToRootDefaultValue)
                {
                    projectStorage.SetPropertyValue(AutoCopyToRootPropertyId, string.Empty,
                        (uint)_PersistStorageType.PST_PROJECT_FILE, value.ToString().ToLower());
                }
            }
        }

        private const string AutoDeployToGACPropertyId = "SPUtilAutoDeployToGAC";
        private const string AutoDeployToGACDefaultValue = "false";

        [DisplayName("Auto deploy to GAC")]
        [DescriptionAttribute("Auto build and deploy project to GAC on build")]
        [DefaultValue(false)]
        [Category("SP Util")]
        public bool AutoDeployToGAC
        {
            get
            {
                Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();

                string propertyValue;

                int hr = projectStorage.GetPropertyValue(AutoDeployToGACPropertyId, string.Empty,
                    (uint)_PersistStorageType.PST_PROJECT_FILE, out propertyValue);

                // Try to get the current value from the project file; if it does not yet exist, return a default value.
                if (!ErrorHandler.Succeeded(hr) || String.IsNullOrEmpty(propertyValue))
                {
                    propertyValue = AutoDeployToGACDefaultValue;
                }

                return bool.Parse(propertyValue);
            }

            set
            {
                Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();

                if (value.ToString().ToLower() != AutoDeployToGACDefaultValue)
                {
                    projectStorage.SetPropertyValue(AutoDeployToGACPropertyId, string.Empty,
                        (uint)_PersistStorageType.PST_PROJECT_FILE, value.ToString().ToLower());
                }
            }
        }
    }
}
