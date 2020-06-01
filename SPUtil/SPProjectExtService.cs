using Microsoft.VisualStudio.SharePoint;
using Microsoft.Web.Administration;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SPExt.Core.Services
{
    public class SPProjectExtService
    {
        private readonly ISharePointProject _spProject;

        public SPProjectExtService(ISharePointProject project)
        {
            _spProject = project;
        }

        public string[] GetDeployableAssemblies(out string[] missing)
        {
            var projectOutput = Path.GetDirectoryName(_spProject.OutputFullPath);

            var assemblies = new List<string>();
            var missingAssemblies = new List<string>();

            var packageAssemblies = 
                _spProject.Package.Model.Assemblies.Select(a => Path.GetFileName(a.Location)).ToList();

            if (_spProject.IncludeAssemblyInPackage)
                packageAssemblies.Add(Path.GetFileName(_spProject.OutputFullPath));

            var outputDlls = Directory.GetFiles(projectOutput, "*.dll");
            foreach (var assembly in packageAssemblies)
            {
                if (outputDlls.Any(dll => string.Compare(Path.GetFileName(dll), assembly, true) == 0))
                    assemblies.Add(Path.Combine(projectOutput, assembly));
                else
                    missingAssemblies.Add(assembly);
            }

            missing = missingAssemblies.ToArray();

            return assemblies.ToArray();
        }

        public void RecyclePools()
        {
            var mgr = new ServerManager();

            var uri = _spProject.SiteUrl;
            if (uri == null)
            {
                _spProject.ProjectService.Logger.WriteLine($"Site url not set, unable to recycle pools", LogCategory.Warning);
                return;
            }

            var scheme = uri.Scheme;
            var host = uri.Host;
            var port = uri.Port;

            foreach (var site in mgr.Sites)
            {
                foreach (var binding in site.Bindings)
                {
                    if (binding.Protocol == scheme && binding.EndPoint.Port == port &&
                        (binding.Host == host || binding.Host == ""))
                    {
                        var pool = site.Applications[0].ApplicationPoolName;
                        var state = mgr.ApplicationPools[pool].Recycle();
                        _spProject.ProjectService.Logger.WriteLine($"Recycling {pool} result {state}", LogCategory.Status);
                    }
                }
            }
        }
    }
}
