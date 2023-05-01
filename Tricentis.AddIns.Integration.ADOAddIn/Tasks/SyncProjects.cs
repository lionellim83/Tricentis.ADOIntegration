using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Tricentis.AddIns.Integration.ADO;
using Tricentis.TCCore.BusinessObjects.Folders;
using Tricentis.TCCore.Persistency;
using Tricentis.TCCore.Persistency.Tasks;

namespace Tricentis.AddIns.Integration.ADOAddIn.Tasks
{
    public class SyncProjects : BaseADOTask
    {
        public SyncProjects(PersistableObject obj)
        {
            _targetObject = obj;
        }

        public override string Name => "Synchronise All Projects";

        public override TaskCategory Category => new TaskCategory(Constants.AddInCategory);

        public override object Execute(PersistableObject obj, ITaskContext context)
        {
            if (obj is TCFolder)
            {
                TCFolder tcFolder = (TCFolder)obj;
                String[] content = tcFolder.PossibleContent.Split(';');

                if (content.Contains(Constants.PossibleContent_TestCase))
                {
                    SyncTestCasesAllProjects(obj, context);
                }
            }

            return null;
        }

        public void SyncTestCasesAllProjects(PersistableObject obj, ITaskContext context)
        {
            List<PersistableObject> tcFolders = obj.SearchByTQL($"=>SUBPARTS:TCFolder[{Constants.ADOSync}==\"True\"]").ToList();

            foreach (TCFolder folder in tcFolders)
            {
                if (folder.TestConfiguration.ConfigurationParamNames.Count() > 0)
                {
                    if (folder.TestConfiguration.ConfigurationParamNames.Contains(Constants.ADOBaseURL) &&
                        folder.TestConfiguration.ConfigurationParamNames.Contains(Constants.ADOPersonalAccessToken) &&
                        folder.TestConfiguration.ConfigurationParamNames.Contains(Constants.ADOProject))
                    {
                        SyncTestCasesTask syncTC = new SyncTestCasesTask(folder);
                        syncTC.Execute(folder, context);
                    }
                }
            }
        }
    }
}
