using System;
using System.Collections.Generic;
using System.Linq;
using Tricentis.AddIns.Integration.ADO;
using Tricentis.AddIns.Integration.ADOAddIn.Tasks;
using Tricentis.TCCore.Base.TestConfigurations;
using Tricentis.TCCore.BusinessObjects;
using Tricentis.TCCore.BusinessObjects.Folders;
using Tricentis.TCCore.BusinessObjects.Testcases;
using Tricentis.TCCore.Persistency;
using Tricentis.TCCore.Persistency.AddInManager;

namespace ToscaAddIn
{
    class ToscaAddInTaskHandler : TaskInterceptor
    {
        public ToscaAddInTaskHandler(PersistableObject obj)
        {
        }

        public override void GetTasks(PersistableObject obj, List<Task> tasks)
        {
            if (obj is TCProject)
            {
                tasks.Add(new PrepareWorkspaceTask());
            }

            if (obj is TCFolder)
            {
                //checks if the folder has the attribute ADOSync
                if (!obj.HasAttribute(Constants.ADOSync))
                {
                    return;
                }

                bool syncFolder = false;
                if (!bool.TryParse(obj.GetAttributeValue(Constants.ADOSync).ToString(), out syncFolder) || !syncFolder)
                {   // Get out if sync is not true
                    return;
                }

                IObjectWithConfiguration configObj = obj as IObjectWithConfiguration;
                string baseAddress = configObj.TestConfiguration.GetConfigurationParamValue(Constants.ADOBaseURL);
                string personalToken = configObj.TestConfiguration.GetConfigurationParamValue(Constants.ADOPersonalAccessToken);
                string project = configObj.TestConfiguration.GetConfigurationParamValue(Constants.ADOProject);

                TCFolder folder = (TCFolder)obj;
                
                String[] content = folder.PossibleContent.Split(';');
                List<PersistableObject> childFolders = new List<PersistableObject>();
                childFolders = folder.SearchByTQL("=>SUBPARTS:TCFolder");
                int count = 0;

                foreach (TCFolder fld in childFolders)
                {
                    if (fld.TestConfiguration.ConfigurationParamNames.Contains(Constants.ADOBaseURL)
                    && fld.TestConfiguration.ConfigurationParamNames.Contains(Constants.ADOPersonalAccessToken)
                    && fld.TestConfiguration.ConfigurationParamNames.Contains(Constants.ADOProject))
                    {
                        count++;
                    }

                    if (count >= 2)
                    {   // Could be multiple projects, we are just concerned with knowing if there are multiple or single projects
                        break;
                    }
                }

                /* checks if 
                 * object property "PossibleContent" == Test Case
                 * TCP are not empty
                 * Property Suite and Plan ID's are "N/A"
                */
                if (content.Contains(Constants.PossibleContent_TestCase) && 
                    !String.IsNullOrEmpty(baseAddress) && 
                    !String.IsNullOrEmpty(personalToken) && 
                    !String.IsNullOrEmpty(project) &&
                    folder.GetAttributeValue(Constants.ADOPlanID).Equals(Constants.NA) &&
                    folder.GetAttributeValue(Constants.ADOSuiteID).Equals(Constants.NA))
                {
                    tasks.Add(new SyncTestCasesTask(obj));
                }
                else if (content.Contains(Constants.PossibleContent_TestCase) && count > 1)
                {
                    tasks.Add(new SyncProjects(obj));
                }
                else if (content.Contains(Constants.PossibleContent_ExecutionList))
                {
                    tasks.Add(new SyncExecutionResultsTask(obj));
                }
                else if (content.Contains(Constants.PossibleContent_Issue))
                {
                    tasks.Add(new SyncNewDefectsTask(obj));
                    tasks.Add(new UpdateSynchronisedDefectsTask(obj));
                }
            }
        }
    }
}
