using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Tricentis.AddIns.Integration.ADO;
using Tricentis.AddIns.Integration.ADO.Objects;
using Tricentis.TCCore.Base.TestConfigurations;
using Tricentis.TCCore.BusinessObjects.ExecutionLists;
using Tricentis.TCCore.BusinessObjects.ExecutionLists.ExecutionLogs;
using Tricentis.TCCore.BusinessObjects.Issues;
using Tricentis.TCCore.BusinessObjects.Testcases;
using Tricentis.TCCore.Persistency;
using Tricentis.TCCore.Persistency.Tasks;

namespace Tricentis.AddIns.Integration.ADOAddIn.Tasks
{
    public class UpdateSynchronisedDefectsTask : BaseADOTask
    {
        public UpdateSynchronisedDefectsTask(PersistableObject obj)
        {
            _targetObject = obj;
        }

        public override string Name => "Update Synchronised Defects";

        public override TaskCategory Category => new TaskCategory(Constants.AddInCategory);

        public override object Execute(PersistableObject obj, ITaskContext context)
        {
            context.ShowStatusInfo("Gathering Issues");
            List<Issue> syncedIssues = obj.SearchByTQL("=>SUBPARTS:Issue[ID!=\"\"]").Select(x => (Issue)x).ToList();

            ADOSync ado = null;
            bool connectionErr = false;
            int syncedCount = 0;
            string prevBaseAddress = "", prevPersonalAccessToken = "", prevProject = "";
            List<Issue> unsyncedResults = new List<Issue>();
            int j = 0;
            //Update Tosca Issue based on search in ADO/TFS by ID
            foreach (Issue i in syncedIssues)
            {
                context.ShowProgressInfo(syncedIssues.Count, j++, "Checking Issue Link");
                IssueLink il = (IssueLink)i.SearchByTQL("->AllReferences:IssueLink").FirstOrDefault();
                if (il != null)
                {
                    ExecutionTestCaseLog log = il.ExecutionTestCaseLog.Get();
                    TestCase tc = log.ExecutionEntry.Get().TestCase.Get();

                    (string add, string token, string proj) = GetConfiguration(tc);

                    // If issue's execution list is not configured, get out
                    if (String.IsNullOrEmpty(add) || string.IsNullOrEmpty(token) || string.IsNullOrEmpty(proj))
                    {
                        unsyncedResults.Add(i);
                        continue;
                    }

                    // Check if current configuration is same as previous issue's execution list
                    // Reinstantiate a new connection if so
                    if (add != prevBaseAddress || token != prevPersonalAccessToken || proj != prevProject)
                    {
                        ado = new ADOSync(add, token, proj);
                        prevBaseAddress = add;
                        prevPersonalAccessToken = token;
                        prevProject = proj;
                    }

                    context.ShowProgressInfo(syncedIssues.Count, j, "Retrieving Bug Info");
                    ADODefect def = null;
                    try
                    {
                        (def, connectionErr) = ado.RetrieveBug(int.Parse(i.GetAttributeValue("ID").ToString()));
                    }
                    catch
                    {
                        connectionErr = true;
                        unsyncedResults.Add(i);
                    }

                    if (def == null)
                    {
                        unsyncedResults.Add(i);
                    }
                    else
                    {
                        context.ShowProgressInfo(syncedIssues.Count, j, "Updating Issue");
                        i.Name = def.Name;
                        i.Severity = def.Severity;
                        i.State = def.State;
                        i.Priority = def.Priority;

                        syncedCount++;
                    }
                }
                else
                {
                    unsyncedResults.Add(i);
                }
            }

            StringBuilder sb = new StringBuilder();
            if (syncedCount > 0)
            {
                sb.AppendLine($"{ syncedCount } issues synced.");
                sb.AppendLine();
            }
            else
            {
                sb.AppendLine("No issues were updated.");
            }

            if (unsyncedResults.Count > 0)
            {
                if (connectionErr.Equals(true))
                {
                    sb.AppendLine("The following issues could not be synced, the personal access token has expired:");
                }
                else if (connectionErr.Equals(false))
                {
                    sb.AppendLine("The following issues could not be synced: ");
                }
                foreach (String name in unsyncedResults.Select(i => i.Name))
                {
                    sb.AppendLine("- " + name);
                }
            }

            context.ShowMessageBox("Sync Status", sb.ToString());

            return null;
        }
    }
}
