using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Tricentis.AddIns.Integration.ADO;
using Tricentis.AddIns.Integration.ADO.Objects;
using Tricentis.TCCore.Base.TestConfigurations;
using Tricentis.TCCore.BusinessObjects.ExecutionLists;
using Tricentis.TCCore.BusinessObjects.ExecutionLists.ExecutionLogs;
using Tricentis.TCCore.BusinessObjects.Folders;
using Tricentis.TCCore.BusinessObjects.Issues;
using Tricentis.TCCore.BusinessObjects.Testcases;
using Tricentis.TCCore.Persistency;
using Tricentis.TCCore.Persistency.Tasks;

namespace Tricentis.AddIns.Integration.ADOAddIn.Tasks
{
    public class SyncNewDefectsTask : BaseADOTask
    {
        public SyncNewDefectsTask(PersistableObject obj)
        {
            _targetObject = obj;
        }

        public override string Name => "Synchronise New Defects";

        public override TaskCategory Category => new TaskCategory(Constants.AddInCategory);

        public override object Execute(PersistableObject obj, ITaskContext context)
        {
            context.ShowStatusInfo("Gathering Issues to Sync");

            // Searches for Issues in the workspace that have no ID, which means it is not synced
            List<Issue> unsyncedIssues = obj.SearchByTQL("=>SUBPARTS:Issue[ID==\"\"]").Select(x => (Issue)x).ToList();
            List<ADODefect> issues = new List<ADODefect>();

            string prevBaseAddress = "", prevPersonalAccessToken = "", prevProject = "";

            ADOSync ado = null;
            int syncedCount = 0;

            StringBuilder sb2 = new StringBuilder();
            bool connectionErr = false;
            bool valueChk = false;
            bool execChk = false;
            bool configChk = false;
            int j = 0;

            //Create Bug in ADO/TFS if it does not have an ID
            foreach (Issue i in unsyncedIssues)
            {
                context.ShowProgressInfo(unsyncedIssues.Count, j++, "Checking Issue Link");
                //gets the issuelink of issues that are unsynced
                IssueLink il = (IssueLink)i.SearchByTQL("->AllReferences:IssueLink").FirstOrDefault();
                if (il != null)
                {
                    // gets execution list that the issue is linked to
                    ExecutionTestCaseLog log = il.ExecutionTestCaseLog.Get();
                    TestCase tc = log.ExecutionEntry.Get().TestCase.Get();

                    (string add, string token, string proj) = GetConfiguration(tc);

                    // If issue's execution list is not configured, get out
                    if (String.IsNullOrEmpty(add) || string.IsNullOrEmpty(token) || string.IsNullOrEmpty(proj))
                    {
                        configChk = true;
                        sb2.AppendLine($"- {i.Name}");
                        continue;
                    }

                    if (log.HasProperty(Constants.ADORunID) && !String.IsNullOrEmpty(log.GetPropertyValue(Constants.ADORunID)))
                    {
                        // Check if current configuration is same as previous issue's execution list
                        // Reinstantiate a new connection if so
                        if (add != prevBaseAddress || token != prevPersonalAccessToken || proj != prevProject)
                        {
                            ado = new ADOSync(add, token, proj);
                            prevBaseAddress = add;
                            prevPersonalAccessToken = token;
                            prevProject = proj;
                        }

                        int defaultSeverity = 3;
                        int defaultPriority = 2;
                        if (i.Severity.Equals(0))
                        {
                            i.Severity = defaultSeverity;
                        }
                        if (i.Priority.Equals(0))
                        {
                            i.Priority = defaultPriority;
                        }

                        ADODefect defect = CreateDefectObject(i);
                        context.ShowProgressInfo(unsyncedIssues.Count, j, "Checking Priority and Severity");
                        if (defect.Priority >= 1 && defect.Priority <= 4 && defect.Severity >= 1 && defect.Severity <= 4)
                        {
                            context.ShowProgressInfo(unsyncedIssues.Count, j, "Creating Bug in ADO");
                            try
                            {
                                (defect.Identification, defect.State) = ado.CreateBug(defect);
                            }
                            catch
                            {
                                sb2.AppendLine($"- {defect.Name}");
                                connectionErr = true;
                            }
                            if (defect.Identification == -1 || defect.State == null)
                            {
                            }
                            else
                            {
                                syncedCount++;
                                i.ID = defect.Identification.ToString();
                                i.State = defect.State;
                            }
                        }
                        else
                        {
                            sb2.AppendLine($"- {defect.Name}");
                            valueChk = true;
                        }
                    }
                    else
                    {
                        execChk = true;
                        sb2.AppendLine($"- {i.Name}");
                    }
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
                sb.AppendLine("There were no issues synced.");
            }

            if (configChk.Equals(true))
            {
                sb.AppendLine("The execution results linked to these issues do not have the correct TCP:");
                sb.AppendLine(sb2.ToString());
            }
            else if (execChk.Equals(true))
            {
                sb.AppendLine("The execution results linked to these issues are not synced:");
                sb.AppendLine(sb2.ToString());
            }
            else if (valueChk.Equals(true))
            {
                sb.AppendLine("The following issues have invalid priority or severity:");
                sb.AppendLine(sb2.ToString());
            }
            else if (connectionErr.Equals(true))
            {
                sb.AppendLine("The connection to ADO was denied, the personal access token for the following has expired:");
                sb.AppendLine(sb2.ToString());
            }
            context.ShowMessageBox("Sync Status", sb.ToString());

            return null;
        }

        public (ADODefect, string) GetPropertyMappings(TCFolder root, ADODefect dft)
        {
            string attValue = (string)root.GetAttributeValue(Constants.ADOCustomProperties);
            List<string> maps = attValue.Split('|').ToList();
            List<string> validMaps = new List<string>();
            string errorMessage = null;

            foreach (string map in maps)
            {
                CustomField customField = new CustomField();
                List<string> techMap = map.Split('=').ToList();
                foreach (string reference in techMap)
                {
                    if (reference.Count(x => x == '.') == 2)
                    {
                        if (reference.ToLower().Contains("ado.b"))
                        {
                            customField.ADOName = reference.Split('.').ElementAt(2);
                        }
                        else if (reference.ToLower().Contains("t.i"))
                        {
                            customField.Name = reference.Split('.').ElementAt(2);
                            dft.CustomFields.Add(customField);
                        }
                    }
                    else
                    {
                        errorMessage = "The format for one or more of the mappings is incorrect.";
                    }
                }
            }

            return (dft, errorMessage);
        }

        private ADODefect CreateDefectObject(Issue issue)
        {
            List<DefectLinks> defLinks = new List<DefectLinks>();
            List<PersistableObject> links = issue.SearchByTQL("->AllReferences:IssueLink");

            foreach (PersistableObject link in links)
            {
                IssueLink il = (IssueLink)link;

                ExecutionTestCaseLog log = il.ExecutionTestCaseLog.Get();

                int tcId = -1;
                int runId = -1;

                if (!int.TryParse(log.ExecutedTestCase.Get().GetAttributeValue(Constants.ADOTestCaseID).ToString(), out tcId)
                    || !int.TryParse(log.GetAttributeValue(Constants.ADORunID).ToString(), out runId))
                {
                    continue;
                }

                defLinks.Add(new DefectLinks(tcId, runId));
            }

            return new ADODefect(-1, issue.Name, issue.Severity, issue.Priority, defLinks, issue.UniqueId);
        }

    }
}