using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Tricentis.AddIns.Integration.ADO;
using Tricentis.TCCore.Base.Folders;
using Tricentis.TCCore.BusinessObjects.ExecutionLists;
using Tricentis.TCCore.BusinessObjects.ExecutionLists.ExecutionLogs;
using Tricentis.TCCore.BusinessObjects.Folders;
using Tricentis.TCCore.BusinessObjects.Testcases;
using Tricentis.TCCore.Persistency;
using Tricentis.TCCore.Persistency.Tasks;

namespace Tricentis.AddIns.Integration.ADOAddIn.Tasks
{
    public class SyncExecutionResultsTask : BaseADOTask
    {
        public SyncExecutionResultsTask(PersistableObject obj)
        {
            _targetObject = obj;
        }

        public override string Name => "Synchronise Test Execution Results";

        public override TaskCategory Category => new TaskCategory(Constants.AddInCategory);

        public override object Execute(PersistableObject obj, ITaskContext context)
        {
            int updateCounter = 0;
            int syncError = 0;

            StringBuilder sb = new StringBuilder();
            StringBuilder errorSB = new StringBuilder();
            sb.AppendLine("The following execution entries are not synced.");

            int skippedCounter = obj.SearchByTQL($"=>SUBPARTS:ExecutionList[{ Constants.ADOSync }==\"True\"]=>SUBPARTS:ExecutionEntry->AllReferences:TestCase[{ Constants.ADOTestCaseID }==\"\" OR { Constants.ADOPlanID }==\"\" OR { Constants.ADOSuiteID }==\"\"]->ExecutionEntries").Count;
            string skipMessage = (skippedCounter > 0 ? $" { skippedCounter } ExecutionEntries were skipped due to missing TestCase ID, Plan ID or Suite ID." : String.Empty);

            //searches for execution entries that are linked to ADO by finding a ADO Testcase ID
            context.ShowStatusInfo("Collecting Execution Entries");
            List<PersistableObject> exEntries = obj.SearchByTQL($"=>SUBPARTS:ExecutionList[{ Constants.ADOSync }==\"True\"]=>SUBPARTS:ExecutionEntry").ToList();


            StringBuilder sbMissingData = new StringBuilder();
            if (exEntries.Count > 0)
            {
                ADOSync ado = new ADOSync();

                string previousExListId = String.Empty;

                string nodePath = obj.NodePath;
                int i = 0;
                foreach (ExecutionEntry ee in exEntries)
                {
                    context.ShowProgressInfo(exEntries.Count, i++, "Syncing Execution Entry");

                    TestCase testCase = ee.TestCase.Get();
                    if (testCase.GetAttributeValue(Constants.ADOPlanID).Equals("") && testCase.GetAttributeValue(Constants.ADOSuiteID).Equals("") && testCase.GetAttributeValue(Constants.ADOTestCaseID).Equals(""))
                    {
                        continue;
                    }

                    ExecutionList el = (ExecutionList)ee.SearchByTQL("=>SUPERPART:ExecutionList").FirstOrDefault();
                    OwnedFolder folder = el.ParentFolder.Get();

                    bool adoSyncCheck = true;

                    // check all folders for ADOSync
                    while (nodePath != folder.NodePath)
                    {
                        if (bool.Parse(folder.GetAttributeValue(Constants.ADOSync).ToString()))
                        {
                            folder = folder.ParentFolder.Get();
                        }
                        else
                        {
                            adoSyncCheck = false;
                            break;
                        }
                    }

                    if (!adoSyncCheck)
                    {
                        continue;
                    }

                    TestCase tc = ee.TestCase.Get();
                    if (String.IsNullOrEmpty(tc.TestConfiguration.GetConfigurationParamValue(Constants.ADOBaseURL)) || String.IsNullOrEmpty(tc.TestConfiguration.GetConfigurationParamValue(Constants.ADOPersonalAccessToken)) || String.IsNullOrEmpty(tc.TestConfiguration.GetConfigurationParamValue(Constants.ADOProject)))
                    {
                        context.ShowErrorMessage("Error", "One or more ADO connection parameters have not been configured. Please define the appropriate TCPs and try again.");
                        return null;
                    }
                    else
                    {
                        ado = new ADOSync(tc.TestConfiguration.GetConfigurationParamValue(Constants.ADOBaseURL), tc.TestConfiguration.GetConfigurationParamValue(Constants.ADOPersonalAccessToken), tc.TestConfiguration.GetConfigurationParamValue(Constants.ADOProject));
                    }

                    string adoPlanId = tc.GetPropertyValue(Constants.ADOPlanID);
                    string adoSuiteId = tc.GetPropertyValue(Constants.ADOSuiteID);
                    string adoTestCaseId = tc.GetPropertyValue(Constants.ADOTestCaseID);

                    if (String.IsNullOrEmpty(adoPlanId) || String.IsNullOrEmpty(adoSuiteId) || String.IsNullOrEmpty(adoTestCaseId))
                    {
                        sbMissingData.AppendLine($"TestCase ID { adoTestCaseId } is missing additional information.{ (String.IsNullOrEmpty(adoPlanId) ? "ADO PlanID is missing." : "") } { (String.IsNullOrEmpty(adoSuiteId) ? "ADO SuiteID is missing." : "")  }");
                        skippedCounter++;
                    }

                    // All execution entries without a ADO_RunID value
                    List<ExecutionTestCaseLog> executionTestCaseLogs = ee.ExecutionLogs.Where(t => t.Result != ExecutionResult.NoResult && t.GetPropertyValue(Constants.ADORunID) == "").ToList();
                    foreach (ExecutionTestCaseLog executionTestCaseLog in executionTestCaseLogs)
                    {
                        string resultsLog = GetTestStepLogs(executionTestCaseLog);
                        string executionResult = executionTestCaseLog.Result.ToString();
                        string endTime = executionTestCaseLog.EndTime.ToString();
                        string startTime = executionTestCaseLog.StartTime.ToString();
                        int testRunId;
                        int testResultId;

                        try
                        {
                            context.ShowProgressInfo(exEntries.Count, i++, "Creating Test Runs In ADO");
                            (testRunId, testResultId) = ado.UpdateExecutionResult(Int32.Parse(adoPlanId), Int32.Parse(adoSuiteId), Int32.Parse(adoTestCaseId), startTime, endTime, executionResult, resultsLog);

                            executionTestCaseLog.SetAttributeValueByString(Constants.ADORunID, testRunId.ToString());

                            updateCounter++;
                        }
                        catch 
                        {
                            errorSB.AppendLine("Could not create Test Run, check the TCP.");
                            sb.AppendLine("- " + ee.Name);
                            syncError++;
                            break;
                        }
                    }

                    context.ShowProgressInfo(exEntries.Count, i, "Test Run Created");
                }
            }

            if (updateCounter > 0 && syncError == 0)
            {
                context.ShowMessageBox("Sync Execution Results", $"{ updateCounter } test execution results successfully synchronised. { (skippedCounter > 0 ? skipMessage : String.Empty) }");
            }
            else if (syncError > 0 && !errorSB.Equals(null))
            {
                context.ShowMessageBox("Sync Execution Results", $"{ updateCounter } test execution results successfully synchronised.\r\n" + errorSB.ToString() + sb.ToString());
            }
            else
            {
                context.ShowMessageBox("Sync Execution Results", $"There were no results to synchronise. { (skippedCounter > 0 ? skipMessage : String.Empty) }");
            }

            return null;
        }

        public string GetTestStepLogs(ExecutionTestCaseLog executionTestCaseLog)
        {
            StringBuilder resultLog = new StringBuilder();

            List<ExecutionTestStepLog> executionTestStepLogs = executionTestCaseLog.TestStepLogsInRightOrder.ToList();
            foreach (ExecutionTestStepLog executionTestStepLog in executionTestStepLogs)
            {
                if (executionTestStepLog.Result == ExecutionResult.Failed)
                {
                    if (resultLog.Length > 0)
                    {
                        resultLog.AppendLine();
                    }

                    resultLog.AppendLine($"TestStep: { executionTestStepLog.Name }");
                    resultLog.AppendLine($"Result: { executionTestStepLog.Result }");

                    List<ExecutionTestStepValueLog> executionTestStepValueLogs = executionTestStepLog.TestStepValueLogs.Where(x => x.Result == ExecutionResult.Failed).ToList();
                    foreach (ExecutionTestStepValueLog executionTestStepValueLog in executionTestStepValueLogs)
                    {
                        resultLog.AppendLine($"-TestStepValue Name: { executionTestStepValueLog.Name }");
                        resultLog.AppendLine($"-TestStepValue Action: { executionTestStepValueLog.ActionMode }");
                        resultLog.AppendLine($"-TestStepValue LogInfo: { executionTestStepValueLog.LogInfo.Replace("\r\n", "\r\n                        ") }");
                    }
                }
            }

            return resultLog.ToString();
        }

    }
}

