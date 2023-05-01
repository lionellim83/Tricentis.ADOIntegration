using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Tricentis.AddIns.Integration.ADO;
using Tricentis.AddIns.Integration.ADO.Enums;
using Tricentis.AddIns.Integration.ADO.Objects;
using Tricentis.TCCore.Base;
using Tricentis.TCCore.BusinessObjects.Folders;
using Tricentis.TCCore.BusinessObjects.Tasks;
using Tricentis.TCCore.BusinessObjects.Testcases;
using Tricentis.TCCore.Persistency;
using Tricentis.TCCore.Persistency.Tasks;
using core = Tricentis.TCCore.BusinessObjects.Testcases;
using xdef = Tricentis.TCAddIns.XDefinitions.Testcases;


namespace Tricentis.AddIns.Integration.ADOAddIn.Tasks
{
    public class SyncTestCasesTask : BaseADOTask
    {
        public SyncTestCasesTask(PersistableObject obj)
        {
            _targetObject = obj;
        }
        public override string Name => "Synchronise Test Cases";

        public override TaskCategory Category => new TaskCategory(Constants.AddInCategory);

        public override object Execute(PersistableObject obj, ITaskContext context)
        {
            DateTime time = DateTime.Now;

            if (String.IsNullOrEmpty(BaseAddress) || String.IsNullOrEmpty(PersonalAccessToken) || String.IsNullOrEmpty(Project))
            {
                context.ShowErrorMessage($"Error in {obj.DisplayedName}", "One or more ADO connection parameters have not been configured. Please define the appropriate TCPs and try again.");
                return null;
            }

            // gets Test cases from ADO
            List<ADOTestCase> testCases = new List<ADOTestCase>();
            string errorStatus = null;
            bool valid = false;

            ADOSync adoSync = new ADOSync(base.BaseAddress, base.PersonalAccessToken, base.Project);

            (testCases, errorStatus, valid) = adoSync.GetTestCases();

            if (errorStatus != "OK" && valid == false)
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendLine($"The connection to {Project} failed: {errorStatus}");
                sb.AppendLine($"Please check that the test configuration values are valid.");

                context.ShowMessageBox($"Connection To {obj.DisplayedName}", sb.ToString());

                return null;
            }

            TCFolder root = (TCFolder)obj;
            if (!RetrieveCustomProperties(context, ref adoSync, root, ref testCases))
            {
                return null;
            }

            //Sets the testcases in tosca to deleted status
            context.ShowStatusInfo("Reinitialising test case status");
            SetTestCaseStatusDeleted(root);

            SyncADOData(root, testCases, time, context);

            context.ShowStatusInfo("Reporting deleted Test Cases");
            ReportDeletedTestCases(root, context);

            // Syncs the test steps in ADO
            context.ShowStatusInfo("Collecting Test Steps for synchronisation");
            testCases = GetTestStepDetails(root, testCases);

            if (testCases.Count != 0)
            {
                int total = testCases.Count();
                int i = 0;

                foreach (ADOTestCase tc in testCases)
                {
                    context.ShowProgressInfo(total, i++, "Sending Test Steps to ADO");
                    if (tc.ToscaTestSteps.Count > 0)
                    {
                        adoSync.SetTestSteps(tc);
                    }
                }
            }

            context.ShowStatusInfo("Synchronisation of Test Cases completed");

            return null;
        }

        private bool RetrieveCustomProperties(ITaskContext context, ref ADOSync adoSync, TCFolder root, ref List<ADOTestCase> testCases)
        {
            string errorMessage = null;
            bool customFieldSuccess = false;
            (customFieldSuccess, errorMessage) = GetPropertyMappings(root, ref testCases);

            if (!customFieldSuccess)
            {
                context.ShowErrorMessage("Custom Field Mapping", errorMessage);
            }

            foreach (ADOTestCase tc in testCases)
            {
                List<String> tcCustomFieldValues = adoSync.GetCustomProperties(tc.TestCaseID, tc.CustomFieldsTC.Select(n => n.ADOName).ToList(), "Test Case");
                List<String> tsCustomFieldValues = adoSync.GetCustomProperties(tc.SuiteID, tc.CustomFieldsTS.Select(n => n.ADOName).ToList(), "Test Suite");
                List<String> tpCustomFieldValues = adoSync.GetCustomProperties(tc.PlanID, tc.CustomFieldsTP.Select(n => n.ADOName).ToList(), "Test Plan");

                if (tc.CustomFieldsTC.Count() != tcCustomFieldValues.Count())
                {
                    context.ShowErrorMessage("Custom Field Error", "One or more ADO Test Case fields specified is invalid");
                    return false;
                }

                if (tc.CustomFieldsTS.Count() != tsCustomFieldValues.Count())
                {
                    context.ShowErrorMessage("Custom Field Error", "One or more ADO Test Suite fields specified is invalid");
                    return false;
                }

                if (tc.CustomFieldsTP.Count() != tpCustomFieldValues.Count())
                {
                    context.ShowErrorMessage("Custom Field Error", "One or more ADO Test Plan fields specified is invalid");
                    return false;
                }

                for (int i = 0; i < tc.CustomFieldsTC.Count(); i++)
                {
                    tc.CustomFieldsTC[i].Value = tcCustomFieldValues[i];
                }

                for (int i = 0; i < tc.CustomFieldsTS.Count(); i++)
                {
                    tc.CustomFieldsTS[i].Value = tsCustomFieldValues[i];
                }

                for (int i = 0; i < tc.CustomFieldsTP.Count(); i++)
                {
                    tc.CustomFieldsTP[i].Value = tpCustomFieldValues[i];
                }
            }

            return true;
        }

        public (bool, string) GetPropertyMappings(TCFolder root, ref List<ADOTestCase> tcs)
        {
            string mappingRuleProperty = (string)root.GetAttributeValue(Constants.ADOCustomProperties);

            List<string> mappingRules = mappingRuleProperty.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries).ToList();

            if (mappingRules.Count > 0)
            {
                List<CustomField> tcCustomFields = new List<CustomField>();
                List<CustomField> tpCustomFields = new List<CustomField>();
                List<CustomField> tsCustomFields = new List<CustomField>();

                // mappingRule: ADO.TC.FieldName=T.TC.PropertyName
                // mappingRule: T.TC.PropertyName=ADO.TC.FieldName
                // mappingRule: ADO.TP.FieldNMame=T.TC.PropertyName
                foreach (string rule in mappingRules)
                {
                    #region General notation validation

                    // Check for 4 "." and ADO.__.___ as well as T.__.___ exists

                    if (rule.Count(c => c == '.') != 4 || rule.Count(x => x == '=') != 1)
                    {
                        return (false, $"'{ rule }' is not a valid mapping of custom fields");
                    }

                    #endregion

                    #region Identifying ADO and Tosca fields
                    string adoComponent = string.Empty;
                    string toscaComponent = string.Empty;
                    string firstHalf = rule.Substring(0, rule.IndexOf("="));

                    if (firstHalf.ToLower().Contains("ado"))
                    {
                        adoComponent = firstHalf;
                    }
                    else
                    {
                        toscaComponent = firstHalf;
                    }

                    if (adoComponent.Equals(string.Empty))
                    {
                        adoComponent = rule.Substring(rule.IndexOf("=") + 1);
                    }
                    else if (toscaComponent.Equals(string.Empty))
                    {
                        toscaComponent = rule.Substring(rule.IndexOf("=") + 1);
                    }

                    if (toscaComponent.Contains("\r\n"))
                    {
                        toscaComponent = toscaComponent.Replace("\r\n", "");
                    }
                    else if (adoComponent.Contains("\r\n"))
                    {
                        adoComponent = adoComponent.Replace("\r\n", "");
                    }
                    if (toscaComponent.Substring(0, toscaComponent.IndexOf(".")).ToLower() != "t" || adoComponent.Substring(0, adoComponent.IndexOf(".")).ToLower() != "ado")
                    {
                        return (false, $"'{ rule }' has an incorrect technology part");
                    }

                    CustomField customField = new CustomField(toscaComponent.Substring(toscaComponent.LastIndexOf(".") + 1), adoComponent.Substring(adoComponent.LastIndexOf(".") + 1));

                    #endregion

                    #region Check if Tosca property exists and add to collection if valid

                    string toscaObjectType = toscaComponent.Substring(toscaComponent.IndexOf(".") + 1);
                    string adoObjectType = adoComponent.Substring(adoComponent.IndexOf(".") + 1);
                    toscaObjectType = toscaObjectType.Substring(0, toscaObjectType.IndexOf("."));
                    adoObjectType = adoObjectType.Substring(0, adoObjectType.IndexOf("."));

                    if (!toscaObjectType.Equals(adoObjectType))
                    {
                        return (false, $"Object types do not match");
                    }

                    TCBaseProject project = Workspace.ActiveWorkspace.Project;

                    switch (toscaObjectType.ToLower())
                    {
                        case "tc":
                            TCObjectPropertiesDefinition tcPropDef = project.GetOrCreatePropertiesDefinition("TestCase");
                            if (tcPropDef.GetTypedProperties().FirstOrDefault(p => p.Name.ToLower() == customField.Name.ToLower()) == null)
                            {
                                return (false, $"Property name '{ customField.Name }' does not exist");
                            }
                            tcCustomFields.Add(customField);
                            break;
                        case "ts":
                            TCObjectPropertiesDefinition tsPropDef = project.GetOrCreatePropertiesDefinition("Folder");
                            if (tsPropDef.GetTypedProperties().FirstOrDefault(p => p.Name.ToLower() == customField.Name.ToLower()) == null)
                            {
                                return (false, $"Property name '{ customField.Name }' does not exist");
                            }
                            tsCustomFields.Add(customField);
                            break;
                        case "tp":
                            TCObjectPropertiesDefinition tpPropDef = project.GetOrCreatePropertiesDefinition("Folder");
                            if (tpPropDef.GetTypedProperties().FirstOrDefault(p => p.Name.ToLower() == customField.Name.ToLower()) == null)
                            {
                                return (false, $"Property name '{ customField.Name }' does not exist");
                            }
                            tpCustomFields.Add(customField);
                            break;
                        default:
                            return (false, $"Property name '{ customField.Name }' is not a support object type for custom fields");    // exit from function if there is an error
                    }

                    #endregion
                }

                //Add the custom field mapping to every Test Case object
                foreach (ADOTestCase tc in tcs)
                {
                    tc.CustomFieldsTC = tcCustomFields.ConvertAll(cf => cf.Copy());
                    tc.CustomFieldsTS = tsCustomFields.ConvertAll(cf => cf.Copy());
                    tc.CustomFieldsTP = tpCustomFields.ConvertAll(cf => cf.Copy());
                }
            }

            return (true, String.Empty);
        }

        public List<ADOTestCase> GetTestCasesToSync(TCFolder root)
        {
            List<PersistableObject> testCases = root.SearchByTQL($"=>SUBPARTS: TestCase[{Constants.ADOStatus} == \"Unsynced\"]");
            List<ADOTestCase> tcList = new List<ADOTestCase>();

            foreach (TestCase tc in testCases)
            {
                // creates a testcase object
                ADOTestCase adoTC = new ADOTestCase();

                // fills the properties with the values for that testcase
                adoTC.PlanID = tc.GetAttributeValue(Constants.ADOPlanID).ToString();
                adoTC.SuiteID = tc.GetAttributeValue(Constants.ADOSuiteID).ToString();
                adoTC.TestCaseID = tc.GetAttributeValue(Constants.ADOTestCaseID).ToString();

                adoTC.PlanName = tc.GetAllDirectSuperParts().First().GetAllDirectSuperParts().First().DisplayedName;
                adoTC.SuiteName = tc.GetAllDirectSuperParts().First().DisplayedName;
                adoTC.TestCaseName = tc.DisplayedName;

                tcList.Add(adoTC);
            }

            return tcList;
        }

        public List<ADOTestCase> GetTestStepDetails(TCFolder folder, List<ADOTestCase> testCaseList)
        {
            // gets all the testcases in the folder
            List<PersistableObject> testCases = folder.SearchByTQL("=>SUBPARTS:TestCase");

            if (testCases.Count != 0)
            {
                foreach (core.TestCase tc in testCases)
                {
                    if (tc.TestConfiguration.GetConfigurationParamValue(Constants.ADOTestStepSwitch) == null)
                    {
                        continue;
                    }
                    else
                    {
                        // Check Config for TestStep Switch
                        if (tc.TestConfiguration.GetConfigurationParamValue(Constants.ADOTestStepSwitch).ToLower() == "true")
                        {
                            // creates a testcase object
                            ADOTestCase adoTC = new ADOTestCase();

                            // fills the properties with the values for that testcase
                            adoTC.PlanID = tc.GetAttributeValue(Constants.ADOPlanID).ToString();
                            adoTC.SuiteID = tc.GetAttributeValue(Constants.ADOSuiteID).ToString();
                            adoTC.TestCaseID = tc.GetAttributeValue(Constants.ADOTestCaseID).ToString();
                            adoTC.PlanName = tc.GetAllDirectSuperParts().First().GetAllDirectSuperParts().First().DisplayedName;
                            adoTC.SuiteName = tc.GetAllDirectSuperParts().First().DisplayedName;
                            adoTC.TestCaseName = tc.DisplayedName;

                            // gets all the test steps including xteststeps and teststeps(Classic)
                            List<core.TestCaseItem> tci = tc.GetAllTestCaseItems().Where(i => i.GetType() == typeof(xdef.XTestStep) || i.GetType() == typeof(core.TestStep)).ToList();

                            foreach (core.TestCaseItem item in tci)
                            {
                                // creates a test step object for the testcase
                                ToscaTestStep ts = new ToscaTestStep();
                                ts.Name = item.Name;

                                // if the item is xteststep
                                if (item.GetType() == typeof(xdef.XTestStep))
                                {
                                    xdef.XTestStep t = (xdef.XTestStep)item;

                                    foreach (xdef.XTestStepValue xtsv in t.TestStepValuesInRightOrder)
                                    {
                                        ToscaTestAction ta = new ToscaTestAction(xtsv.Name, xtsv.Value, ConvertToADOActionModeX(xtsv.ActionModeToUse));
                                        ts.TestActions.Add(ta);
                                    }
                                }
                                // if the item is teststep (classic)
                                else if (item.GetType() == typeof(core.TestStep))
                                {
                                    core.TestStep t = (core.TestStep)item;

                                    foreach (core.TestStepValue ctsv in t.GetTestStepValuesOrdered())
                                    {
                                        ToscaTestAction ta = new ToscaTestAction(ctsv.Name, ctsv.Value, ConvertToADOActionMode(ctsv.ActionMode));
                                        ts.TestActions.Add(ta);
                                    }
                                }

                                adoTC.ToscaTestSteps.Add(ts);
                            }
                            testCaseList.Add(adoTC);
                        }

                    }
                }
            }

            return testCaseList;
        }

        private ActionMode ConvertToADOActionModeX(xdef.XTestStepActionMode actionMode)
        {
            switch (actionMode)
            {
                case xdef.XTestStepActionMode.Select:
                    return ActionMode.Select;
                case xdef.XTestStepActionMode.Constraint:
                    return ActionMode.Constraints;
                case xdef.XTestStepActionMode.Insert:
                    return ActionMode.Insert;
                case xdef.XTestStepActionMode.Input:
                    return ActionMode.Input;
                case xdef.XTestStepActionMode.Verify:
                    return ActionMode.Verify;
                case xdef.XTestStepActionMode.WaitOn:
                    return ActionMode.WaitOn;
                case xdef.XTestStepActionMode.Buffer:
                    return ActionMode.Buffer;
                case xdef.XTestStepActionMode.Output:
                    return ActionMode.Output;
                default:
                    break;
            }

            throw new Exception($"ActionMode [{actionMode.ToString()}] is not supported");
        }

        private ActionMode ConvertToADOActionMode(core.TestStepActionMode actionMode)
        {
            switch (actionMode)
            {
                case core.TestStepActionMode.DoNothing:
                    return ActionMode.DoNothing;
                case core.TestStepActionMode.Input:
                    return ActionMode.Input;
                case core.TestStepActionMode.Verify:
                    return ActionMode.Verify;
                case core.TestStepActionMode.WaitOn:
                    return ActionMode.WaitOn;
                case core.TestStepActionMode.External:
                    return ActionMode.External;
                case core.TestStepActionMode.Buffer:
                    return ActionMode.Buffer;
                case core.TestStepActionMode.Output:
                    return ActionMode.Output;
                case core.TestStepActionMode.Expert:
                    return ActionMode.Expert;
                default:
                    break;
            }

            throw new Exception($"ActionMode [{actionMode.ToString()}] is not supported");
        }

        public void ReportDeletedTestCases(TCFolder root, ITaskContext task)
        {
            List<PersistableObject> deletedTestCases = root.SearchByTQL($"=>SUBPARTS: TestCase[{Constants.ADOStatus} == \"Deleted\"]");

            if (deletedTestCases.Count > 0)
            {
                StringBuilder sb = new StringBuilder();

                sb.AppendLine($"The following {deletedTestCases.Count} testcases do not exist in ADO: ");
                foreach (TestCase tc in deletedTestCases)
                {
                    sb.AppendLine("- " + tc.Name);
                }

                task.ShowMessageBox("Attention", sb.ToString());
            }
        }

        public void SetTestCaseStatusDeleted(PersistableObject folder)
        {
            foreach (TestCase tc in folder.SearchByTQL($"=>SUBPARTS: TestCase[{Constants.ADOStatus} == \"Exists\"]"))
            {
                tc.SetAttributeValueByString(Constants.ADOStatus, "Deleted");
            }
        }

        public void SyncADOData(TCFolder rootFolder, List<ADOTestCase> ado_testCases, DateTime time, ITaskContext task = null)
        {
            int total = ado_testCases.Count();
            int i = 0;

            bool skipNameChange = false;
            foreach (ADOTestCase ado_Instance in ado_testCases)
            {
                task.ShowProgressInfo(total, i++, "Synchronising Test Cases");
                //check if test plan folder exist inside ADO
                TCFolder testPlanFolder = CreateFolder(rootFolder, ado_Instance, "Test Plan");

                //check if test suite folder exists
                TCFolder testSuiteFolder = CreateFolder(testPlanFolder, ado_Instance, "Test Suite");

                //check if testcase exists
                skipNameChange = CreateUpdateTestCase(testSuiteFolder, ado_Instance, task, true, skipNameChange, time);
            }
        }

        public TCFolder CreateFolder(TCFolder parentFolder, ADOTestCase testCase, string folderType)
        {
            if (folderType == "Test Plan")
            {
                int count = parentFolder.SearchByTQL($"=>SUBPARTS[{Constants.ADOPlanID} == \"{ testCase.PlanID }\" AND { Constants.ADOSuiteID} == \"{ Constants.NA }\"]").Count();
                if (count == 0)
                {
                    TCFolder childFolder = parentFolder.GetOrCreateTestCaseSubFolder(testCase.PlanName);
                    childFolder.SetAttributeValueByString(Constants.ADOPlanID, testCase.PlanID);
                    childFolder.SetAttributeValueByString(Constants.ADOSuiteID, Constants.NA);

                    foreach (CustomField cf in testCase.CustomFieldsTP)
                    {
                        childFolder.SetAttributeValueByString(cf.Name, cf.Value);
                    }

                    return childFolder;
                }
                else
                {
                    TCFolder childFolder = parentFolder.GetOrCreateTestCaseSubFolder(testCase.PlanName);

                    foreach (CustomField cf in testCase.CustomFieldsTP)
                    {
                        childFolder.SetAttributeValueByString(cf.Name, cf.Value);
                    }

                    return childFolder;
                }
            }
            else
            {
                int count = parentFolder.SearchByTQL($"=>SUBPARTS[{Constants.ADOPlanID} == \"{ testCase.PlanID }\" AND { Constants.ADOSuiteID} == \"{ testCase.SuiteID }\"]").Count();
                if (count == 0)
                {
                    TCFolder childFolder = parentFolder.GetOrCreateTestCaseSubFolder(testCase.SuiteName);
                    childFolder.SetAttributeValueByString(Constants.ADOPlanID, testCase.PlanID);
                    childFolder.SetAttributeValueByString(Constants.ADOSuiteID, testCase.SuiteID);

                    foreach (CustomField cf in testCase.CustomFieldsTS)
                    {
                        childFolder.SetAttributeValueByString(cf.Name, cf.Value);
                    }

                    return childFolder;
                }
                else
                {
                    TCFolder childFolder = parentFolder.GetOrCreateTestCaseSubFolder(testCase.SuiteName);

                    foreach (CustomField cf in testCase.CustomFieldsTS)
                    {
                        childFolder.SetAttributeValueByString(cf.Name, cf.Value);
                    }

                    return childFolder;
                }
            }
        }

        public bool CreateUpdateTestCase(TCFolder parentFolder, ADOTestCase testCase, ITaskContext task, bool warnOnNameChange, bool skipNameChange, DateTime time)
        {
            List<PersistableObject> testCases = parentFolder.SearchByTQL($"=>SUBPARTS:TestCase[{ Constants.ADOPlanID }==\"{ testCase.PlanID }\" AND { Constants.ADOSuiteID }==\"{ testCase.SuiteID }\" AND { Constants.ADOTestCaseID }==\"{ testCase.TestCaseID }\"]");

            TestCase tc = null;
            if (testCases.Count == 0)
            {
                CreateTestCaseTask ct = new CreateTestCaseTask();
                tc = (TestCase)ct.Execute(parentFolder, task);
                tc.Name = testCase.TestCaseName;
                tc.SetAttributeValueByString(Constants.ADOPlanID, testCase.PlanID);
                tc.SetAttributeValueByString(Constants.ADOSuiteID, testCase.SuiteID);
                tc.SetAttributeValueByString(Constants.ADOTestCaseID, testCase.TestCaseID);

                tc.SetAttributeValueByString(Constants.ADOStatus, "Exists");

                foreach (CustomField cf in testCase.CustomFieldsTC)
                {
                    tc.SetAttributeValueByString(cf.Name, cf.Value);
                }
            }
            else if (warnOnNameChange && testCases.Count == 1)
            {
                tc = (TestCase)testCases[0];

                if (tc.DisplayedName != testCase.TestCaseName)
                {
                    if (warnOnNameChange)
                    {
                        MsgBoxResult resp = MsgBoxResult.None;
                        // will always hit first go
                        if (skipNameChange.Equals(false))
                        {
                            resp = task.ShowMessageBox_Ok_Cancel_DontShowAgain("Attention", "TestCase \"" + tc.Name + "\" name is different from ADO. Would you like to sync the name from ADO?");
                        }

                        if (resp == MsgBoxResult.OK)
                        {
                            tc.Name = testCase.TestCaseName;
                        }
                        else if (resp.Equals(MsgBoxResult.OkDontShowAgain) || skipNameChange.Equals(true))
                        {
                            tc.Name = testCase.TestCaseName;
                            skipNameChange = true;
                        }
                    }
                    else
                    {
                        tc.Name = testCase.TestCaseName;
                    }
                }

                foreach (CustomField cf in testCase.CustomFieldsTC)
                {
                    tc.SetAttributeValueByString(cf.Name, cf.Value);
                }


                tc.SetAttributeValueByString(Constants.LastSynced, time.ToString("d MMM yyyy hh:mm:ss tt"));
                tc.SetAttributeValueByString(Constants.ADOStatus, "Exists");
            }

            return skipNameChange;
        }

    }
}
