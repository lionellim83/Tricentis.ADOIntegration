using Microsoft.TeamFoundation.TestManagement.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
using Microsoft.VisualStudio.Services.Common;
using Microsoft.VisualStudio.Services.WebApi;
using Microsoft.VisualStudio.Services.WebApi.Patch;
using Microsoft.VisualStudio.Services.WebApi.Patch.Json;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using Tricentis.AddIns.Integration.ADO.Objects;

namespace Tricentis.AddIns.Integration.ADO
{
    public class ADOSync
    {
        private string _personalToken;
        private string _baseAddress;
        private string _project;

        private WorkItemTrackingHttpClient _workItemTrackingHttpClient;
        private TestManagementHttpClient _testManagementHttpClient;

        public ADOSync()
        {

        }
        public ADOSync(String baseAddress, string personalAccessToken, string project)
        {
            _baseAddress = baseAddress;
            _personalToken = personalAccessToken;
            _project = project;

            InitClients();
        }

        public void InitClients()
        {
            var credentials = new VssBasicCredential(string.Empty, _personalToken);
            var connection = new VssConnection(new Uri(_baseAddress), credentials);
            try
            {
                _workItemTrackingHttpClient = connection.GetClient<WorkItemTrackingHttpClient>();
                _testManagementHttpClient = connection.GetClient<TestManagementHttpClient>();
            }
            catch
            {
                return;
            }
        }

        public List<ADOTestCase> GetEpicAndUserStory()
        {
            List<ADOTestCase> tcs = new List<ADOTestCase>();
            WIQL wiql = new WIQL($"SELECT [System.Id], [System.Title] FROM WorkItems WHERE [System.WorkItemType] = 'User Story'");
            string API_resource = "_apis/wit/wiql?api-version=4.1";
            StringContent content = new StringContent(JsonConvert.SerializeObject(wiql), Encoding.UTF8, "application/json");
            string credentials = Convert.ToBase64String(ASCIIEncoding.ASCII.GetBytes(string.Format("{0}:{1}", "", _personalToken)));

            using (var client = new HttpClient())
            {
                client.BaseAddress = new Uri(_baseAddress + "/" + _project + "/");
                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", credentials);

                //connect to the REST endpoint            
                HttpResponseMessage response = client.PostAsync(API_resource, content).Result;

                if (response.IsSuccessStatusCode)
                {
                    JObject jObj_workItems = JObject.Parse(response.Content.ReadAsStringAsync().Result);
                    JToken jToken_workItems = jObj_workItems.SelectToken("columns");
                    List<string> customFieldReference = jToken_workItems.Select(s => (string)s.SelectToken("referenceName")).ToList();
                }
            }

            return tcs;
        }

        public List<String> GetCustomProperties(string id, List<String> customFieldNames, string workItemType)
        {
            List<String> customFieldValues = new List<string>();

            if (customFieldNames.Count > 0)
            {
                string aggFieldNames = "[" + customFieldNames.Aggregate((c1, c2) => c1 + "], [" + c2) + "]";

                WIQL wiql = new WIQL($"SELECT { aggFieldNames } FROM WorkItems WHERE [Work Item Type] = '{ workItemType }' AND [ID] = '{ id }'");
                string API_resource = "_apis/wit/wiql?api-version=4.1";
                StringContent content = new StringContent(JsonConvert.SerializeObject(wiql), Encoding.UTF8, "application/json");
                string credentials = Convert.ToBase64String(ASCIIEncoding.ASCII.GetBytes(string.Format("{0}:{1}", "", _personalToken)));

                using (var client = new HttpClient())
                {
                    client.BaseAddress = new Uri(_baseAddress + "/" + _project + "/");
                    client.DefaultRequestHeaders.Accept.Clear();
                    client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", credentials);

                    //connect to the REST endpoint            
                    HttpResponseMessage response = client.PostAsync(API_resource, content).Result;

                    if (response.IsSuccessStatusCode)
                    {
                        JObject jObj_workItems = JObject.Parse(response.Content.ReadAsStringAsync().Result);
                        JToken jToken_workItems = jObj_workItems.SelectToken("columns");
                        WorkItem wItem = _workItemTrackingHttpClient.GetWorkItemAsync(Int32.Parse(id), expand: WorkItemExpand.All).Result;

                        foreach (String customFieldName in customFieldNames)
                        {
                            string customFieldReference = jToken_workItems.Where(t => (string)t.SelectToken("name") == customFieldName).Select(s => (string)s.SelectToken("referenceName")).FirstOrDefault();
                            Object o = null;
                            if (wItem.Fields.TryGetValue(customFieldReference, out o))
                            {
                                if (o is IdentityRef)
                                {
                                    IdentityRef ir = (IdentityRef)o;
                                    customFieldValues.Add(ir.DisplayName);
                                }
                                else
                                {
                                    customFieldValues.Add(o.ToString());
                                }
                            }
                            else
                            {
                                customFieldValues.Add("");
                            }
                        }
                    }
                }
            }

            return customFieldValues;
        }

        public (List<ADOTestCase>, string, bool) GetTestCases()
        {
            string plan_name = "", suite_name = "";
            List<ADOTestCase> adoTestCases = new List<ADOTestCase>();
            List<string> testPlanIds = new List<string>();
            bool valid = false;

            WIQL wiql = new WIQL("Select [System.Id], [System.Title], [System.State] From WorkItems Where [System.WorkItemType] = 'Test Plan'");

            string API_resource = "_apis/wit/wiql?api-version=4.1";
            StringContent content = new StringContent(JsonConvert.SerializeObject(wiql), Encoding.UTF8, "application/json");
            string credentials = Convert.ToBase64String(ASCIIEncoding.ASCII.GetBytes(string.Format("{0}:{1}", "", _personalToken)));
            HttpResponseMessage response;
            using (var client = new HttpClient())
            {
                client.BaseAddress = new Uri(_baseAddress + "/" + _project + "/");
                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", credentials);

                //connect to the REST endpoint            
                response = client.PostAsync(API_resource, content).Result;

                //check to see if we have a successful response
                if (response.IsSuccessStatusCode)
                {
                    JObject jObj_workItems = JObject.Parse(response.Content.ReadAsStringAsync().Result);
                    JToken jToken_workItems = jObj_workItems.SelectToken("workItems");
                    testPlanIds = jToken_workItems.Select(s => (string)s.SelectToken("id")).ToList();

                    foreach (string testPlan_id in testPlanIds)
                    {
                        plan_name = GetResourceNameById(testPlan_id);
                        API_resource = "_apis/test/plans/" + testPlan_id + "/suites?api-version=4.1";

                        response = client.GetAsync(API_resource).Result;
                        if (response.IsSuccessStatusCode)
                        {
                            JObject jObj_testSuites = JObject.Parse(response.Content.ReadAsStringAsync().Result);
                            List<JToken> jToken_testSuites = jObj_testSuites.SelectToken("value").Children().Where(s => !s.SelectToken("testCaseCount").ToString().Equals("0")).ToList();
                            List<string> testSuite_Ids = jToken_testSuites.Select(s => (string)s.SelectToken("id")).ToList();

                            foreach (string testSuite_id in testSuite_Ids)
                            {
                                suite_name = GetResourceNameById(testSuite_id);
                                API_resource = "_apis/test/plans/" + testPlan_id + "/suites/" + testSuite_id + "/testcases?api-version=4.1";

                                response = client.GetAsync(API_resource).Result;
                                if (response.IsSuccessStatusCode)
                                {
                                    JObject jObj_testCases = JObject.Parse(response.Content.ReadAsStringAsync().Result);
                                    List<JToken> jToken_testCases = jObj_testCases.SelectToken("value").Children().ToList();

                                    foreach (JToken jToken_testCase in jToken_testCases)
                                    {
                                        string tc_id = jToken_testCase.SelectToken("testCase.id").ToString();
                                        string url = jToken_testCase.SelectToken("testCase.url").ToString();
                                        response = client.GetAsync(url).Result;

                                        JObject jObject_TC = JObject.Parse(response.Content.ReadAsStringAsync().Result);
                                        string tc_name = jObject_TC.SelectToken("fields.['System.Title']").ToString();

                                        adoTestCases.Add(new ADOTestCase() { PlanID = testPlan_id, PlanName = plan_name, SuiteID = testSuite_id, SuiteName = suite_name, TestCaseID = tc_id, TestCaseName = tc_name });
                                        valid = true;
                                    }
                                }
                                else
                                {
                                    return (adoTestCases, response.ReasonPhrase, valid);
                                }
                            }
                        }
                    }
                }
                else
                {
                    return (adoTestCases, response.ReasonPhrase, valid);
                }
            }
            return (adoTestCases, response.ReasonPhrase, valid);
        }

        public (int, int) UpdateExecutionResult(int TestPlanId, int testSuiteId, int TestCaseId, string start_time, string end_time, string result, string resultsLog)
        {
            CultureInfo c = CultureInfo.CurrentCulture;
            TestPoint testPoint = _testManagementHttpClient.GetPointsAsync(_project, TestPlanId, testSuiteId, testCaseId: TestCaseId.ToString()).Result.FirstOrDefault();

            string testPlanName = GetResourceNameById(TestPlanId.ToString());
            string testCaseName = GetResourceNameById(TestCaseId.ToString());
            string TestRunState = (result == "Passed" ? "Completed" : "NeedsInvestigation");

            var testPlanRef = new ShallowReference(TestPlanId.ToString(), name: testPlanName);

            RunCreateModel runCreate = new RunCreateModel(testSettings: default,
                                                            name: testCaseName,
                                                            plan: testPlanRef,
                                                            startedDate: DateTime.Parse(start_time, c).ToUniversalTime().ToString("o"),
                                                            pointIds: new int[] { testPoint.Id });

            TestRun testRun = _testManagementHttpClient.CreateTestRunAsync(runCreate, _project).Result;

            TestCaseResult testResult = CreateTestResult(result, end_time, resultsLog);
            _ = _testManagementHttpClient.UpdateTestResultsAsync(new TestCaseResult[] { testResult }, _project, testRun.Id).Result;

            RunUpdateModel runUpdateModel = new RunUpdateModel(completedDate: DateTime.Parse(end_time, c).ToUniversalTime().ToString("o"), state: TestRunState);
            _ = _testManagementHttpClient.UpdateTestRunAsync(runUpdateModel, _project, testRun.Id).Result;

            return (testRun.Id, testResult.Id);
        }

        public TestCaseResult CreateTestResult(string result, string end_time, string resultLog)
        {
            CultureInfo c = CultureInfo.CurrentCulture;
            TestCaseResult testCaseResult = new TestCaseResult
            {
                Id = 100000,
                Outcome = result,
                CompletedDate = DateTime.Parse(end_time, c).ToUniversalTime(),
                State = "Completed",
                ErrorMessage = resultLog
            };

            return testCaseResult;
        }

        public (int, string) CreateBug(ADODefect defectInstance)
        {
            WorkItem w = null;
            if (defectInstance.Identification == -1)
            {
                JsonPatchDocument patchDocument = new JsonPatchDocument();

                //Add field and their values to your patch document
                patchDocument.Add(new JsonPatchOperation()
                {
                    Operation = Operation.Add,
                    Path = "/fields/System.Title",
                    Value = defectInstance.Name
                });

                //Body message
                patchDocument.Add(
                    new JsonPatchOperation()
                    {
                        Operation = Operation.Add,
                        Path = "/fields/Microsoft.VSTS.TCM.ReproSteps",
                        Value = "testing bug creation"
                    }
                );
                //adds the priority
                patchDocument.Add(
                    new JsonPatchOperation()
                    {
                        Operation = Operation.Add,
                        Path = "/fields/Microsoft.VSTS.Common.Priority",
                        Value = defectInstance.Priority
                    }
                );
                //adds the severity
                patchDocument.Add(
                    new JsonPatchOperation()
                    {
                        Operation = Operation.Add,
                        Path = "/fields/Microsoft.VSTS.Common.Severity",
                        Value = defectInstance.SeverityText
                    }
                );

                try
                {
                    w = _workItemTrackingHttpClient.CreateWorkItemAsync(patchDocument, _project, "Bug").Result;

                    foreach (DefectLinks defectLink in defectInstance.DefectLinks)
                    {
                        LinkBugsToTestCases(defectLink, w);
                    }
                }
                catch (Exception e)
                {
                    return (-1, e.InnerException.Message);
                }
            }

            if (w == null)
            {
                return (-1, String.Empty);
            }
            else
            {
                return (w.Id.Value, w.Fields["System.State"].ToString());
            }
        }

        private void LinkBugsToTestCases(DefectLinks defectLink, WorkItem wItemBug)
        {
            // gets the test case work item
            WorkItem wItemTC = _workItemTrackingHttpClient.GetWorkItemAsync(defectLink._TestCaseID, expand: WorkItemExpand.All).Result;
            //create the patchDocument adding the link to the test case
            JsonPatchDocument patchUpdateDocument = new JsonPatchDocument();
            patchUpdateDocument.Add(
                new JsonPatchOperation()
                {
                    Operation = Operation.Add,
                    Path = "/relations/-",
                    Value = new
                    {
                        rel = "Microsoft.VSTS.Common.TestedBy-Forward",
                        url = wItemTC.Url
                    }
                }
            );

            // updates the changes to the work item in ADO
            wItemBug = _workItemTrackingHttpClient.UpdateWorkItemAsync(patchUpdateDocument, wItemBug.Id.Value).Result;

            TestCaseResult f;
            f = _testManagementHttpClient.GetTestResultByIdAsync(_project, defectLink._RunID, defectLink._ResultID).Result;

            List<ShallowReference> bugs = new List<ShallowReference>();
            bugs.Add(new ShallowReference(wItemBug.Id.ToString(), url: wItemBug.Url));
            f.AssociatedBugs = bugs;

            _ = _testManagementHttpClient.UpdateTestResultsAsync(new TestCaseResult[] { f }, _project, defectLink._RunID).Result;
        }

        public (ADODefect, bool) RetrieveBug(int bugId)
        {
            try
            {
                WorkItem w = _workItemTrackingHttpClient.GetWorkItemAsync(bugId, expand: WorkItemExpand.All).Result;
                return (new ADODefect(w), false);
            }
            catch
            {
                return (null, true);
            }
        }

        private string GetResourceNameById(string resourceId)
        {
            WorkItem wItem = _workItemTrackingHttpClient.GetWorkItemAsync(Int32.Parse(resourceId), expand: WorkItemExpand.All).Result;
            return wItem.Fields["System.Title"].ToString();
        }

        public string SetTestSteps(ADOTestCase tc)
        {
            WorkItem wItem = _workItemTrackingHttpClient.GetWorkItemAsync(Int32.Parse(tc.TestCaseID), expand: WorkItemExpand.All).Result;
            string test = GetResourceNameById(tc.TestCaseID);

            // Create and instance of TestBaseHelper Class and generate ITestBase Object using that
            TestBaseHelper helper = new TestBaseHelper();
            ITestBase testBase = helper.Create();

            ITestStep testStep;
            foreach (ToscaTestStep ts in tc.ToscaTestSteps)
            {
                testStep = testBase.CreateTestStep();
                testStep.Title = ts.GenerateTestStepAction();
                testStep.ExpectedResult = ts.GenerateTestStepExpectedResult();
                testBase.Actions.Add(testStep);
            }
            //Update Test case object
            JsonPatchDocument updatePatchDocument = new JsonPatchDocument();
            updatePatchDocument = testBase.SaveActions(updatePatchDocument);
            // update testcase wit using new json
            wItem = _workItemTrackingHttpClient.UpdateWorkItemAsync(updatePatchDocument, wItem.Id.Value).Result;

            return test;
        }

    }
}
