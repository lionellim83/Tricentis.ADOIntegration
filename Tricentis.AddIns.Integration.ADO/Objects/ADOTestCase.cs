using System.Collections.Generic;

namespace Tricentis.AddIns.Integration.ADO.Objects
{
    public class ADOTestCase
    {
        public string PlanID { get; set; }
        public string PlanName { get; set; }
        public string SuiteID { get; set; }
        public string ADOSuiteID { get; set; }
        public string SuiteName { get; set; }
        public string TestCaseID { get; set; }
        public string TestCaseName { get; set; }
        public string AssignedTo { get; set; }
        public bool SyncTestStep { get; set; }
        public string PropertyMapping { get; set; }

        public List<CustomField> CustomFieldsTC = new List<CustomField>();
        public List<CustomField> CustomFieldsTS = new List<CustomField>();
        public List<CustomField> CustomFieldsTP = new List<CustomField>();

        public List<ToscaTestStep> ToscaTestSteps = new List<ToscaTestStep>();

    }
}
