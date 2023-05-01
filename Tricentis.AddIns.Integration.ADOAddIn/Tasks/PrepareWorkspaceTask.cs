using System;
using System.Collections.Generic;
using System.Linq;
using Tricentis.AddIns.Integration.ADO;
using Tricentis.TCCore.Base;
using Tricentis.TCCore.Base.Configurations;
using Tricentis.TCCore.BusinessObjects.Folders;
using Tricentis.TCCore.Persistency;
using Tricentis.TCCore.Persistency.Tasks;


namespace Tricentis.AddIns.Integration.ADOAddIn.Tasks
{
    public class PrepareWorkspaceTask : Tricentis.TCCore.Persistency.Task
    {
        public PrepareWorkspaceTask()
        {
        }

        public override TaskCategory Category => new TaskCategory(Constants.AddInCategory);

        public override string Name => "Prepare Workspace";

        public override bool RequiresChangeRights => true;

        public override object Execute(PersistableObject obj, ITaskContext context)
        {
            TCBaseProject project = Workspace.ActiveWorkspace.Project;

            //TestCase Properties
            TCObjectPropertiesDefinition tcPropDef = project.GetOrCreatePropertiesDefinition("TestCase");
            CreateProperty(tcPropDef, Constants.ADOPlanID);
            CreateProperty(tcPropDef, Constants.ADOSuiteID);
            CreateProperty(tcPropDef, Constants.ADOTestCaseID);
            CreateProperty(tcPropDef, Constants.ADOStatus, "Deleted;Exists;Unsynced");
            CreateProperty(tcPropDef, Constants.LastSynced);

            //Execution Properties
            TCObjectPropertiesDefinition elPropDef = project.GetOrCreatePropertiesDefinition("ExecutionTestCaseLog");
            CreateProperty(elPropDef, Constants.ADORunID);

            //Execution List Properties
            TCObjectPropertiesDefinition ePropDef = project.GetOrCreatePropertiesDefinition("ExecutionList");
            CreateProperty(ePropDef, Constants.ADOSync, "True;False", "True");
            //CreateProperty(ePropDef, Constants.Configuration, "", "Default");

            //Folder Properties
            TCObjectPropertiesDefinition cfPropDef = project.GetOrCreatePropertiesDefinition("Folder");
            CreateProperty(cfPropDef, Constants.ADOSync, "True;False", "False");
            CreateProperty(cfPropDef, Constants.ADOPlanID, "", "N/A");
            CreateProperty(cfPropDef, Constants.ADOSuiteID, "", "N/A");
            CreateProperty(cfPropDef, Constants.ADOCustomProperties, "", "");

            // Create Sample Configuration
            CreateConfiguration(project);

            return obj;
        }

        private void CreateProperty(TCObjectPropertiesDefinition category, string propertyName, string valueRange = "", string defaultValue = "")
        {
            List<TCTypedObjectProperty> propList = category.GetTypedProperties().ToList();

            if (!propList.Exists(p => p.Name == propertyName))
            {
                TCTypedObjectProperty property = TCTypedObjectProperty.Create();
                property.Name = propertyName;
                property.Visible = true;
                property.ValueRange = valueRange;
                property.Value = defaultValue;
                category.Properties.Add(property);
            }
        }

        private void CreateConfiguration(TCBaseProject project)
        {
            TCFolder cf = (TCFolder)project.SearchByTQL("=>SUBPARTS:TCFolder[PossibleContent==\"Configuration\"]").FirstOrDefault();

            TCConfiguration config = (TCConfiguration)cf.Items.FirstOrDefault(c => c.GetType() == typeof(TCConfiguration) && c.Name == Constants.ADOConfiguration);
            if (config == null)
            {
                config = TCConfiguration.Create();
                config.Name = Constants.ADOConfiguration;
                cf.Items.Add(config);
            }

            if (!config.TestConfiguration.ConfigurationParamNames.Contains(Constants.ADOBaseURL))
            {
                config.TestConfiguration.SetConfigurationParam(Constants.ADOBaseURL, "");
            }

            if (!config.TestConfiguration.ConfigurationParamNames.Contains(Constants.ADOPersonalAccessToken))
            {
                config.TestConfiguration.SetConfigurationParam(Constants.ADOPersonalAccessToken, "");
            }

            if (!config.TestConfiguration.ConfigurationParamNames.Contains(Constants.ADOProject))
            {
                config.TestConfiguration.SetConfigurationParam(Constants.ADOProject, "");
            }

            if (!config.TestConfiguration.ConfigurationParamNames.Contains(Constants.ADOTestStepSwitch))
            {
                config.TestConfiguration.SetConfigurationParam(Constants.ADOTestStepSwitch, "");
            }
        }

    }
}