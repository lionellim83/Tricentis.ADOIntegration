using Tricentis.AddIns.Integration.ADO;
using Tricentis.TCCore.Base.TestConfigurations;
using Tricentis.TCCore.Persistency;

namespace Tricentis.AddIns.Integration.ADOAddIn.Tasks
{
    public abstract class BaseADOTask : Task
    {
        public PersistableObject _targetObject;

        protected string BaseAddress => ((IObjectWithConfiguration)_targetObject).TestConfiguration.GetConfigurationParamValue(Constants.ADOBaseURL);

        protected string PersonalAccessToken => ((IObjectWithConfiguration)_targetObject).TestConfiguration.GetConfigurationParamValue(Constants.ADOPersonalAccessToken);
      
        protected string Project => ((IObjectWithConfiguration)_targetObject).TestConfiguration.GetConfigurationParamValue(Constants.ADOProject);

        protected (string baseAddress, string pat, string project) GetConfiguration(PersistableObject obj)
        {
            IObjectWithConfiguration configObj = obj as IObjectWithConfiguration;

            string baseAddress = configObj.TestConfiguration.GetConfigurationParamValue(Constants.ADOBaseURL);
            string personalToken = configObj.TestConfiguration.GetConfigurationParamValue(Constants.ADOPersonalAccessToken);
            string project = configObj.TestConfiguration.GetConfigurationParamValue(Constants.ADOProject);

            return (baseAddress, personalToken, project);
        }
    }
}
    