using Tricentis.AddIns.Integration.ADO;
using Tricentis.TCCore.Persistency;
using Tricentis.TCCore.Persistency.AddInManager;
using Tricentis.TCCore.Persistency.Tasks;

namespace ToscaAddIn
{
    class ToscaAddIn : TCAddIn
    {
        public override string UniqueName => Constants.AddInCategory;

        public override string DisplayedName => Constants.AddInCategory;

        public override void InitializeTaskContext(ITaskContext taskContext, CommanderInterfaceType interfaceType)
        {
            base.InitializeTaskContext(taskContext, interfaceType);
        }

        public override void InitializeAfterOpenWorkspace()
        {
            base.InitializeAfterOpenWorkspace();
        }
    }

}
