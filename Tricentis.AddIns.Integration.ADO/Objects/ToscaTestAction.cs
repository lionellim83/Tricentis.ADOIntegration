using Tricentis.AddIns.Integration.ADO.Enums;

namespace Tricentis.AddIns.Integration.ADO.Objects
{
    public class ToscaTestAction
    {
        public ToscaTestAction(string name, string value, ActionMode actionMode)
        {
            Name = name;
            Value = value;
            ActionMode = actionMode;
        }

        public string Name;
        public string Value;
        public ActionMode ActionMode;
    }
}
