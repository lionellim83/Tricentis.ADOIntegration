using System;
using System.Collections.Generic;
using System.Text;
using Tricentis.AddIns.Integration.ADO.Enums;

namespace Tricentis.AddIns.Integration.ADO.Objects
{
    public class ToscaTestStep
    {
        public string Name;

        public List<ToscaTestAction> TestActions = new List<ToscaTestAction>();

        public String GenerateTestStepAction()
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine(Name);

            for (int i = 0; i < TestActions.Count; i++)
            {
                ToscaTestAction ta = TestActions[i];

                if (ta.ActionMode == ActionMode.Input)
                {
                    sb.AppendLine($"({i + 1}) Input '{ta.Value}' into {ta.Name}");
                }
                else if (ta.ActionMode == ActionMode.Verify)
                {
                    sb.AppendLine($"({i + 1}) Verify {ta.Name} value '{ta.Value}");
                }
                else if (ta.ActionMode == ActionMode.WaitOn)
                {
                    sb.AppendLine($"({i + 1}) WaitOn {ta.Name}");
                }
                else if (ta.ActionMode == ActionMode.Select)
                {
                    sb.AppendLine($"({i + 1}) The element {ta.Name} is Selected: {ta.Value}");
                }
                else if (ta.ActionMode == ActionMode.Buffer)
                {
                    sb.AppendLine($"({i + 1}) Buffer the element at {ta.Name} with the name {ta.Value}");
                }
                else if (ta.ActionMode == ActionMode.Constraints)
                {
                    sb.AppendLine($"({i + 1}) The element {ta.Name} is Selected to the value {ta.Value}");
                }
            }
            sb.ToString().Replace("\\n", null);
            return sb.ToString();
        }

        public String GenerateTestStepExpectedResult()
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine();

            for (int i = 0; i < TestActions.Count; i++)
            {
                ToscaTestAction ta = TestActions[i];

                if (ta.ActionMode == ActionMode.Input)
                {
                    sb.AppendLine($"({i + 1}) '{ta.Value}' was inputted");
                }
                else if (ta.ActionMode == ActionMode.Verify)
                {
                    sb.AppendLine($"({i + 1}) Value is '{ta.Value}'");
                }
                else if (ta.ActionMode == ActionMode.WaitOn)
                {
                    sb.AppendLine($"({i + 1}) WaitOn for '{ta.Value}' was successful");
                }
                else if (ta.ActionMode == ActionMode.Select)
                {
                    sb.AppendLine($"({i + 1}) The element {ta.Name} is Selected: {ta.Value}");
                }
                else if (ta.ActionMode == ActionMode.Buffer)
                {
                    sb.AppendLine($"({i + 1}) Buffer was successfully created");
                }
                else if (ta.ActionMode == ActionMode.Constraints)
                {
                    sb.AppendLine($"({i + 1}) The Constraint was successful{ta.Value}");
                }
            }

            return sb.ToString();
        }
    }
}
