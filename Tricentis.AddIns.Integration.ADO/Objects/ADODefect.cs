using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Tricentis.AddIns.Integration.ADO.Enums;

namespace Tricentis.AddIns.Integration.ADO.Objects
{
    public class ADODefect
    {
        private string _name;
        public string Name
        {
            get { return _name; }
            set { _name = value; }
        }

        private int _identification;
        public int Identification
        {
            get { return _identification; }
            set { _identification = value; }
        }

        private string _state;
        public string State
        {
            get { return _state; }
            set { _state = value; }
        }

        private int _priority;
        public int Priority
        {
            get { return _priority; }
            set { _priority = value; }
        }

        private int _severity;
        public int Severity
        {
            get { return _severity; }
            set { _severity = value; }
        }
        public string SeverityText
        {
            get
            {
                SeverityValue sv = (SeverityValue)_severity;
                switch (sv)
                {
                    case SeverityValue.Critical:
                        return "1 - Critical";
                    case SeverityValue.High:
                        return "2 - High";
                    case SeverityValue.Medium:
                        return "3 - Medium";
                    case SeverityValue.Low:
                        return "4 - Low";
                    default: return _severity.ToString();
                }
            }
            set
            {
                Match m = new Regex("^[0-9]{1,}").Match(value);
                _severity = int.Parse(m.Value);
            }
        }

        private string _uniqueId;
        public string UniqueId
        {
            get { return _uniqueId; }
            set { _uniqueId = value; }
        }

        private List<DefectLinks> _defectLinks;
        public List<DefectLinks> DefectLinks
        {
            get { return _defectLinks; }
            set { _defectLinks = value; }
        }

        public ADODefect(int id, string name, int severity, int priority, List<DefectLinks> defectLinks, string uniqueID)
        {
            _identification = id;
            _name = name;
            _priority = priority;
            _uniqueId = uniqueID;
            _defectLinks = defectLinks;

            SeverityValue value = (SeverityValue)severity;
            switch (value)
            {
                case SeverityValue.Critical:
                    SeverityText = "1 - Critical";
                    break;
                case SeverityValue.High:
                    SeverityText = "2 - High";
                    break;
                case SeverityValue.Medium:
                    SeverityText = "3 - Medium";
                    break;
                case SeverityValue.Low:
                    SeverityText = "4 - Low";
                    break;
                default:
                    break;
            }
        }

        public ADODefect(WorkItem w)
        {
            _identification = w.Id.Value;
            _name = w.Fields["System.Title"].ToString();
            _state = w.Fields["System.State"].ToString();
            _priority = Int32.Parse(w.Fields["Microsoft.VSTS.Common.Priority"].ToString());
            SeverityText = w.Fields["Microsoft.VSTS.Common.Severity"].ToString();
        }

        public List<CustomField> CustomFields = new List<CustomField>();
    }
}
