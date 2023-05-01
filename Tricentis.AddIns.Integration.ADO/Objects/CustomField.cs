using System;

namespace Tricentis.AddIns.Integration.ADO.Objects
{
    public class CustomField
    {
        public string Name { get; set; }
        public string Value;
        public string ADOName;

        public CustomField()
        {

        }

        public CustomField(string name, string adoName)
        {
            Name = name;
            ADOName = adoName;
        }

        public CustomField Copy()
        {
            return new CustomField { Name = this.Name, Value = this.Value, ADOName = this.ADOName };
        }
    }

}
