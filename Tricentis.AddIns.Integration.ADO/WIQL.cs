namespace Tricentis.AddIns.Integration.ADO
{
    public class WIQL
    {
        public WIQL(string query)
        {
            this.query = query;
        }

        public string query { get; set; }
    }
}
