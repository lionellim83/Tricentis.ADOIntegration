namespace Tricentis.AddIns.Integration.ADO.Objects
{
    public class DefectLinks
    {
        public int _TestCaseID { get; set; }
        public int _RunID { get; set; }
        public int _ResultID { get; set; }

        public DefectLinks(int TestCaseID, int RunID)
        {
            this._ResultID = 100000;
            this._RunID = RunID;
            this._TestCaseID = TestCaseID;
        }
    }
}
