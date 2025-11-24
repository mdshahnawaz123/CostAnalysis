namespace CostAnalysis.Model
{
    public class UserRecord
    {
        public string Username { get; set; }
        public string Password { get; set; } // plaintext in JSON today — migrate to hash later
        public bool Active { get; set; }
        public System.DateTime Expires { get; set; }
    }
}
