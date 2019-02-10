
namespace LetterRollout
{
    public class Model
    {
        public string name { get; set; }
        public string empId { get; set; }
        public string emailId { get; set; }
        public string firstName { get; set; }

        public string agc { get; set; }
        public string apb { get; set; }
        public string atc { get; set; }
        public string basic { get; set; }
        public string hra { get; set; }
        public string pf { get; set; }
        public string flexible { get; set; }
        public string total { get; set; }

        //Title for promotion only
        public string title { get; set; }

        //Share
        public string shareoptions { get; set; }
        public string shareGrant { get; set; }

        public string cc { get; set; }
    }
}
