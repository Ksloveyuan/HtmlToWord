using System.Runtime.Serialization;

namespace HtmlToWord.Contract
{
    [DataContract(Name = "ConvertResult")]
    public class CovertResult
    {
        [DataMember] public string FileUrl { get; set; }

        [DataMember] public bool Success { get; set; }

        [DataMember] public string Message { get; set; }
    }
}