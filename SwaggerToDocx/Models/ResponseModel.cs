using System.Collections.Generic;

namespace SwaggerToDocx.Models
{
    public class ResponseModel
    {
        public string Description { get; set; }
        public Dictionary<string, object> Schema { get; set; }
    }
}