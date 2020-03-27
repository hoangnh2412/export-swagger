using System.Collections.Generic;

namespace SwaggerToDocx.Models
{
    public class ApiModel
    {
        public string Summary { get; set; }
        public string Description { get; set; }
        public string OperationId { get; set; }
        public List<string> Tags { get; set; }
        public List<string> Consumes { get; set; }
        public List<string> Produces { get; set; }
        public List<ParameterModel> Parameters { get; set; }
        public Dictionary<string, ResponseModel> Responses { get; set; }
    }
}