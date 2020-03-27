using System.Collections.Generic;

namespace SwaggerToDocx.Models
{
    public class DefinitionModel
    {
        public List<string> Required { get; set; }
        public string Type { get; set; }
        public Dictionary<string, ParameterModel> Properties { get; set; }
    }
}