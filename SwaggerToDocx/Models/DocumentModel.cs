using System.Collections.Generic;

namespace SwaggerToDocx.Models
{
    public class DocumentModel
    {
        public string Swagger { get; set; }
        public InfoModel Info { get; set; }
        public Dictionary<string, Dictionary<string, ApiModel>> Paths { get; set; }
        public Dictionary<string, DefinitionModel> Definitions { get; set; }
        public Dictionary<string, object> SecurityDefinitions { get; set; }
        public List<string> Tags { get; set; }
    }
}