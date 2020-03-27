using System;
using System.Collections.Generic;

namespace SwaggerToDocx.Models
{
    public class ParameterModel
    {
        public string Name { get; set; }
        public string In { get; set; }
        public string Description { get; set; }
        public bool Required { get; set; }
        public bool UniqueItems { get; set; }
        public string Type { get; set; }
        public string Format { get; set; }
        public Dictionary<string, object> Schema { get; set; }
        public Dictionary<string, string> Items { get; set; }

        public Dictionary<string, object> Data { get; set; }

        public string Ref { get; set; }

    }
}