using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using SwaggerToDocx.Models;

namespace SwaggerToDocx
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");

            var pathInput = Path.Combine(AppContext.BaseDirectory, "api.json");
            var json = File.ReadAllText(pathInput);
            json = json.Replace("$ref", "ref");
            var document = JsonConvert.DeserializeObject<DocumentModel>(json);

            var pathOutput = Path.Combine(AppContext.BaseDirectory, "api.xlsx");
            if (File.Exists(pathOutput))
                File.Delete(pathOutput);

            using (var package = new ExcelPackage(new FileInfo(pathOutput)))
            {
                var workSheet = package.Workbook.Worksheets.Add("APIs");
                workSheet.PrinterSettings.LeftMargin = (decimal)0.2;
                workSheet.PrinterSettings.RightMargin = 0;
                workSheet.View.PageLayoutView = true;
                workSheet.View.ShowGridLines = false;

                workSheet.Cells.Style.WrapText = true;
                workSheet.Cells.Style.Font.Name = "Calibri (Body)";
                workSheet.Cells.Style.Font.Size = 10;

                workSheet.Column(1).Width = 5;
                workSheet.Column(2).Width = 3;
                workSheet.Column(3).Width = 13;
                workSheet.Column(4).Width = 15;
                workSheet.Column(5).Width = 47.6;
                workSheet.Column(6).Width = 9;

                //Title
                workSheet.Cells[1, 1, 1, 6].Merge = true;
                workSheet.Row(1).Height = 25;
                workSheet.Cells[1, 1, 1, 6].Style.Font.Size = 24;
                workSheet.Cells[1, 1, 1, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[1, 1, 1, 6].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells[1, 1, 1, 6].Style.Font.Color.SetColor(1, 48, 84, 150);
                workSheet.Cells[1, 1, 1, 6].Value = document.Info.Title;

                //Description
                workSheet.Cells[2, 1, 2, 6].Merge = true;
                workSheet.Cells[2, 1, 2, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[2, 1, 2, 6].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells[2, 1, 2, 6].Value = document.Info.Description;
                workSheet.Row(2).CustomHeight = true;

                //Version
                workSheet.Cells[3, 1, 3, 6].Merge = true;
                workSheet.Cells[3, 1, 3, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[3, 1, 3, 6].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells[3, 1, 3, 6].Value = $"Version: {document.Info.Version}";


                var row = 5;
                var indexApi = 1;
                foreach (var path in document.Paths)
                {
                    var url = path.Key;
                    var methods = path.Value.First();
                    var method = methods.Key;
                    var api = methods.Value;
                    var summary = api.Summary;
                    var inputType = api.Consumes.FirstOrDefault();
                    var outputType = api.Produces.FirstOrDefault();
                    var paramQuery = api.Parameters.Where(x => x.In == "query").ToList();
                    var paramBody = api.Parameters.FirstOrDefault(x => x.In == "body");
                    var response = api.Responses.ContainsKey("200") ? api.Responses["200"] : null;

                    //Info
                    workSheet.Cells[row, 1, row, 6].Style.Font.Bold = true;
                    workSheet.Cells[row, 1, row, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    workSheet.Cells[row, 1, row, 6].Style.Fill.BackgroundColor.SetColor(1, 180, 198, 231);
                    workSheet.Cells[row, 1].Value = indexApi;
                    workSheet.Cells[row, 2, row, 6].Style.Font.Bold = true;
                    workSheet.Cells[row, 2, row, 6].Merge = true;
                    workSheet.Cells[row, 2, row, 6].Value = url;
                    workSheet.Row(row).CustomHeight = true;
                    row++;
                    workSheet.Cells[row, 2, row, 3].Merge = true;
                    workSheet.Cells[row, 2, row, 3].Value = "Summary";
                    workSheet.Cells[row, 4, row, 6].Merge = true;
                    workSheet.Cells[row, 4, row, 6].Value = summary;
                    workSheet.Row(row).CustomHeight = true;
                    row++;
                    workSheet.Cells[row, 2, row, 3].Merge = true;
                    workSheet.Cells[row, 2, row, 3].Value = "Method";
                    workSheet.Cells[row, 4, row, 6].Merge = true;
                    workSheet.Cells[row, 4, row, 6].Value = method.ToUpper();
                    row++;

                    //Parameter
                    if (paramQuery.Count > 0)
                    {
                        row++;
                        workSheet.Cells[row, 2, row, 3].Merge = true;
                        workSheet.Cells[row, 2, row, 3].Value = "Parameters";
                        workSheet.Cells[row, 4, row, 6].Merge = true;
                        workSheet.Cells[row, 4, row, 6].Value = "Query string";
                        row++;
                        workSheet.Cells[row, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        workSheet.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(1, 169, 208, 142);
                        workSheet.Cells[row, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        workSheet.Cells[row, 1].Style.Font.Bold = true;
                        workSheet.Cells[row, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[row, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        workSheet.Cells[row, 1].Value = "STT";
                        workSheet.Cells[row, 2, row, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        workSheet.Cells[row, 2, row, 3].Style.Fill.BackgroundColor.SetColor(1, 169, 208, 142);
                        workSheet.Cells[row, 2, row, 3].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        workSheet.Cells[row, 2, row, 3].Style.Font.Bold = true;
                        workSheet.Cells[row, 2, row, 3].Merge = true;
                        workSheet.Cells[row, 2, row, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[row, 2, row, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        workSheet.Cells[row, 2, row, 3].Value = "Field";
                        workSheet.Cells[row, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        workSheet.Cells[row, 4].Style.Fill.BackgroundColor.SetColor(1, 169, 208, 142);
                        workSheet.Cells[row, 4].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        workSheet.Cells[row, 4].Style.Font.Bold = true;
                        workSheet.Cells[row, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[row, 4].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        workSheet.Cells[row, 4].Value = "Type";
                        workSheet.Cells[row, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        workSheet.Cells[row, 5].Style.Fill.BackgroundColor.SetColor(1, 169, 208, 142);
                        workSheet.Cells[row, 5].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        workSheet.Cells[row, 5].Style.Font.Bold = true;
                        workSheet.Cells[row, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[row, 5].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        workSheet.Cells[row, 5].Value = "Description";
                        workSheet.Cells[row, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        workSheet.Cells[row, 6].Style.Fill.BackgroundColor.SetColor(1, 169, 208, 142);
                        workSheet.Cells[row, 6].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        workSheet.Cells[row, 6].Style.Font.Bold = true;
                        workSheet.Cells[row, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[row, 6].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        workSheet.Cells[row, 6].Value = "Required";
                        row++;

                        for (int i = 0; i < paramQuery.Count; i++)
                        {
                            var parameter = paramQuery[i];

                            workSheet.Cells[row, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            workSheet.Cells[row, 1].Value = i + 1;

                            workSheet.Cells[row, 2, row, 3].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            workSheet.Cells[row, 2, row, 3].Merge = true;
                            workSheet.Cells[row, 2, row, 3].Value = parameter.Name;

                            workSheet.Cells[row, 4].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            workSheet.Cells[row, 4].Value = string.IsNullOrWhiteSpace(parameter.Format) ? $"{parameter.Type}" : $"{parameter.Type} ({parameter.Format})";

                            workSheet.Cells[row, 5].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            workSheet.Cells[row, 5].Value = parameter.Description;
                            workSheet.Row(row).CustomHeight = true;

                            workSheet.Cells[row, 6].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            workSheet.Cells[row, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            workSheet.Cells[row, 6].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            if (parameter.Required)
                            {
                                workSheet.Cells[row, 6].Value = "x";
                            }
                            row++;
                        }
                    }

                    //Body
                    if (paramBody != null)
                    {
                        row++;
                        workSheet.Cells[row, 2, row, 3].Merge = true;
                        workSheet.Cells[row, 2, row, 3].Value = "Body type";
                        workSheet.Cells[row, 4, row, 6].Merge = true;
                        workSheet.Cells[row, 4, row, 6].Value = inputType;
                        row++;
                        workSheet.Cells[row, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        workSheet.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(1, 169, 208, 142);
                        workSheet.Cells[row, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        workSheet.Cells[row, 1].Style.Font.Bold = true;
                        workSheet.Cells[row, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[row, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        workSheet.Cells[row, 1].Value = "STT";
                        workSheet.Cells[row, 2, row, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        workSheet.Cells[row, 2, row, 3].Style.Fill.BackgroundColor.SetColor(1, 169, 208, 142);
                        workSheet.Cells[row, 2, row, 3].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        workSheet.Cells[row, 2, row, 3].Style.Font.Bold = true;
                        workSheet.Cells[row, 2, row, 3].Merge = true;
                        workSheet.Cells[row, 2, row, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[row, 2, row, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        workSheet.Cells[row, 2, row, 3].Value = "Field";
                        workSheet.Cells[row, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        workSheet.Cells[row, 4].Style.Fill.BackgroundColor.SetColor(1, 169, 208, 142);
                        workSheet.Cells[row, 4].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        workSheet.Cells[row, 4].Style.Font.Bold = true;
                        workSheet.Cells[row, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[row, 4].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        workSheet.Cells[row, 4].Value = "Type";
                        workSheet.Cells[row, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        workSheet.Cells[row, 5].Style.Fill.BackgroundColor.SetColor(1, 169, 208, 142);
                        workSheet.Cells[row, 5].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        workSheet.Cells[row, 5].Style.Font.Bold = true;
                        workSheet.Cells[row, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[row, 5].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        workSheet.Cells[row, 5].Value = "Description";
                        workSheet.Cells[row, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        workSheet.Cells[row, 6].Style.Fill.BackgroundColor.SetColor(1, 169, 208, 142);
                        workSheet.Cells[row, 6].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        workSheet.Cells[row, 6].Style.Font.Bold = true;
                        workSheet.Cells[row, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[row, 6].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        workSheet.Cells[row, 6].Value = "Required";
                        row++;

                        if (paramBody.Schema != null)
                        {
                            var schema = "";
                            if (paramBody.Schema.ContainsKey("ref"))
                            {
                                schema = (paramBody.Schema["ref"].ToString()).Replace("#/definitions/", "");
                            }
                            if (paramBody.Schema.ContainsKey("items"))
                            {
                                schema = (JsonConvert.DeserializeObject<Dictionary<string, string>>(paramBody.Schema["items"].ToString())["ref"]).Replace("#/definitions/", "");
                            }

                            //var schema = paramBody.Schema["ref"].Replace("#/definitions/", "");
                            if (document.Definitions.ContainsKey(schema))
                            {
                                var modelInput = document.Definitions[schema];
                                var indexProperty = 0;
                                foreach (var property in modelInput.Properties)
                                {
                                    indexProperty++;
                                    workSheet.Cells[row, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                    workSheet.Cells[row, 1].Value = indexProperty;

                                    workSheet.Cells[row, 2, row, 3].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                    workSheet.Cells[row, 2, row, 3].Merge = true;
                                    workSheet.Cells[row, 2, row, 3].Value = property.Key;

                                    workSheet.Cells[row, 4].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                    workSheet.Cells[row, 4].Value = string.IsNullOrWhiteSpace(property.Value.Format) ? $"{property.Value.Type}" : $"{property.Value.Type} ({property.Value.Format})";

                                    workSheet.Cells[row, 5].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                    workSheet.Cells[row, 5].Value = property.Value.Description;
                                    workSheet.Row(row).CustomHeight = true;

                                    workSheet.Cells[row, 6].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                    workSheet.Cells[row, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                    workSheet.Cells[row, 6].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                    if (modelInput.Required != null && modelInput.Required.Any(x => x == property.Key))
                                    {
                                        workSheet.Cells[row, 6].Value = "x";
                                    }
                                    row++;
                                }

                                workSheet.Cells[row, 2, row, 3].Merge = true;
                                workSheet.Cells[row, 2, row, 3].Value = "Input example";
                                row++;
                                workSheet.Cells[row, 3, row, 6].Merge = true;
                                workSheet.Cells[row, 3, row, 6].Style.WrapText = true;
                                workSheet.Row(row).CustomHeight = true;
                                workSheet.Cells[row, 3, row, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                workSheet.Cells[row, 3, row, 6].Style.Fill.BackgroundColor.SetColor(1, 244, 176, 132);
                                workSheet.Cells[row, 3, row, 6].Value = DefinitionToJson(document.Definitions, modelInput);
                                row++;
                            }
                        }

                    }

                    //Response
                    //nếu response là 1 object
                    if (response != null && response.Schema != null)
                    {
                        row++;
                        workSheet.Cells[row, 2, row, 3].Merge = true;
                        workSheet.Cells[row, 2, row, 3].Value = "Output";
                        workSheet.Cells[row, 4, row, 6].Merge = true;
                        workSheet.Cells[row, 4, row, 6].Value = outputType;

                        if (response.Schema.ContainsKey("ref"))
                        {
                            row++;
                            workSheet.Cells[row, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            workSheet.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(1, 169, 208, 142);
                            workSheet.Cells[row, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            workSheet.Cells[row, 1].Style.Font.Bold = true;
                            workSheet.Cells[row, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            workSheet.Cells[row, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            workSheet.Cells[row, 1].Value = "STT";
                            workSheet.Cells[row, 2, row, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            workSheet.Cells[row, 2, row, 3].Style.Fill.BackgroundColor.SetColor(1, 169, 208, 142);
                            workSheet.Cells[row, 2, row, 3].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            workSheet.Cells[row, 2, row, 3].Style.Font.Bold = true;
                            workSheet.Cells[row, 2, row, 3].Merge = true;
                            workSheet.Cells[row, 2, row, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            workSheet.Cells[row, 2, row, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            workSheet.Cells[row, 2, row, 3].Value = "Field";
                            workSheet.Cells[row, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            workSheet.Cells[row, 4].Style.Fill.BackgroundColor.SetColor(1, 169, 208, 142);
                            workSheet.Cells[row, 4].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            workSheet.Cells[row, 4].Style.Font.Bold = true;
                            workSheet.Cells[row, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            workSheet.Cells[row, 4].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            workSheet.Cells[row, 4].Value = "Type";
                            workSheet.Cells[row, 5, row, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            workSheet.Cells[row, 5, row, 6].Style.Fill.BackgroundColor.SetColor(1, 169, 208, 142);
                            workSheet.Cells[row, 5, row, 6].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            workSheet.Cells[row, 5, row, 6].Style.Font.Bold = true;
                            workSheet.Cells[row, 5, row, 6].Merge = true;
                            workSheet.Cells[row, 5, row, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            workSheet.Cells[row, 5, row, 6].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            workSheet.Cells[row, 5, row, 6].Value = "Description";
                            row++;

                            var modelOutput = GetDefinition(document.Definitions, response.Schema["ref"].ToString());

                            var index = 0;
                            foreach (var item in modelOutput.Properties)
                            {
                                workSheet.Cells[row, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                workSheet.Cells[row, 1].Value = index + 1;

                                workSheet.Cells[row, 2, row, 3].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                workSheet.Cells[row, 2, row, 3].Merge = true;
                                workSheet.Cells[row, 2, row, 3].Value = item.Key;

                                workSheet.Cells[row, 4].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                workSheet.Cells[row, 4].Value = string.IsNullOrWhiteSpace(item.Value.Format) ? $"{item.Value.Type}" : $"{item.Value.Type} ({item.Value.Format})";

                                workSheet.Cells[row, 5, row, 6].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                workSheet.Cells[row, 5, row, 6].Value = item.Value.Description;
                                workSheet.Row(row).CustomHeight = true;
                                workSheet.Cells[row, 5, row, 6].Merge = true;
                                row++;

                                if (item.Value.Type == "array")
                                {
                                    if (item.Value.Items.ContainsKey("ref"))
                                    {
                                        var model = GetDefinition(document.Definitions, item.Value.Items["ref"]);
                                        foreach (var property in model.Properties)
                                        {
                                            workSheet.Cells[row, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                            workSheet.Cells[row, 1].Value = index + 1;

                                            workSheet.Cells[row, 2, row, 3].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                            workSheet.Cells[row, 2, row, 3].Merge = true;
                                            workSheet.Cells[row, 2, row, 3].Value = property.Key;

                                            workSheet.Cells[row, 4].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                            workSheet.Cells[row, 4].Value = string.IsNullOrWhiteSpace(property.Value.Format) ? $"{property.Value.Type}" : $"{property.Value.Type} ({property.Value.Format})";

                                            workSheet.Cells[row, 5, row, 6].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                            workSheet.Cells[row, 5, row, 6].Value = property.Value.Description;
                                            workSheet.Row(row).CustomHeight = true;
                                            workSheet.Cells[row, 5, row, 6].Merge = true;
                                            row++;
                                            index++;
                                        }
                                    }
                                    else
                                    {
                                        //TODO: có trường hợp là error => có item là type
                                    }
                                }
                                if (item.Value.Ref != null)
                                {
                                    var referenceOutModel = GetDefinition(document.Definitions, item.Value.Ref);
                                    modelOutput.Properties[item.Key].Schema = referenceOutModel.Properties.ToDictionary(x => x.Key, x => (object)x.Value);

                                    foreach (var property in referenceOutModel.Properties)
                                    {
                                        workSheet.Cells[row, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                        workSheet.Cells[row, 1].Value = index + 1;

                                        workSheet.Cells[row, 2, row, 3].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                        workSheet.Cells[row, 2, row, 3].Merge = true;
                                        workSheet.Cells[row, 2, row, 3].Value = $"{item.Key}.{property.Key}";

                                        workSheet.Cells[row, 4].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                        workSheet.Cells[row, 4].Value = string.IsNullOrWhiteSpace(property.Value.Format) ? $"{property.Value.Type}" : $"{property.Value.Type} ({property.Value.Format})";

                                        workSheet.Cells[row, 5, row, 6].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                        workSheet.Cells[row, 5, row, 6].Value = property.Value.Description;
                                        workSheet.Row(row).CustomHeight = true;
                                        workSheet.Cells[row, 5, row, 6].Merge = true;
                                        row++;
                                        index++;
                                    }
                                }
                                else
                                {
                                    index++;
                                }
                            }

                            workSheet.Cells[row, 2, row, 3].Merge = true;
                            workSheet.Cells[row, 2, row, 3].Value = "Output example";
                            row++;
                            workSheet.Cells[row, 3, row, 6].Style.WrapText = true;
                            workSheet.Cells[row, 3, row, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            workSheet.Cells[row, 3, row, 6].Style.Fill.BackgroundColor.SetColor(1, 244, 176, 132);
                            workSheet.Cells[row, 3, row, 6].Merge = true;
                            workSheet.Row(row).CustomHeight = true;
                            var output = DefinitionToJson(document.Definitions, modelOutput);
                            workSheet.Cells[row, 3, row, 6].Value = output;

                            double totalWidth = 0;
                            for (int i = 3; i <= 6; i++)
                            {
                                totalWidth += workSheet.Column(i).Width;
                            }
                            workSheet.Cells[row, 7].Value = output;
                            workSheet.Column(7).Width = totalWidth;
                            workSheet.Row(row).CustomHeight = true;
                            row++;
                        }

                        //Response
                        //nếu response là 1 mảng
                        else if (response.Schema.ContainsKey("type") && response.Schema["type"].ToString() == "array")
                        {
                            row++;
                            workSheet.Cells[row, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            workSheet.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(1, 169, 208, 142);
                            workSheet.Cells[row, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            workSheet.Cells[row, 1].Style.Font.Bold = true;
                            workSheet.Cells[row, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            workSheet.Cells[row, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            workSheet.Cells[row, 1].Value = "STT";
                            workSheet.Cells[row, 2, row, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            workSheet.Cells[row, 2, row, 3].Style.Fill.BackgroundColor.SetColor(1, 169, 208, 142);
                            workSheet.Cells[row, 2, row, 3].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            workSheet.Cells[row, 2, row, 3].Style.Font.Bold = true;
                            workSheet.Cells[row, 2, row, 3].Merge = true;
                            workSheet.Cells[row, 2, row, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            workSheet.Cells[row, 2, row, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            workSheet.Cells[row, 2, row, 3].Value = "Field";
                            workSheet.Cells[row, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            workSheet.Cells[row, 4].Style.Fill.BackgroundColor.SetColor(1, 169, 208, 142);
                            workSheet.Cells[row, 4].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            workSheet.Cells[row, 4].Style.Font.Bold = true;
                            workSheet.Cells[row, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            workSheet.Cells[row, 4].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            workSheet.Cells[row, 4].Value = "Type";
                            workSheet.Cells[row, 5, row, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            workSheet.Cells[row, 5, row, 6].Style.Fill.BackgroundColor.SetColor(1, 169, 208, 142);
                            workSheet.Cells[row, 5, row, 6].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            workSheet.Cells[row, 5, row, 6].Style.Font.Bold = true;
                            workSheet.Cells[row, 5, row, 6].Merge = true;
                            workSheet.Cells[row, 5, row, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            workSheet.Cells[row, 5, row, 6].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            workSheet.Cells[row, 5, row, 6].Value = "Description";
                            row++;

                            DefinitionModel modelOutput = null;
                            var typeResponse = JsonConvert.DeserializeObject<Dictionary<string, string>>(response.Schema["items"].ToString())["ref"];
                            modelOutput = GetDefinition(document.Definitions, typeResponse);

                            var index = 0;
                            foreach (var item in modelOutput.Properties)
                            {
                                workSheet.Cells[row, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                workSheet.Cells[row, 1].Value = index + 1;

                                workSheet.Cells[row, 2, row, 3].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                workSheet.Cells[row, 2, row, 3].Merge = true;
                                workSheet.Cells[row, 2, row, 3].Value = item.Key;

                                workSheet.Cells[row, 4].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                workSheet.Cells[row, 4].Value = string.IsNullOrWhiteSpace(item.Value.Format) ? $"{item.Value.Type}" : $"{item.Value.Type} ({item.Value.Format})";

                                workSheet.Cells[row, 5, row, 6].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                workSheet.Cells[row, 5, row, 6].Value = item.Value.Description;
                                workSheet.Row(row).CustomHeight = true;
                                workSheet.Cells[row, 5, row, 6].Merge = true;
                                row++;

                                if (item.Value.Type == "array")
                                {
                                    if (item.Value.Items.ContainsKey("ref"))
                                    {
                                        var model = GetDefinition(document.Definitions, item.Value.Items["ref"]);
                                        foreach (var property in model.Properties)
                                        {
                                            workSheet.Cells[row, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                            workSheet.Cells[row, 1].Value = index + 1;

                                            workSheet.Cells[row, 2, row, 3].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                            workSheet.Cells[row, 2, row, 3].Merge = true;
                                            workSheet.Cells[row, 2, row, 3].Value = property.Key;

                                            workSheet.Cells[row, 4].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                            workSheet.Cells[row, 4].Value = string.IsNullOrWhiteSpace(property.Value.Format) ? $"{property.Value.Type}" : $"{property.Value.Type} ({property.Value.Format})";

                                            workSheet.Cells[row, 5, row, 6].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                            workSheet.Cells[row, 5, row, 6].Value = property.Value.Description;
                                            workSheet.Row(row).CustomHeight = true;
                                            workSheet.Cells[row, 5, row, 6].Merge = true;
                                            row++;
                                            index++;
                                        }
                                    }
                                    //TODO: có trường hợp là error => có item là type
                                }
                                else
                                {
                                    index++;
                                }
                            }

                            workSheet.Cells[row, 2, row, 3].Merge = true;
                            workSheet.Cells[row, 2, row, 3].Value = "Output example";
                            row++;
                            workSheet.Cells[row, 3, row, 6].Style.WrapText = true;
                            workSheet.Cells[row, 3, row, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            workSheet.Cells[row, 3, row, 6].Style.Fill.BackgroundColor.SetColor(1, 244, 176, 132);
                            workSheet.Cells[row, 3, row, 6].Merge = true;
                            workSheet.Row(row).CustomHeight = true;
                            var output = DefinitionToJson(document.Definitions, modelOutput);
                            workSheet.Cells[row, 3, row, 6].Value = output;

                            double totalWidth = 0;
                            for (int i = 3; i <= 6; i++)
                            {
                                totalWidth += workSheet.Column(i).Width;
                            }
                            workSheet.Cells[row, 7].Value = output;
                            workSheet.Column(7).Width = totalWidth;
                            workSheet.Row(row).CustomHeight = true;
                            row++;
                        }
                    }

                    workSheet.Cells[row, 1, row, 6].Style.Border.Bottom.Style = ExcelBorderStyle.Double;
                    row++;

                    row++;
                    indexApi++;
                }

                workSheet.Column(7).Hidden = true;
                package.Save();
            }

            Console.WriteLine("Done");
            // Console.ReadLine();
        }

        private static DefinitionModel GetDefinition(Dictionary<string, DefinitionModel> definitions, string reference)
        {
            reference = reference.Replace("#/definitions/", "");
            if (!definitions.ContainsKey(reference))
                return null;

            return definitions[reference];
        }

        private static string DefinitionToJson(Dictionary<string, DefinitionModel> definitions, DefinitionModel model)
        {
            if (model.Type == "object")
                return GenerateObject(definitions, model.Properties);

            if (model.Type == "array")
                return GenerateArray(definitions, model.Properties);

            throw new Exception($"Definition type {model.Type.ToUpper()} is not found");
        }

        private static string GenerateObject(Dictionary<string, DefinitionModel> definitions, Dictionary<string, ParameterModel> properties)
        {
            var items = new List<string>();
            foreach (var property in properties)
            {
                var item = $"\"{property.Key}\":";
                item += GenerateExampleValue(property.Value.Type);

                if (property.Value.Type == "array")
                {
                    if (property.Value.Items.ContainsKey("ref"))
                    {
                        var reference = property.Value.Items["ref"].Replace("#/definitions/", "");
                        if (definitions.ContainsKey(reference))
                        {
                            var model = definitions[reference];
                            item += $"[{DefinitionToJson(definitions, model)}]";
                        }
                    }
                    else if (property.Value.Items.ContainsKey("type"))
                    {
                        var typeElement = property.Value.Items["type"];
                        item += $"[{GenerateExampleValue(typeElement)}]";
                    }
                }

                if (property.Value.Ref != null)
                {
                    var reference = property.Value.Ref.Replace("#/definitions/", "");
                    if (definitions.ContainsKey(reference))
                    {
                        var model = definitions[reference];
                        item += $"[{DefinitionToJson(definitions, model)}]";
                    }
                }

                items.Add(item);
            }
            var json = "{" + string.Join(',', items) + "}";
            return json;
        }

        private static string GenerateExampleValue(string type)
        {
            if (type == "string")
                return $"\"string\"";
            if (type == "integer" || type == "number")
                return "0";
            if (type == "boolean")
                return "false";

            return "";
        }

        private static string GenerateArray(Dictionary<string, DefinitionModel> definitions, Dictionary<string, ParameterModel> properties)
        {
            return $"[{GenerateObject(definitions, properties)}]";
        }
    }
}
