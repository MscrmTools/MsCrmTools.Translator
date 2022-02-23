using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Linq;

namespace MsCrmTools.Translator.AppCode
{
    public class VisualizationTranslation : BaseTranslation
    {
        public VisualizationTranslation()
        {
            name = "Charts";
        }

        /// <summary>
        ///
        /// </summary>
        /// <example>
        /// visualizationId;entityLogicalName;visualizationName;LCID1;LCID2;...;LCODX
        /// </example>
        /// <param name="entities"></param>
        /// <param name="languages"></param>
        /// <param name="sheet"></param>
        /// <param name="service"></param>
        /// <param name="settings"></param>
        public void Export(List<EntityMetadata> entities, List<int> languages, ExcelWorksheet sheet, IOrganizationService service, ExportSettings settings)
        {
            var line = 0;
            int cell;

            AddHeader(sheet, languages);

            var crmVisualizations = new List<CrmVisualization>();

            foreach (var entity in entities.OrderBy(e => e.LogicalName))
            {
                if (!entity.MetadataId.HasValue)
                    continue;

                var visualizations = RetrieveVisualizations(entity.ObjectTypeCode.Value, service);

                foreach (var visualization in visualizations)
                {
                    var crmVisualization = crmVisualizations.FirstOrDefault(cv => cv.Id == visualization.Id);
                    if (crmVisualization == null)
                    {
                        crmVisualization = new CrmVisualization
                        {
                            Id = visualization.Id,
                            Entity = visualization.GetAttributeValue<string>("primaryentitytypecode"),
                            Names = new Dictionary<int, string>(),
                            Descriptions = new Dictionary<int, string>()
                        };
                        crmVisualizations.Add(crmVisualization);
                    }

                    RetrieveLocLabelsRequest request;
                    RetrieveLocLabelsResponse response;

                    if (settings.ExportNames)
                    {
                        // Names
                        request = new RetrieveLocLabelsRequest
                        {
                            AttributeName = "name",
                            EntityMoniker = new EntityReference("savedqueryvisualization", visualization.Id)
                        };

                        response = (RetrieveLocLabelsResponse)service.Execute(request);
                        foreach (var locLabel in response.Label.LocalizedLabels)
                        {
                            crmVisualization.Names.Add(locLabel.LanguageCode, locLabel.Label);
                        }
                    }

                    if (settings.ExportDescriptions)
                    {
                        // Descriptions
                        request = new RetrieveLocLabelsRequest
                        {
                            AttributeName = "description",
                            EntityMoniker = new EntityReference("savedqueryvisualization", visualization.Id)
                        };

                        response = (RetrieveLocLabelsResponse)service.Execute(request);
                        foreach (var locLabel in response.Label.LocalizedLabels)
                        {
                            crmVisualization.Descriptions.Add(locLabel.LanguageCode, locLabel.Label);
                        }
                    }
                }
            }

            foreach (var crmVisualization in crmVisualizations.OrderBy(cv => cv.Entity))
            {
                if (settings.ExportNames)
                {
                    line++;
                    cell = 0;
                    ZeroBasedSheet.Cell(sheet, line, cell++).Value = crmVisualization.Id.ToString("B");
                    ZeroBasedSheet.Cell(sheet, line, cell++).Value = crmVisualization.Entity;
                    ZeroBasedSheet.Cell(sheet, line, cell++).Value = "Name";

                    foreach (var lcid in languages)
                    {
                        var name = crmVisualization.Names.FirstOrDefault(n => n.Key == lcid);
                        if (name.Value != null)
                            ZeroBasedSheet.Cell(sheet, line, cell++).Value = name.Value;
                        else
                        {
                            cell++;
                        }
                    }
                }

                if (settings.ExportDescriptions)
                {
                    line++;
                    cell = 0;
                    ZeroBasedSheet.Cell(sheet, line, cell++).Value = crmVisualization.Id.ToString("B");
                    ZeroBasedSheet.Cell(sheet, line, cell++).Value = crmVisualization.Entity;
                    ZeroBasedSheet.Cell(sheet, line, cell++).Value = "Description";

                    foreach (var lcid in languages)
                    {
                        var desc = crmVisualization.Descriptions.FirstOrDefault(n => n.Key == lcid);
                        if (desc.Value != null)
                            ZeroBasedSheet.Cell(sheet, line, cell++).Value = desc.Value;
                        else
                        {
                            cell++;
                        }
                    }
                }
            }

            // Applying style to cells
            for (int i = 0; i < (3 + languages.Count); i++)
            {
                StyleMutator.TitleCell(ZeroBasedSheet.Cell(sheet, 0, i).Style);
            }

            for (int i = 1; i <= line; i++)
            {
                for (int j = 0; j < 3; j++)
                {
                    StyleMutator.HighlightedCell(ZeroBasedSheet.Cell(sheet, i, j).Style);
                }
            }
        }

        public void Import(ExcelWorksheet sheet, IOrganizationService service, BackgroundWorker worker)
        {
            OnLog(new LogEventArgs($"Reading {sheet.Name}"));

            var rowsCount = sheet.Dimension.Rows;
            var cellsCount = sheet.Dimension.Columns;
            var requests = new List<SetLocLabelsRequest>();
            for (var rowI = 1; rowI < rowsCount; rowI++)
            {
                if (HasEmptyCells(sheet, rowI, 2)) continue;

                var currentVisualizationId = new Guid(ZeroBasedSheet.Cell(sheet, rowI, 0).Value.ToString());

                var locLabel = ((RetrieveLocLabelsResponse)service.Execute(new RetrieveLocLabelsRequest
                {
                    EntityMoniker = new EntityReference("savedqueryvisualization", currentVisualizationId),
                    AttributeName = ZeroBasedSheet.Cell(sheet, rowI, 3).Value.ToString() == "Name" ? "name" : "description"
                })).Label;

                var request = new SetLocLabelsRequest
                {
                    EntityMoniker = new EntityReference("savedqueryvisualization", currentVisualizationId),
                    AttributeName = ZeroBasedSheet.Cell(sheet, rowI, 2).Value.ToString() == "Name" ? "name" : "description"
                };

                var labels = locLabel.LocalizedLabels.ToList();

                var columnIndex = 3;
                while (columnIndex < cellsCount)
                {
                    if (ZeroBasedSheet.Cell(sheet, rowI, columnIndex).Value != null)
                    {
                        var lcid = int.Parse(ZeroBasedSheet.Cell(sheet, 0, columnIndex).Value.ToString());
                        var label = ZeroBasedSheet.Cell(sheet, rowI, columnIndex).Value.ToString();

                        var translatedLabel = labels.FirstOrDefault(x => x.LanguageCode == lcid);
                        if (translatedLabel == null)
                        {
                            translatedLabel = new LocalizedLabel(label, lcid);
                            labels.Add(translatedLabel);
                        }
                        else
                        {
                            translatedLabel.Label = label;
                        }
                    }

                    columnIndex++;
                }

                request.Labels = labels.ToArray();

                requests.Add(request);
            }

            OnLog(new LogEventArgs($"Importing {sheet.Name} translations"));

            var arg = new TranslationProgressEventArgs { SheetName = sheet.Name };
            foreach (var request in requests)
            {
                AddRequest(request);
                ExecuteMultiple(service, arg, requests.Count);
            }
            ExecuteMultiple(service, arg, requests.Count, true);
        }

        private void AddHeader(ExcelWorksheet sheet, IEnumerable<int> languages)
        {
            var cell = 0;

            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "Chart Id";
            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "Entity Logical Name";
            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "Type";

            foreach (var lcid in languages)
            {
                ZeroBasedSheet.Cell(sheet, 0, cell++).Value = lcid.ToString(CultureInfo.InvariantCulture);
            }
        }

        private List<Entity> RetrieveVisualizations(int objectTypeCode, IOrganizationService service)
        {
            try
            {
                var qba = new QueryByAttribute
                {
                    EntityName = "savedqueryvisualization",
                    ColumnSet = new ColumnSet(true)
                };

                qba.Attributes.Add("primaryentitytypecode");
                qba.Values.Add(objectTypeCode);

                EntityCollection visualizations = service.RetrieveMultiple(qba);

                return visualizations.Entities.ToList();
            }
            catch (Exception error)
            {
                throw new Exception("Error while retrieving visualizations: " + error.Message);
            }
        }
    }
}