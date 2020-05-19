using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Linq;

namespace MsCrmTools.Translator.AppCode
{
    internal class RelationshipNnTranslation : BaseTranslation
    {
        public RelationshipNnTranslation()
        {
            name = "NN Relationships";
        }

        public void Export(List<EntityMetadata> entities, List<int> languages, ExcelWorksheet sheet)
        {
            var line = 1;

            AddHeader(sheet, languages);

            foreach (var entity in entities.OrderBy(e => e.LogicalName))
            {
                foreach (var rel in entity.ManyToManyRelationships.ToList())
                {
                    var cell = 0;

                    var amc = rel.Entity1LogicalName == entity.LogicalName ? rel.Entity1AssociatedMenuConfiguration : rel.Entity2AssociatedMenuConfiguration;

                    if (!(amc.Behavior.HasValue && amc.Behavior.Value == AssociatedMenuBehavior.UseLabel))
                        continue;

                    ZeroBasedSheet.Cell(sheet, line, cell++).Value = entity.LogicalName;
                    ZeroBasedSheet.Cell(sheet, line, cell++).Value = rel.MetadataId.Value.ToString("B");
                    ZeroBasedSheet.Cell(sheet, line, cell++).Value = rel.IntersectEntityName;

                    foreach (var lcid in languages)
                    {
                        var entity1Label = string.Empty;

                        if (amc.Label != null)
                        {
                            var displayNameLabel =
                                amc.Label.LocalizedLabels.FirstOrDefault(l => l.LanguageCode == lcid);
                            if (displayNameLabel != null)
                            {
                                entity1Label = displayNameLabel.Label;
                            }
                        }

                        ZeroBasedSheet.Cell(sheet, line, cell++).Value = entity1Label;
                    }

                    line++;
                }
            }

            // Applying style to cells
            for (int i = 0; i < (3 + languages.Count); i++)
            {
                StyleMutator.TitleCell(ZeroBasedSheet.Cell(sheet, 0, i).Style);
            }

            for (int i = 1; i < line; i++)
            {
                for (int j = 0; j < 3; j++)
                {
                    StyleMutator.HighlightedCell(ZeroBasedSheet.Cell(sheet, i, j).Style);
                }
            }
        }

        public void Import(ExcelWorksheet sheet, List<EntityMetadata> emds, IOrganizationService service, BackgroundWorker worker)
        {
            var rmds = new List<ManyToManyRelationshipMetadata>();

            var rowsCount = sheet.Dimension.Rows;
            var cellsCount = sheet.Dimension.Columns;
            for (var rowI = 1; rowI < rowsCount; rowI++)
            {
                var rmd = rmds.FirstOrDefault(r => r.MetadataId == new Guid(ZeroBasedSheet.Cell(sheet, rowI, 1).Value.ToString()));
                if (rmd == null)
                {
                    var currentEntity = emds.FirstOrDefault(e => e.LogicalName == ZeroBasedSheet.Cell(sheet, rowI, 0).Value.ToString());
                    if (currentEntity == null)
                    {
                        var request = new RetrieveEntityRequest
                        {
                            LogicalName = ZeroBasedSheet.Cell(sheet, rowI, 0).Value.ToString(),
                            EntityFilters = EntityFilters.Entity | EntityFilters.Attributes | EntityFilters.Relationships
                        };

                        var response = ((RetrieveEntityResponse)service.Execute(request));
                        currentEntity = response.EntityMetadata;

                        emds.Add(currentEntity);
                    }

                    rmd = currentEntity.ManyToManyRelationships.FirstOrDefault(r => r.IntersectEntityName == ZeroBasedSheet.Cell(sheet, rowI, 2).Value.ToString());
                    rmds.Add(rmd);
                }

                if (rmd == null)
                {
                    OnLog(new LogEventArgs
                    {
                        Type = LogType.Warning,
                        Message = $"Unable to find relationship '{ZeroBasedSheet.Cell(sheet, rowI, 2).Value}' for entity '{ZeroBasedSheet.Cell(sheet, rowI, 0).Value}"
                    });
                    continue;
                }

                int columnIndex = 3;

                if (rmd.Entity1LogicalName == ZeroBasedSheet.Cell(sheet, rowI, 0).Value.ToString())
                {
                    rmd.Entity1AssociatedMenuConfiguration.Label = new Label();

                    while (columnIndex < cellsCount)
                    {
                        if (ZeroBasedSheet.Cell(sheet, rowI, columnIndex).Value != null)
                        {
                            var lcid = int.Parse(ZeroBasedSheet.Cell(sheet, 0, columnIndex).Value.ToString());
                            var label = ZeroBasedSheet.Cell(sheet, rowI, columnIndex).Value.ToString();

                            rmd.Entity1AssociatedMenuConfiguration.Label.LocalizedLabels.Add(new LocalizedLabel(label, lcid));
                        }

                        columnIndex++;
                    }
                }
                else if (rmd.Entity2LogicalName == ZeroBasedSheet.Cell(sheet, rowI, 0).Value.ToString())
                {
                    rmd.Entity2AssociatedMenuConfiguration.Label = new Label();

                    while (columnIndex < cellsCount)
                    {
                        if (ZeroBasedSheet.Cell(sheet, rowI, columnIndex).Value != null)
                        {
                            var lcid = int.Parse(ZeroBasedSheet.Cell(sheet, 0, columnIndex).Value.ToString());
                            var label = ZeroBasedSheet.Cell(sheet, rowI, columnIndex).Value.ToString();

                            rmd.Entity2AssociatedMenuConfiguration.Label.LocalizedLabels.Add(new LocalizedLabel(label,
                                lcid));
                        }
                        columnIndex++;
                    }
                }
                else
                {
                    OnLog(new LogEventArgs
                    {
                        Type = LogType.Warning,
                        Message = $"Unable to find entity '{ZeroBasedSheet.Cell(sheet, rowI, 0).Value}' in relationship '{ZeroBasedSheet.Cell(sheet, rowI, 2).Value}"
                    });
                }
            }

            var arg = new TranslationProgressEventArgs { SheetName = sheet.Name };
            foreach (var rmd in rmds)
            {
                var request = new UpdateRelationshipRequest
                {
                    Relationship = rmd,
                    MergeLabels = true
                };

                AddRequest(request);
                ExecuteMultiple(service, arg);
            }
            ExecuteMultiple(service, arg, true);
        }

        private void AddHeader(ExcelWorksheet sheet, IEnumerable<int> languages)
        {
            var cell = 0;

            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "Entity";
            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "Relationship Id";
            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "Relationship Intersect Entity";

            foreach (var lcid in languages)
            {
                ZeroBasedSheet.Cell(sheet, 0, cell++).Value = lcid.ToString(CultureInfo.InvariantCulture);
            }
        }
    }
}