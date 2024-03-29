﻿using Microsoft.Xrm.Sdk;
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
    internal class RelationshipTranslation : BaseTranslation
    {
        public RelationshipTranslation()
        {
            name = "Relationships";
        }

        public void Export(List<EntityMetadata> entities, List<int> languages, ExcelWorksheet sheet)
        {
            var line = 1;

            AddHeader(sheet, languages);
            var exportedRelationships = new List<Guid>();

            foreach (var entity in entities.OrderBy(e => e.LogicalName))
            {
                var relationships = new List<OneToManyRelationshipMetadata>();
                relationships.AddRange(entity.OneToManyRelationships);
                relationships.AddRange(entity.ManyToOneRelationships);

                foreach (var rel in relationships)
                {
                    if (exportedRelationships.Contains(rel.MetadataId.Value))
                    {
                        continue;
                    }
                    exportedRelationships.Add(rel.MetadataId.Value);

                    var cell = 0;

                    if (!rel.AssociatedMenuConfiguration.Behavior.HasValue ||
                         rel.AssociatedMenuConfiguration.Behavior.Value != AssociatedMenuBehavior.UseLabel)
                        continue;

                    // entity1Label
                    ZeroBasedSheet.Cell(sheet, line, cell++).Value = rel.ReferencedEntity;
                    ZeroBasedSheet.Cell(sheet, line, cell++).Value = rel.MetadataId.Value.ToString("B");
                    ZeroBasedSheet.Cell(sheet, line, cell++).Value = rel.SchemaName;
                    ZeroBasedSheet.Cell(sheet, line, cell++).Value = rel.ReferencingEntity;

                    foreach (var lcid in languages)
                    {
                        var entity1Label = string.Empty;

                        if (rel.AssociatedMenuConfiguration.Label != null)
                        {
                            var displayNameLabel =
                                rel.AssociatedMenuConfiguration.Label.LocalizedLabels.FirstOrDefault(
                                    l => l.LanguageCode == lcid);
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
            for (int i = 0; i < (4 + languages.Count); i++)
            {
                StyleMutator.TitleCell(ZeroBasedSheet.Cell(sheet, 0, i).Style);
            }

            for (int i = 1; i < line; i++)
            {
                for (int j = 0; j < 4; j++)
                {
                    StyleMutator.HighlightedCell(ZeroBasedSheet.Cell(sheet, i, j).Style);
                }
            }
        }

        public void Import(ExcelWorksheet sheet, List<EntityMetadata> emds, IOrganizationService service, BackgroundWorker worker)
        {
            OnLog(new LogEventArgs($"Reading {sheet.Name}"));

            var rmds = new List<OneToManyRelationshipMetadata>();

            var rowsCount = sheet.Dimension.Rows;
            var cellsCount = sheet.Dimension.Columns;
            for (var rowI = 1; rowI < rowsCount; rowI++)
            {
                if (HasEmptyCells(sheet, rowI, 3)) continue;

                var rmd = rmds.FirstOrDefault(r => ZeroBasedSheet.Cell(sheet, rowI, 1).Value != null && r.MetadataId == new Guid(ZeroBasedSheet.Cell(sheet, rowI, 1).Value.ToString()));
                if (rmd == null)
                {
                    var currentEntity = emds.FirstOrDefault(e => e.LogicalName == ZeroBasedSheet.Cell(sheet, rowI, 0).Value?.ToString());
                    if (currentEntity == null)
                    {
                        var request = new RetrieveEntityRequest
                        {
                            LogicalName = ZeroBasedSheet.Cell(sheet, rowI, 0).Value.ToString(),
                            EntityFilters = EntityFilters.Entity | EntityFilters.Attributes | EntityFilters.Relationships
                        };

                        var response = (RetrieveEntityResponse)service.Execute(request);
                        currentEntity = response.EntityMetadata;

                        emds.Add(currentEntity);
                    }
                    rmd =
                        currentEntity.OneToManyRelationships.FirstOrDefault(
                            r => r.SchemaName == ZeroBasedSheet.Cell(sheet, rowI, 2).Value?.ToString());
                    if (rmd == null)
                    {
                        rmd =
                            currentEntity.ManyToOneRelationships.FirstOrDefault(
                                r => r.SchemaName == ZeroBasedSheet.Cell(sheet, rowI, 2).Value?.ToString());
                    }

                    rmds.Add(rmd);
                }

                int columnIndex = 4;

                if (rmd.AssociatedMenuConfiguration.Label == null) rmd.AssociatedMenuConfiguration.Label = new Label();

                while (columnIndex < cellsCount)
                {
                    if (ZeroBasedSheet.Cell(sheet, rowI, columnIndex).Value != null)
                    {
                        var lcid = int.Parse(ZeroBasedSheet.Cell(sheet, 0, columnIndex).Value.ToString());
                        var label = ZeroBasedSheet.Cell(sheet, rowI, columnIndex).Value.ToString();

                        var translatedLabel = rmd.AssociatedMenuConfiguration.Label.LocalizedLabels.FirstOrDefault(x => x.LanguageCode == lcid);
                        if (translatedLabel == null)
                        {
                            translatedLabel = new LocalizedLabel(label, lcid);
                            rmd.AssociatedMenuConfiguration.Label.LocalizedLabels.Add(translatedLabel);
                        }
                        else
                        {
                            translatedLabel.Label = label;
                        }
                    }

                    columnIndex++;
                }
            }

            OnLog(new LogEventArgs($"Importing {sheet.Name} translations"));

            var arg = new TranslationProgressEventArgs { SheetName = sheet.Name };
            foreach (var rmd in rmds)
            {
                var request = new UpdateRelationshipRequest
                {
                    Relationship = rmd,
                    MergeLabels = true
                };

                AddRequest(request);
                ExecuteMultiple(service, arg, rmds.Count);
            }

            ExecuteMultiple(service, arg, rmds.Count, true);
        }

        private void AddHeader(ExcelWorksheet sheet, IEnumerable<int> languages)
        {
            var cell = 0;

            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "Entity";
            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "Relationship Id";
            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "Relationship Name";
            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "Relationship entity";

            foreach (var lcid in languages)
            {
                ZeroBasedSheet.Cell(sheet, 0, cell++).Value = lcid.ToString(CultureInfo.InvariantCulture);
            }
        }
    }
}