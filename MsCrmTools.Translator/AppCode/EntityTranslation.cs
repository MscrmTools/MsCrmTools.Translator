﻿using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using OfficeOpenXml;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Linq;

namespace MsCrmTools.Translator.AppCode
{
    public class EntityTranslation : BaseTranslation
    {
        /// <summary>
        ///
        /// </summary>
        /// <example>
        /// entityId;entityLogicalName;Type;LCID1;LCID2;...;LCODX
        /// </example>
        /// <param name="entities"></param>
        /// <param name="languages"></param>
        /// <param name="sheet"></param>
        public void Export(List<EntityMetadata> entities, List<int> languages, ExcelWorksheet sheet, ExportSettings settings)
        {
            var line = 0;
            var cell = 0;

            AddHeader(sheet, languages);

            foreach (var entity in entities.OrderBy(e => e.LogicalName))
            {
                if (!entity.MetadataId.HasValue)
                    continue;

                if (settings.ExportNames)
                {
                    line++;
                    cell = 0;

                    ZeroBasedSheet.Cell(sheet, line, cell++).Value = entity.MetadataId.Value.ToString("B");
                    ZeroBasedSheet.Cell(sheet, line, cell++).Value = entity.LogicalName;

                    // DisplayName
                    ZeroBasedSheet.Cell(sheet, line, cell++).Value = "DisplayName";

                    foreach (var lcid in languages)
                    {
                        var displayName = string.Empty;

                        if (entity.DisplayName != null)
                        {
                            var displayNameLabel =
                                entity.DisplayName.LocalizedLabels.FirstOrDefault(l => l.LanguageCode == lcid);
                            if (displayNameLabel != null)
                            {
                                displayName = displayNameLabel.Label;
                            }
                        }

                        ZeroBasedSheet.Cell(sheet, line, cell++).Value = displayName;
                    }

                    // Plural Name
                    line++;
                    cell = 0;
                    ZeroBasedSheet.Cell(sheet, line, cell++).Value = entity.MetadataId.Value.ToString("B");
                    ZeroBasedSheet.Cell(sheet, line, cell++).Value = entity.LogicalName;
                    ZeroBasedSheet.Cell(sheet, line, cell++).Value = "DisplayCollectionName";

                    foreach (var lcid in languages)
                    {
                        var collectionName = string.Empty;

                        if (entity.DisplayCollectionName != null)
                        {
                            var collectionNameLabel =
                                entity.DisplayCollectionName.LocalizedLabels.FirstOrDefault(l =>
                                    l.LanguageCode == lcid);
                            if (collectionNameLabel != null)
                            {
                                collectionName = collectionNameLabel.Label;
                            }
                        }

                        ZeroBasedSheet.Cell(sheet, line, cell++).Value = collectionName;
                    }
                }

                if (settings.ExportDescriptions)
                {
                    // Description
                    line++;
                    cell = 0;
                    ZeroBasedSheet.Cell(sheet, line, cell++).Value = entity.MetadataId.Value.ToString("B");
                    ZeroBasedSheet.Cell(sheet, line, cell++).Value = entity.LogicalName;
                    ZeroBasedSheet.Cell(sheet, line, cell++).Value = "Description";

                    foreach (var lcid in languages)
                    {
                        var description = string.Empty;

                        if (entity.Description != null)
                        {
                            var descriptionLabel =
                                entity.Description.LocalizedLabels.FirstOrDefault(l => l.LanguageCode == lcid);
                            if (descriptionLabel != null)
                            {
                                description = descriptionLabel.Label;
                            }
                        }

                        ZeroBasedSheet.Cell(sheet, line, cell++).Value = description;
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

        public void Import(ExcelWorksheet sheet, List<EntityMetadata> emds, IOrganizationService service, BackgroundWorker worker)
        {
            var rowsCount = sheet.Dimension.Rows;
            var cellsCount = sheet.Dimension.Columns;

            for (var rowI = 1; rowI < rowsCount; rowI++)
            {
                var emd = emds.FirstOrDefault(e => e.LogicalName == ZeroBasedSheet.Cell(sheet, rowI, 1).Value.ToString());
                if (emd == null)
                {
                    var request = new RetrieveEntityRequest
                    {
                        LogicalName = ZeroBasedSheet.Cell(sheet, rowI, 1).Value.ToString(),
                        EntityFilters = EntityFilters.Entity | EntityFilters.Attributes | EntityFilters.Relationships
                    };

                    var response = ((RetrieveEntityResponse)service.Execute(request));
                    emd = response.EntityMetadata;

                    OnResult(new TranslationResultEventArgs
                    {
                        Success = true,
                        SheetName = sheet.Name,
                        Message = $"Entity: {emd.LogicalName}"
                    });

                    emds.Add(emd);
                }

                if (ZeroBasedSheet.Cell(sheet, rowI, 2).Value.ToString() == "DisplayName")
                {
                    emd.DisplayName = new Label();
                    int columnIndex = 3;

                    while (columnIndex < cellsCount)
                    {
                        if (ZeroBasedSheet.Cell(sheet, rowI, columnIndex).Value != null)
                        {
                            var lcid = int.Parse(ZeroBasedSheet.Cell(sheet, 0, columnIndex).Value.ToString());
                            var label = ZeroBasedSheet.Cell(sheet, rowI, columnIndex).Value.ToString();

                            emd.DisplayName.LocalizedLabels.Add(new LocalizedLabel(label, lcid));
                        }

                        columnIndex++;
                    }
                }
                else if (ZeroBasedSheet.Cell(sheet, rowI, 2).Value.ToString() == "DisplayCollectionName")
                {
                    emd.DisplayCollectionName = new Label();
                    int columnIndex = 3;

                    while (columnIndex < cellsCount)
                    {
                        if (ZeroBasedSheet.Cell(sheet, rowI, columnIndex).Value != null)
                        {
                            var lcid = int.Parse(ZeroBasedSheet.Cell(sheet, 0, columnIndex).Value.ToString());
                            var label = ZeroBasedSheet.Cell(sheet, rowI, columnIndex).Value.ToString();

                            emd.DisplayCollectionName.LocalizedLabels.Add(new LocalizedLabel(label, lcid));
                        }

                        columnIndex++;
                    }
                }
                else if (ZeroBasedSheet.Cell(sheet, rowI, 2).Value.ToString() == "Description")
                {
                    emd.Description = new Label();
                    int columnIndex = 3;

                    while (columnIndex < cellsCount)
                    {
                        if (ZeroBasedSheet.Cell(sheet, rowI, columnIndex).Value != null)
                        {
                            var lcid = int.Parse(ZeroBasedSheet.Cell(sheet, 0, columnIndex).Value.ToString());
                            var label = ZeroBasedSheet.Cell(sheet, rowI, columnIndex).Value.ToString();

                            emd.Description.LocalizedLabels.Add(new LocalizedLabel(label, lcid));
                        }

                        columnIndex++;
                    }
                }
            }

            var entities = emds.Where(e => e.IsRenameable.Value).ToList();
            int i = 0;
            foreach (var emd in entities)
            {
                var entityUpdate = new EntityMetadata();
                entityUpdate.LogicalName = emd.LogicalName;
                entityUpdate.DisplayName = emd.DisplayName;
                entityUpdate.Description = emd.Description;
                entityUpdate.DisplayCollectionName = emd.DisplayCollectionName;

                try
                {
                    var request = new UpdateEntityRequest { Entity = entityUpdate };

                    service.Execute(request);

                    OnResult(new TranslationResultEventArgs
                    {
                        Success = true,
                        SheetName = sheet.Name
                    });
                }
                catch (Exception error)
                {
                    OnResult(new TranslationResultEventArgs
                    {
                        Success = false,
                        SheetName = sheet.Name,
                        Message = $"{emd.LogicalName}: {error.Message}"
                    });
                }

                i++;
                worker.ReportProgressIfPossible(0, new ProgressInfo
                {
                    Item = i * 100 / entities.Count
                });
            }
        }

        private void AddHeader(ExcelWorksheet sheet, IEnumerable<int> languages)
        {
            var cell = 0;

            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "Entity Id";
            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "Entity Logical Name";
            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "Type";

            foreach (var lcid in languages)
            {
                ZeroBasedSheet.Cell(sheet, 0, cell++).Value = lcid.ToString(CultureInfo.InvariantCulture);
            }
        }
    }
}