﻿using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using Label = Microsoft.Xrm.Sdk.Label;

namespace MsCrmTools.Translator.AppCode
{
    public class AttributeTranslation : BaseTranslation
    {
        public AttributeTranslation()
        {
            name = "Attributes";
        }

        /// <summary>
        ///
        /// </summary>
        /// <example>
        /// attributeId;entityLogicalName;attributeLogicalName;Type;LCID1;LCID2;...;LCODX
        /// </example>
        /// <param name="entities"></param>
        /// <param name="languages"></param>
        /// <param name="sheet"></param>
        /// <param name="settings"></param>
        public void Export(List<EntityMetadata> entities, List<int> languages, ExcelWorksheet sheet, ExportSettings settings)
        {
            var line = 0;
            int cell;

            AddHeader(sheet, languages);

            foreach (var entity in entities.OrderBy(e => e.LogicalName))
            {
                foreach (var attribute in entity.Attributes.OrderBy(a => a.LogicalName))
                {
                    if (attribute.AttributeType == null
                        || attribute.AttributeType.Value == AttributeTypeCode.BigInt
                        || attribute.AttributeType.Value == AttributeTypeCode.CalendarRules
                        || attribute.AttributeType.Value == AttributeTypeCode.EntityName
                        || attribute.AttributeType.Value == AttributeTypeCode.ManagedProperty
                        || attribute.AttributeType.Value == AttributeTypeCode.Uniqueidentifier
                        || attribute.AttributeType.Value == AttributeTypeCode.Virtual && !(attribute is MultiSelectPicklistAttributeMetadata)
                        || attribute.AttributeOf != null
                        || !attribute.MetadataId.HasValue
                        || !attribute.IsRenameable.Value)
                        continue;

                    if (attribute.DisplayName != null && attribute.DisplayName.LocalizedLabels.All(l => string.IsNullOrEmpty(l.Label)))
                        continue;

                    // If derived attribute from calculated field, don't process it
                    if (attribute.LogicalName.EndsWith("_state"))
                    {
                        var baseName = attribute.LogicalName.Remove(attribute.LogicalName.Length - 6, 6);

                        if (entity.Attributes.Any(a => a.LogicalName == baseName) &&
                            entity.Attributes.Any(a => a.LogicalName == baseName + "_date"))
                        {
                            continue;
                        }
                    }
                    if (attribute.LogicalName.EndsWith("_date"))
                    {
                        var baseName = attribute.LogicalName.Remove(attribute.LogicalName.Length - 5, 5);

                        if (entity.Attributes.Any(a => a.LogicalName == baseName) &&
                            entity.Attributes.Any(a => a.LogicalName == baseName + "_state"))
                        {
                            continue;
                        }
                    }

                    if (settings.ExportNames)
                    {
                        // DisplayName
                        line++;
                        cell = 0;
                        ZeroBasedSheet.Cell(sheet, line, cell++).Value = attribute.MetadataId.Value.ToString("B");
                        ZeroBasedSheet.Cell(sheet, line, cell++).Value = entity.LogicalName;
                        ZeroBasedSheet.Cell(sheet, line, cell++).Value = attribute.LogicalName;
                        ZeroBasedSheet.Cell(sheet, line, cell++).Value = "DisplayName";

                        foreach (var lcid in languages)
                        {
                            var displayName = string.Empty;

                            if (attribute.DisplayName != null)
                            {
                                var displayNameLabel =
                                    attribute.DisplayName.LocalizedLabels.FirstOrDefault(l => l.LanguageCode == lcid);
                                if (displayNameLabel != null)
                                {
                                    displayName = displayNameLabel.Label;
                                }
                            }

                            ZeroBasedSheet.Cell(sheet, line, cell++).Value = displayName;
                        }
                    }

                    if (settings.ExportDescriptions)
                    {
                        // Description
                        line++;
                        cell = 0;
                        ZeroBasedSheet.Cell(sheet, line, cell++).Value = attribute.MetadataId.Value.ToString("B");
                        ZeroBasedSheet.Cell(sheet, line, cell++).Value = entity.LogicalName;
                        ZeroBasedSheet.Cell(sheet, line, cell++).Value = attribute.LogicalName;
                        ZeroBasedSheet.Cell(sheet, line, cell++).Value = "Description";

                        foreach (var lcid in languages)
                        {
                            var description = string.Empty;

                            if (attribute.Description != null)
                            {
                                var descriptionLabel = attribute.Description.LocalizedLabels.FirstOrDefault(l => l.LanguageCode == lcid);
                                if (descriptionLabel != null)
                                {
                                    description = descriptionLabel.Label;
                                }
                            }

                            ZeroBasedSheet.Cell(sheet, line, cell++).Value = description;
                        }
                    }
                }
            }

            // Applying style to cells
            for (int i = 0; i < (4 + languages.Count); i++)
            {
                StyleMutator.TitleCell(ZeroBasedSheet.Cell(sheet, 0, i).Style);
            }

            for (int i = 1; i <= line; i++)
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

            var amds = new List<MasterAttribute>();

            var rowsCount = sheet.Dimension.Rows;
            var cellsCount = sheet.Dimension.Columns;
            for (var rowI = 1; rowI < rowsCount; rowI++)
            {
                if (HasEmptyCells(sheet, rowI, 3)) continue;

                var amd = amds.FirstOrDefault(a => a.Amd.MetadataId == new Guid(ZeroBasedSheet.Cell(sheet, rowI, 0).Value?.ToString()));
                if (amd == null)
                {
                    var currentEntity = emds.FirstOrDefault(e => e.LogicalName == ZeroBasedSheet.Cell(sheet, rowI, 1).Value?.ToString());
                    if (currentEntity == null)
                    {
                        var request = new RetrieveEntityRequest
                        {
                            LogicalName = ZeroBasedSheet.Cell(sheet, rowI, 1).Value.ToString(),
                            EntityFilters = EntityFilters.Entity | EntityFilters.Attributes | EntityFilters.Relationships
                        };

                        var response = ((RetrieveEntityResponse)service.Execute(request));
                        currentEntity = response.EntityMetadata;

                        emds.Add(currentEntity);
                    }

                    amd = new MasterAttribute();
                    amd.Amd = currentEntity.Attributes.FirstOrDefault(a => string.Equals(a.LogicalName,
                        ZeroBasedSheet.Cell(sheet, rowI, 2).Value?.ToString(), StringComparison.OrdinalIgnoreCase));

                    //still null? someone deleted the attribute while we were busy translating. Let's skip it!
                    if (amd.Amd == null)
                    {
                        OnLog(new LogEventArgs
                        {
                            Type = LogType.Warning,
                            Message = $"Attribute {ZeroBasedSheet.Cell(sheet, rowI, 1).Value} - {ZeroBasedSheet.Cell(sheet, rowI, 2).Value} is missing in CRM!"
                        });
                        continue;
                    }

                    amds.Add(amd);
                }

                int columnIndex = 4;

                if (ZeroBasedSheet.Cell(sheet, rowI, 3).Value.ToString() == "DisplayName")
                {
                    if (amd.Amd.DisplayName == null) amd.Amd.DisplayName = new Label();

                    while (columnIndex < cellsCount)
                    {
                        if (ZeroBasedSheet.Cell(sheet, rowI, columnIndex).Value != null)
                        {
                            var lcid = int.Parse(ZeroBasedSheet.Cell(sheet, 0, columnIndex).Value.ToString());
                            var label = ZeroBasedSheet.Cell(sheet, rowI, columnIndex).Value.ToString();

                            var translatedLabel = amd.Amd.DisplayName.LocalizedLabels.FirstOrDefault(x => x.LanguageCode == lcid);
                            if (translatedLabel == null)
                            {
                                translatedLabel = new LocalizedLabel(label, lcid);
                                amd.Amd.DisplayName.LocalizedLabels.Add(translatedLabel);
                            }
                            else
                            {
                                translatedLabel.Label = label;
                            }
                        }
                        columnIndex++;
                    }
                }
                else if (ZeroBasedSheet.Cell(sheet, rowI, 3).Value.ToString() == "Description")
                {
                    amd.Amd.Description = new Label();

                    while (columnIndex < cellsCount)
                    {
                        if (ZeroBasedSheet.Cell(sheet, rowI, columnIndex).Value != null)
                        {
                            var lcid = int.Parse(ZeroBasedSheet.Cell(sheet, 0, columnIndex).Value.ToString());
                            var label = ZeroBasedSheet.Cell(sheet, rowI, columnIndex).Value.ToString();
                            var translatedLabel = amd.Amd.Description.LocalizedLabels.FirstOrDefault(x => x.LanguageCode == lcid);
                            if (translatedLabel == null)
                            {
                                translatedLabel = new LocalizedLabel(label, lcid);
                                amd.Amd.Description.LocalizedLabels.Add(translatedLabel);
                            }
                            else
                            {
                                translatedLabel.Label = label;
                            }
                        }

                        columnIndex++;
                    }
                }
            }

            int i = 0;

            OnLog(new LogEventArgs($"Importing {sheet.Name} translations"));

            var arg = new TranslationProgressEventArgs();
            foreach (var amd in amds)
            {
                if (amd.Amd.DisplayName.LocalizedLabels.All(l => string.IsNullOrEmpty(l.Label))
                    || amd.Amd.IsRenameable.Value == false)
                {
                    continue;
                }

                AddRequest(new UpdateAttributeRequest
                {
                    Attribute = amd.Amd,
                    EntityName = amd.Amd.EntityLogicalName,
                    MergeLabels = true
                });

                ExecuteMultiple(service, arg, amds.Count);
            }
            ExecuteMultiple(service, arg, amds.Count, true);
        }

        private void AddHeader(ExcelWorksheet sheet, IEnumerable<int> languages)
        {
            var cell = 0;

            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "Attribute Id";
            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "Entity Logical Name";
            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "Attribute Logical Name";
            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "Type";

            foreach (var lcid in languages)
            {
                ZeroBasedSheet.Cell(sheet, 0, cell++).Value = lcid.ToString(CultureInfo.InvariantCulture);
            }
        }
    }
}