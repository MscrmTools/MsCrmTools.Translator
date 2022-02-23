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
    public class BooleanTranslation : BaseTranslation
    {
        public BooleanTranslation()
        {
            name = "Boolean values";
        }

        /// <summary>
        ///
        /// </summary>
        /// <example>
        /// attributeId;entityLogicalName;attributeLogicalName;OptionSetValue;LCID1;LCID2;...;LCODX
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
                        || attribute.AttributeType.Value != AttributeTypeCode.Boolean
                        || !attribute.MetadataId.HasValue)
                        continue;

                    var bAmd = (BooleanAttributeMetadata)attribute;

                    if (bAmd.OptionSet?.IsGlobal ?? false)
                        continue;

                    if (settings.ExportNames)
                    {
                        line++;
                        cell = 0;

                        ZeroBasedSheet.Cell(sheet, line, cell++).Value = attribute.MetadataId.Value.ToString("B");
                        ZeroBasedSheet.Cell(sheet, line, cell++).Value = entity.LogicalName;
                        ZeroBasedSheet.Cell(sheet, line, cell++).Value = attribute.LogicalName;
                        ZeroBasedSheet.Cell(sheet, line, cell++).Value = bAmd.OptionSet.FalseOption.Value;
                        ZeroBasedSheet.Cell(sheet, line, cell++).Value = "Label";

                        foreach (var lcid in languages)
                        {
                            var label = string.Empty;

                            if (bAmd.OptionSet.FalseOption.Label != null)
                            {
                                var optionLabel =
                                    bAmd.OptionSet.FalseOption.Label.LocalizedLabels.FirstOrDefault(l =>
                                        l.LanguageCode == lcid);
                                if (optionLabel != null)
                                {
                                    label = optionLabel.Label;
                                }
                            }

                            ZeroBasedSheet.Cell(sheet, line, cell++).Value = label;
                        }
                    }

                    if (settings.ExportDescriptions)
                    {
                        line++;
                        cell = 0;

                        ZeroBasedSheet.Cell(sheet, line, cell++).Value = attribute.MetadataId.Value.ToString("B");
                        ZeroBasedSheet.Cell(sheet, line, cell++).Value = entity.LogicalName;
                        ZeroBasedSheet.Cell(sheet, line, cell++).Value = attribute.LogicalName;
                        ZeroBasedSheet.Cell(sheet, line, cell++).Value = bAmd.OptionSet.FalseOption.Value;
                        ZeroBasedSheet.Cell(sheet, line, cell++).Value = "Description";

                        foreach (var lcid in languages)
                        {
                            var label = string.Empty;

                            if (bAmd.OptionSet.FalseOption.Description != null)
                            {
                                var optionLabel =
                                    bAmd.OptionSet.FalseOption.Description.LocalizedLabels.FirstOrDefault(l =>
                                        l.LanguageCode == lcid);
                                if (optionLabel != null)
                                {
                                    label = optionLabel.Label;
                                }
                            }

                            ZeroBasedSheet.Cell(sheet, line, cell++).Value = label;
                        }
                    }

                    if (settings.ExportNames)
                    {
                        line++;
                        cell = 0;

                        ZeroBasedSheet.Cell(sheet, line, cell++).Value = attribute.MetadataId.Value.ToString("B");
                        ZeroBasedSheet.Cell(sheet, line, cell++).Value = entity.LogicalName;
                        ZeroBasedSheet.Cell(sheet, line, cell++).Value = attribute.LogicalName;
                        ZeroBasedSheet.Cell(sheet, line, cell++).Value = bAmd.OptionSet.TrueOption.Value;
                        ZeroBasedSheet.Cell(sheet, line, cell++).Value = "Label";

                        foreach (var lcid in languages)
                        {
                            var label = string.Empty;

                            if (bAmd.OptionSet.TrueOption.Label != null)
                            {
                                var optionLabel =
                                    bAmd.OptionSet.TrueOption.Label.LocalizedLabels.FirstOrDefault(l =>
                                        l.LanguageCode == lcid);
                                if (optionLabel != null)
                                {
                                    label = optionLabel.Label;
                                }
                            }

                            ZeroBasedSheet.Cell(sheet, line, cell++).Value = label;
                        }
                    }

                    if (settings.ExportDescriptions)
                    {
                        line++;
                        cell = 0;

                        ZeroBasedSheet.Cell(sheet, line, cell++).Value = attribute.MetadataId.Value.ToString("B");
                        ZeroBasedSheet.Cell(sheet, line, cell++).Value = entity.LogicalName;
                        ZeroBasedSheet.Cell(sheet, line, cell++).Value = attribute.LogicalName;
                        ZeroBasedSheet.Cell(sheet, line, cell++).Value = bAmd.OptionSet.TrueOption.Value;
                        ZeroBasedSheet.Cell(sheet, line, cell++).Value = "Description";

                        foreach (var lcid in languages)
                        {
                            var label = string.Empty;

                            if (bAmd.OptionSet.TrueOption.Description != null)
                            {
                                var optionLabel =
                                    bAmd.OptionSet.TrueOption.Description.LocalizedLabels.FirstOrDefault(l =>
                                        l.LanguageCode == lcid);
                                if (optionLabel != null)
                                {
                                    label = optionLabel.Label;
                                }
                            }

                            ZeroBasedSheet.Cell(sheet, line, cell++).Value = label;
                        }
                    }
                }
            }

            // Applying style to cells
            for (int i = 0; i < (5 + languages.Count); i++)
            {
                StyleMutator.TitleCell(ZeroBasedSheet.Cell(sheet, 0, i).Style);
            }

            for (int i = 1; i <= line; i++)
            {
                for (int j = 0; j < 5; j++)
                {
                    StyleMutator.HighlightedCell(ZeroBasedSheet.Cell(sheet, i, j).Style);
                }
            }
        }

        public void Import(ExcelWorksheet sheet, List<EntityMetadata> emds, IOrganizationService service, BackgroundWorker worker)
        {
            OnLog(new LogEventArgs($"Reading {sheet.Name}"));

            var requests = new List<UpdateOptionValueRequest>();

            var rowsCount = sheet.Dimension.Rows;
            var cellsCount = sheet.Dimension.Columns;
            for (var rowI = 1; rowI < rowsCount; rowI++)
            {
                if (HasEmptyCells(sheet, rowI, 4)) continue;

                var value = int.Parse(ZeroBasedSheet.Cell(sheet, rowI, 3).Value.ToString());

                UpdateOptionValueRequest request = requests.FirstOrDefault(r => r.EntityLogicalName == ZeroBasedSheet.Cell(sheet, rowI, 1).Value.ToString()
                && r.AttributeLogicalName == ZeroBasedSheet.Cell(sheet, rowI, 2).Value.ToString()
                && r.Value == value);

                if (request == null)
                {
                    var emd = emds.FirstOrDefault(e => e.LogicalName == ZeroBasedSheet.Cell(sheet, rowI, 1).Value.ToString());
                    if (emd == null)
                    {
                        var mdRequest = new RetrieveEntityRequest
                        {
                            LogicalName = ZeroBasedSheet.Cell(sheet, rowI, 1).Value.ToString(),
                            EntityFilters = EntityFilters.Entity | EntityFilters.Attributes | EntityFilters.Relationships
                        };

                        var response = ((RetrieveEntityResponse)service.Execute(mdRequest));
                        emd = response.EntityMetadata;

                        emds.Add(emd);
                    }

                    var amd = (BooleanAttributeMetadata)emd.Attributes.FirstOrDefault(a => a.LogicalName == ZeroBasedSheet.Cell(sheet, rowI, 2).Value.ToString());
                    var dLabel = (value == 0 ? amd?.OptionSet.FalseOption.Label : amd?.OptionSet.TrueOption.Label) ?? new Label();
                    var descLabel = (value == 0 ? amd?.OptionSet.FalseOption.Description : amd?.OptionSet.TrueOption.Description) ?? new Label();

                    request = new UpdateOptionValueRequest
                    {
                        AttributeLogicalName = ZeroBasedSheet.Cell(sheet, rowI, 2).Value.ToString(),
                        EntityLogicalName = ZeroBasedSheet.Cell(sheet, rowI, 1).Value.ToString(),
                        Value = int.Parse(ZeroBasedSheet.Cell(sheet, rowI, 3).Value.ToString()),
                        Label = dLabel,
                        Description = descLabel,
                        MergeLabels = true
                    };

                    requests.Add(request);
                }

                int columnIndex = 5;

                if (ZeroBasedSheet.Cell(sheet, rowI, 4).Value.ToString() == "Label")
                {
                    while (columnIndex < cellsCount)
                    {
                        var lcid = int.Parse(ZeroBasedSheet.Cell(sheet, 0, columnIndex).Value.ToString());
                        var label = ZeroBasedSheet.Cell(sheet, rowI, columnIndex).Value.ToString();

                        if (label.Length > 0)
                        {
                            var translatedLabel = request.Label.LocalizedLabels.FirstOrDefault(x => x.LanguageCode == lcid);
                            if (translatedLabel == null)
                            {
                                translatedLabel = new LocalizedLabel(label, lcid);
                                request.Label.LocalizedLabels.Add(translatedLabel);
                            }
                            else
                            {
                                translatedLabel.Label = label;
                            }
                        }
                        columnIndex++;
                    }
                }
                else if (ZeroBasedSheet.Cell(sheet, rowI, 4).Value.ToString() == "Description")
                {
                    while (columnIndex < cellsCount)
                    {
                        var lcid = int.Parse(ZeroBasedSheet.Cell(sheet, 0, columnIndex).Value.ToString());
                        var label = ZeroBasedSheet.Cell(sheet, rowI, columnIndex).Value.ToString();

                        if (label.Length > 0)
                        {
                            var translatedLabel = request.Description.LocalizedLabels.FirstOrDefault(x => x.LanguageCode == lcid);
                            if (translatedLabel == null)
                            {
                                translatedLabel = new LocalizedLabel(label, lcid);
                                request.Description.LocalizedLabels.Add(translatedLabel);
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

            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "Attribute Id";
            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "Entity Logical Name";
            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "Attribute Logical Name";
            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "Value";
            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "Type";

            foreach (var lcid in languages)
            {
                ZeroBasedSheet.Cell(sheet, 0, cell++).Value = lcid.ToString(CultureInfo.InvariantCulture);
            }
        }
    }
}