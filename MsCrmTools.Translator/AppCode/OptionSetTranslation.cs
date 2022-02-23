using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using OfficeOpenXml;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using Label = Microsoft.Xrm.Sdk.Label;

namespace MsCrmTools.Translator.AppCode
{
    public class OptionSetTranslation : BaseTranslation
    {
        public OptionSetTranslation()
        {
            name = "OptionSet values";
        }

        /// <summary>
        ///
        /// </summary>
        /// <example>
        /// attributeId;entityLogicalName;attributeLogicalName;OptionSetValue;LCID1;LCID2;...;LCIDX
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
                        || attribute.AttributeType.Value != AttributeTypeCode.Picklist
                        && attribute.AttributeType.Value != AttributeTypeCode.State
                        && attribute.AttributeType.Value != AttributeTypeCode.Status
                        && !(attribute is MultiSelectPicklistAttributeMetadata)
                        || !attribute.MetadataId.HasValue)
                        continue;

                    OptionSetMetadata omd = null;

                    switch (attribute.AttributeType.Value)
                    {
                        case AttributeTypeCode.Picklist:
                            omd = ((PicklistAttributeMetadata)attribute).OptionSet;
                            break;

                        case AttributeTypeCode.State:
                            omd = ((StateAttributeMetadata)attribute).OptionSet;
                            break;

                        case AttributeTypeCode.Status:
                            omd = ((StatusAttributeMetadata)attribute).OptionSet;
                            break;

                        case AttributeTypeCode.Virtual:
                            omd = ((MultiSelectPicklistAttributeMetadata)attribute).OptionSet;
                            break;
                    }

                    if (omd.IsGlobal.Value)
                        continue;

                    foreach (var option in omd.Options.OrderBy(o => o.Value))
                    {
                        if (settings.ExportNames)
                        {
                            line++;
                            cell = 0;

                            ZeroBasedSheet.Cell(sheet, line, cell++).Value = attribute.MetadataId.Value.ToString("B");
                            ZeroBasedSheet.Cell(sheet, line, cell++).Value = entity.LogicalName;
                            ZeroBasedSheet.Cell(sheet, line, cell++).Value = attribute.LogicalName;
                            ZeroBasedSheet.Cell(sheet, line, cell++).Value = attribute.AttributeType.Value.ToString();
                            ZeroBasedSheet.Cell(sheet, line, cell++).Value = option.Value;
                            ZeroBasedSheet.Cell(sheet, line, cell++).Value = "Label";

                            foreach (var lcid in languages)
                            {
                                var label = string.Empty;

                                if (option.Label != null)
                                {
                                    var optionLabel =
                                        option.Label.LocalizedLabels.FirstOrDefault(l => l.LanguageCode == lcid);
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
                            ZeroBasedSheet.Cell(sheet, line, cell++).Value = attribute.AttributeType.Value.ToString();
                            ZeroBasedSheet.Cell(sheet, line, cell++).Value = option.Value;
                            ZeroBasedSheet.Cell(sheet, line, cell++).Value = "Description";

                            foreach (var lcid in languages)
                            {
                                var label = string.Empty;

                                if (option.Description != null)
                                {
                                    var optionLabel =
                                        option.Description.LocalizedLabels.FirstOrDefault(l => l.LanguageCode == lcid);
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
            }

            // Applying style to cells
            for (int i = 0; i < (6 + languages.Count); i++)
            {
                StyleMutator.TitleCell(ZeroBasedSheet.Cell(sheet, 0, i).Style);
            }

            for (int i = 1; i <= line; i++)
            {
                for (int j = 0; j < 6; j++)
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

            for (var rowI = 1; rowI < rowsCount; rowI++)
            {
                if (HasEmptyCells(sheet, rowI, 5)) continue;

                var value = int.Parse(ZeroBasedSheet.Cell(sheet, rowI, 4).Value.ToString());

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

                var amd = emd.Attributes.FirstOrDefault(a => a.LogicalName == ZeroBasedSheet.Cell(sheet, rowI, 2).Value.ToString());
                OptionMetadata option = null;
                if (amd is PicklistAttributeMetadata pamd)
                {
                    option = pamd.OptionSet.Options.FirstOrDefault(o => o.Value == value);
                }
                else if (amd is StateAttributeMetadata samd)
                {
                    option = samd.OptionSet.Options.FirstOrDefault(o => o.Value == value);
                }
                else if (amd is StatusAttributeMetadata ssamd)
                {
                    option = ssamd.OptionSet.Options.FirstOrDefault(o => o.Value == value);
                }

                if (option == null)
                {
                    OnLog(new LogEventArgs($"Unable to determine type of the AttributeMetadata for attribute {amd.LogicalName}") { Type = LogType.Error });
                    continue;
                }

                UpdateOptionValueRequest request =
                    requests
                    .FirstOrDefault(
                        r => r.OptionSetName == ZeroBasedSheet.Cell(sheet, rowI, 1).Value.ToString() &&
                        r.Value == value);

                if (request == null)
                {
                    request = new UpdateOptionValueRequest
                    {
                        AttributeLogicalName = ZeroBasedSheet.Cell(sheet, rowI, 2).Value.ToString(),
                        EntityLogicalName = ZeroBasedSheet.Cell(sheet, rowI, 1).Value.ToString(),
                        Value = value,
                        Label = option.Label ?? new Label(),
                        Description = option.Description ?? new Label(),
                        MergeLabels = true
                    };

                    requests.Add(request);
                }

                int columnIndex = 6;

                if (ZeroBasedSheet.Cell(sheet, rowI, 5).Value.ToString() == "Label")
                {
                    // WTF: QUESTIONABLE DELETION: row.Cells.Count() > columnIndex &&
                    while (ZeroBasedSheet.Cell(sheet, rowI, columnIndex) != null && ZeroBasedSheet.Cell(sheet, rowI, columnIndex).Value != null)
                    {
                        var lcid = int.Parse(ZeroBasedSheet.Cell(sheet, 0, columnIndex).Value.ToString());
                        var label = ZeroBasedSheet.Cell(sheet, rowI, columnIndex).Value.ToString();

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
                        columnIndex++;
                    }
                }
                else if (ZeroBasedSheet.Cell(sheet, rowI, 5).Value.ToString() == "Description")
                {
                    // WTF: QUESTIONABLE DELETION: row.Cells.Count() > columnIndex &&
                    while (ZeroBasedSheet.Cell(sheet, rowI, columnIndex) != null && ZeroBasedSheet.Cell(sheet, rowI, columnIndex).Value != null)
                    {
                        var lcid = int.Parse(ZeroBasedSheet.Cell(sheet, 0, columnIndex).Value.ToString());
                        var label = ZeroBasedSheet.Cell(sheet, rowI, columnIndex).Value.ToString();

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
            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "Attribute Type";
            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "Value";
            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "Type";

            foreach (var lcid in languages)
            {
                ZeroBasedSheet.Cell(sheet, 0, cell++).Value = lcid.ToString(CultureInfo.InvariantCulture);
            }
        }
    }
}