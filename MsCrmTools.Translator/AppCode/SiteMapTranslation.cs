using McTools.Xrm.Connection;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml;

namespace MsCrmTools.Translator.AppCode
{
    public class SiteMapTranslation : BaseTranslation
    {
        private EntityCollection siteMaps;
        private List<Guid> siteMapsToTranslate = new List<Guid>();

        public SiteMapTranslation()
        {
            name = "Sitemaps";
        }

        /// <summary>
        ///
        /// </summary>
        /// <example>
        /// viewId;entityLogicalName;viewName;ViewType;Type;LCID1;LCID2;...;LCODX
        /// </example>
        /// <param name="languages"></param>
        /// <param name="file"></param>
        /// <param name="service"></param>
        /// <param name="settings"></param>
        public void Export(List<int> languages, ExcelWorkbook file, IOrganizationService service, ExportSettings settings, ConnectionDetail detail)
        {
            var line = 0;
            int cell;

            siteMaps = GetSiteMaps(settings, service, detail);

            var areaSheet = file.Worksheets.Add("SiteMap Areas");
            var groupSheet = file.Worksheets.Add("SiteMap Groups");
            var subAreaSheet = file.Worksheets.Add("SiteMap SubAreas");
            AddAreaHeader(areaSheet, languages);
            AddGroupHeader(groupSheet, languages);
            AddSubAreaHeader(subAreaSheet, languages);

            var crmSiteMapAreas = new List<CrmSiteMapArea>();
            var crmSiteMapGroups = new List<CrmSiteMapGroup>();
            var crmSiteMapSubAreas = new List<CrmSiteMapSubArea>();

            foreach (var siteMap in siteMaps.Entities)
            {
                var siteMapDoc = new XmlDocument();
                siteMapDoc.LoadXml(siteMap["sitemapxml"].ToString());

                #region Export Area

                var areaNodes = siteMapDoc.SelectNodes("SiteMap/Area");
                foreach (XmlNode areaNode in areaNodes)
                {
                    var area = new CrmSiteMapArea
                    {
                        Id = areaNode.Attributes["Id"].Value,
                        SiteMapId = siteMap.Id,
                        SiteMapName = siteMap.GetAttributeValue<string>("sitemapname")
                    };

                    if (settings.ExportNames)
                    {
                        foreach (XmlNode titleNode in areaNode.SelectNodes("Titles/Title"))
                        {
                            area.Titles.Add(int.Parse(titleNode.Attributes["LCID"].Value),
                                titleNode.Attributes["Title"].Value);
                        }
                    }

                    if (settings.ExportDescriptions)
                    {
                        foreach (XmlNode titleNode in areaNode.SelectNodes("Descriptions/Description"))
                        {
                            area.Descriptions.Add(int.Parse(titleNode.Attributes["LCID"].Value),
                                titleNode.Attributes["Description"].Value);
                        }
                    }

                    crmSiteMapAreas.Add(area);

                    #region Export Groups

                    var groupNodes = areaNode.SelectNodes("Group");
                    foreach (XmlNode groupNode in groupNodes)
                    {
                        var group = new CrmSiteMapGroup
                        {
                            Id = groupNode.Attributes["Id"].Value,
                            AreaId = areaNode.Attributes["Id"].Value,
                            SiteMapId = siteMap.Id,
                            SiteMapName = siteMap.GetAttributeValue<string>("sitemapname")
                        };

                        if (settings.ExportNames)
                        {
                            foreach (XmlNode titleNode in groupNode.SelectNodes("Titles/Title"))
                            {
                                group.Titles.Add(int.Parse(titleNode.Attributes["LCID"].Value),
                                    titleNode.Attributes["Title"].Value);
                            }
                        }

                        if (settings.ExportDescriptions)
                        {
                            foreach (XmlNode titleNode in groupNode.SelectNodes("Descriptions/Description"))
                            {
                                group.Descriptions.Add(int.Parse(titleNode.Attributes["LCID"].Value),
                                    titleNode.Attributes["Description"].Value);
                            }
                        }

                        crmSiteMapGroups.Add(group);

                        #region Export SubArea

                        var subAreaNodes = groupNode.SelectNodes("SubArea");
                        foreach (XmlNode subAreaNode in subAreaNodes)
                        {
                            var subArea = new CrmSiteMapSubArea()
                            {
                                Id = subAreaNode.Attributes["Id"].Value,
                                GroupId = groupNode.Attributes["Id"].Value,
                                AreaId = areaNode.Attributes["Id"].Value,
                                SiteMapId = siteMap.Id,
                                SiteMapName = siteMap.GetAttributeValue<string>("sitemapname")
                            };

                            if (settings.ExportNames)
                            {
                                foreach (XmlNode titleNode in subAreaNode.SelectNodes("Titles/Title"))
                                {
                                    subArea.Titles.Add(int.Parse(titleNode.Attributes["LCID"].Value),
                                        titleNode.Attributes["Title"].Value);
                                }
                            }

                            if (settings.ExportDescriptions)
                            {
                                foreach (XmlNode titleNode in subAreaNode.SelectNodes("Descriptions/Description"))
                                {
                                    subArea.Descriptions.Add(int.Parse(titleNode.Attributes["LCID"].Value),
                                        titleNode.Attributes["Description"].Value);
                                }
                            }

                            crmSiteMapSubAreas.Add(subArea);
                        }

                        #endregion Export SubArea
                    }

                    #endregion Export Groups
                }

                #endregion Export Area
            }

            #region Area sheet

            foreach (var crmArea in crmSiteMapAreas)
            {
                if (settings.ExportNames)
                {
                    line++;
                    cell = 0;
                    ZeroBasedSheet.Cell(areaSheet, line, cell++).Value = crmArea.SiteMapName;
                    ZeroBasedSheet.Cell(areaSheet, line, cell++).Value = crmArea.SiteMapId.ToString();
                    ZeroBasedSheet.Cell(areaSheet, line, cell++).Value = crmArea.Id;
                    ZeroBasedSheet.Cell(areaSheet, line, cell++).Value = "Title";

                    foreach (var lcid in languages)
                    {
                        ZeroBasedSheet.Cell(areaSheet, line, cell++).Value =
                            crmArea.Titles.FirstOrDefault(n => n.Key == lcid).Value;
                    }
                }

                if (settings.ExportDescriptions)
                {
                    line++;
                    cell = 0;
                    ZeroBasedSheet.Cell(areaSheet, line, cell++).Value = crmArea.SiteMapName;
                    ZeroBasedSheet.Cell(areaSheet, line, cell++).Value = crmArea.SiteMapId.ToString();
                    ZeroBasedSheet.Cell(areaSheet, line, cell++).Value = crmArea.Id;
                    ZeroBasedSheet.Cell(areaSheet, line, cell++).Value = "Description";

                    foreach (var lcid in languages)
                    {
                        ZeroBasedSheet.Cell(areaSheet, line, cell++).Value =
                            crmArea.Descriptions.FirstOrDefault(n => n.Key == lcid).Value;
                    }
                }
            }

            #endregion Area sheet

            #region Group sheet

            line = 0;
            foreach (var crmGroup in crmSiteMapGroups)
            {
                if (settings.ExportNames)
                {
                    line++;
                    cell = 0;
                    ZeroBasedSheet.Cell(groupSheet, line, cell++).Value = crmGroup.SiteMapName;
                    ZeroBasedSheet.Cell(groupSheet, line, cell++).Value = crmGroup.SiteMapId.ToString();
                    ZeroBasedSheet.Cell(groupSheet, line, cell++).Value = crmGroup.AreaId;
                    ZeroBasedSheet.Cell(groupSheet, line, cell++).Value = crmGroup.Id;
                    ZeroBasedSheet.Cell(groupSheet, line, cell++).Value = "Title";

                    foreach (var lcid in languages)
                    {
                        ZeroBasedSheet.Cell(groupSheet, line, cell++).Value =
                            crmGroup.Titles.FirstOrDefault(n => n.Key == lcid).Value;
                    }
                }

                if (settings.ExportDescriptions)
                {
                    line++;
                    cell = 0;
                    ZeroBasedSheet.Cell(groupSheet, line, cell++).Value = crmGroup.SiteMapName;
                    ZeroBasedSheet.Cell(groupSheet, line, cell++).Value = crmGroup.SiteMapId.ToString();
                    ZeroBasedSheet.Cell(groupSheet, line, cell++).Value = crmGroup.AreaId;
                    ZeroBasedSheet.Cell(groupSheet, line, cell++).Value = crmGroup.Id;
                    ZeroBasedSheet.Cell(groupSheet, line, cell++).Value = "Description";

                    foreach (var lcid in languages)
                    {
                        ZeroBasedSheet.Cell(groupSheet, line, cell++).Value =
                            crmGroup.Descriptions.FirstOrDefault(n => n.Key == lcid).Value;
                    }
                }
            }

            #endregion Group sheet

            #region SubArea sheet

            line = 0;
            foreach (var crmSubArea in crmSiteMapSubAreas)
            {
                if (settings.ExportNames)
                {
                    line++;
                    cell = 0;
                    ZeroBasedSheet.Cell(subAreaSheet, line, cell++).Value = crmSubArea.SiteMapName;
                    ZeroBasedSheet.Cell(subAreaSheet, line, cell++).Value = crmSubArea.SiteMapId.ToString();
                    ZeroBasedSheet.Cell(subAreaSheet, line, cell++).Value = crmSubArea.AreaId;
                    ZeroBasedSheet.Cell(subAreaSheet, line, cell++).Value = crmSubArea.GroupId;
                    ZeroBasedSheet.Cell(subAreaSheet, line, cell++).Value = crmSubArea.Id;
                    ZeroBasedSheet.Cell(subAreaSheet, line, cell++).Value = "Title";

                    foreach (var lcid in languages)
                    {
                        ZeroBasedSheet.Cell(subAreaSheet, line, cell++).Value =
                            crmSubArea.Titles.FirstOrDefault(n => n.Key == lcid).Value;
                    }
                }

                if (settings.ExportDescriptions)
                {
                    line++;
                    cell = 0;
                    ZeroBasedSheet.Cell(subAreaSheet, line, cell++).Value = crmSubArea.SiteMapName;
                    ZeroBasedSheet.Cell(subAreaSheet, line, cell++).Value = crmSubArea.SiteMapId.ToString();
                    ZeroBasedSheet.Cell(subAreaSheet, line, cell++).Value = crmSubArea.AreaId;
                    ZeroBasedSheet.Cell(subAreaSheet, line, cell++).Value = crmSubArea.GroupId;
                    ZeroBasedSheet.Cell(subAreaSheet, line, cell++).Value = crmSubArea.Id;
                    ZeroBasedSheet.Cell(subAreaSheet, line, cell++).Value = "Description";

                    foreach (var lcid in languages)
                    {
                        ZeroBasedSheet.Cell(subAreaSheet, line, cell++).Value =
                            crmSubArea.Descriptions.FirstOrDefault(n => n.Key == lcid).Value;
                    }
                }
            }

            #endregion SubArea sheet

            // Applying style to cells
            for (int i = 0; i < 4 + languages.Count; i++)
            {
                StyleMutator.TitleCell(ZeroBasedSheet.Cell(areaSheet, 0, i).Style);
            }
            for (int i = 1; i <= crmSiteMapAreas.Select(c => c.Titles).Count() + crmSiteMapAreas.Select(c => c.Descriptions).Count(); i++)
            {
                for (int j = 0; j < 4; j++)
                {
                    StyleMutator.HighlightedCell(ZeroBasedSheet.Cell(areaSheet, i, j).Style);
                }
            }

            for (int i = 0; i < 5 + languages.Count; i++)
            {
                StyleMutator.TitleCell(ZeroBasedSheet.Cell(groupSheet, 0, i).Style);
            }

            for (int i = 1; i <= crmSiteMapGroups.Select(c => c.Titles).Count() + crmSiteMapGroups.Select(c => c.Descriptions).Count(); i++)
            {
                for (int j = 0; j < 5; j++)
                {
                    StyleMutator.HighlightedCell(ZeroBasedSheet.Cell(groupSheet, i, j).Style);
                }
            }

            for (int i = 0; i < 6 + languages.Count; i++)
            {
                StyleMutator.TitleCell(ZeroBasedSheet.Cell(subAreaSheet, 0, i).Style);
            }

            for (int i = 1; i <= crmSiteMapSubAreas.Select(c => c.Titles).Count() + crmSiteMapSubAreas.Select(c => c.Descriptions).Count(); i++)
            {
                for (int j = 0; j < 6; j++)
                {
                    StyleMutator.HighlightedCell(ZeroBasedSheet.Cell(subAreaSheet, i, j).Style);
                }
            }
        }

        public EntityCollection GetSiteMaps(ExportSettings settings, IOrganizationService service, ConnectionDetail detail)
        {
            if (siteMaps != null) return siteMaps;

            var ec = new EntityCollection();

            if (detail.OrganizationMajorVersion >= 8 && detail.OrganizationMinorVersion >= 2)
            {
                List<Guid> ids = new List<Guid>();

                if ((settings?.SolutionId ?? Guid.Empty) != Guid.Empty)
                {
                    var components = service.RetrieveMultiple(new QueryExpression("solutioncomponent")
                    {
                        ColumnSet = new ColumnSet("objectid"),
                        Criteria = new FilterExpression
                        {
                            Conditions =
                            {
                                new ConditionExpression("componenttype", ConditionOperator.Equal, 62),
                                new ConditionExpression("solutionid", ConditionOperator.Equal, settings?.SolutionId ?? Guid.Empty)
                            }
                        }
                    });

                    ids.AddRange(components.Entities.Select(e => e.GetAttributeValue<Guid>("objectid")));
                }
                else
                {
                    var sitemapsIds = service.RetrieveMultiple(new QueryExpression("appmodulecomponent")
                    {
                        ColumnSet = new ColumnSet(true),
                        Criteria = new FilterExpression
                        {
                            Conditions =
                            {
                                new ConditionExpression("componenttype", ConditionOperator.Equal, 62)
                            }
                        }
                    });

                    foreach (var siteMapId in sitemapsIds.Entities)
                    {
                        if (ec.Entities.Any(ent => ent.Id == siteMapId.GetAttributeValue<Guid>("objectid") ||
                                                   siteMapId.GetAttributeValue<EntityReference>("appmoduleidunique").Name == null))
                        {
                            continue;
                        }

                        ids.Add(siteMapId.GetAttributeValue<Guid>("objectid"));
                    }
                }

                foreach (var id in ids)
                {
                    try
                    {
                        var tmpSiteMap = service.Retrieve("sitemap", id, new ColumnSet(true));

                        if (tmpSiteMap.GetAttributeValue<OptionSetValue>("componentstate")?.Value != 0)
                            continue;

                        ec.Entities.Add(tmpSiteMap);
                    }
                    catch (Exception error)
                    { }
                }
            }

            if (settings == null || settings.SolutionId == Guid.Empty)
            {
                // Adding default sitemap
                var qe = new QueryExpression("sitemap");
                qe.ColumnSet = new ColumnSet(true);
                EntityCollection ecDefault = service.RetrieveMultiple(qe);
                ecDefault.Entities.First()["sitemapname"] = "Default";
                ec.Entities.Add(ecDefault.Entities.First());
            }
            siteMaps = ec;

            return siteMaps;
        }

        public void Import(IOrganizationService service)
        {
            OnLog(new LogEventArgs("Importing SiteMap translations"));

            var arg = new TranslationProgressEventArgs { SheetName = "SiteMaps" };

            foreach (var siteMap in siteMaps.Entities)
            {
                if (siteMapsToTranslate.Contains(siteMap.Id))
                {
                    AddRequest(new UpdateRequest { Target = siteMap });
                    ExecuteMultiple(service, arg, siteMaps.Entities.Count);
                }
            }

            ExecuteMultiple(service, arg, siteMaps.Entities.Count, true);
        }

        public void PrepareAreas(ExcelWorksheet sheet, IOrganizationService service, ConnectionDetail detail)
        {
            OnLog(new LogEventArgs($"Reading {sheet.Name}"));

            GetSiteMaps(null, service, detail);

            foreach (var siteMap in siteMaps.Entities)
            {
                var siteMapDoc = new XmlDocument();
                siteMapDoc.LoadXml(siteMap["sitemapxml"].ToString());

                var rowsCount = sheet.Dimension.Rows;
                var cellsCount = sheet.Dimension.Columns;
                for (var rowI = 1; rowI < rowsCount; rowI++)
                {
                    if (HasEmptyCells(sheet, rowI, 3)) continue;

                    if (ZeroBasedSheet.Cell(sheet, rowI, 0).Value == null) break;
                    if (ZeroBasedSheet.Cell(sheet, rowI, 1).Value.ToString() != siteMap.Id.ToString()) continue;

                    if (!siteMapsToTranslate.Contains(siteMap.Id)) siteMapsToTranslate.Add(siteMap.Id);

                    var areaId = ZeroBasedSheet.Cell(sheet, rowI, 2).Value.ToString();
                    var areaNode = siteMapDoc.SelectSingleNode("SiteMap/Area[@Id='" + areaId + "']");
                    if (areaNode == null)
                    {
                        throw new Exception("Unable to find area with id " + areaId);
                    }

                    var columnIndex = 4;
                    while (columnIndex < cellsCount)
                    {
                        if (ZeroBasedSheet.Cell(sheet, rowI, columnIndex).Value != null)
                        {
                            var lcid = ZeroBasedSheet.Cell(sheet, 0, columnIndex).Value.ToString();
                            var label = ZeroBasedSheet.Cell(sheet, rowI, columnIndex).Value.ToString();

                            if (ZeroBasedSheet.Cell(sheet, rowI, 3).Value.ToString() == "Title")
                            {
                                UpdateXmlNode(areaNode, "Titles", "Title", lcid, label);
                            }
                            else
                            {
                                UpdateXmlNode(areaNode, "Descriptions", "Description", lcid, label);
                            }
                        }

                        columnIndex++;
                    }
                }

                siteMap["sitemapxml"] = siteMapDoc.OuterXml;
            }
        }

        public void PrepareGroups(ExcelWorksheet sheet, IOrganizationService service, ConnectionDetail detail)
        {
            OnLog(new LogEventArgs($"Reading {sheet.Name}"));

            GetSiteMaps(null, service, detail);

            foreach (var siteMap in siteMaps.Entities)
            {
                var siteMapDoc = new XmlDocument();
                siteMapDoc.LoadXml(siteMap["sitemapxml"].ToString());

                var rowsCount = sheet.Dimension.Rows;
                var cellsCount = sheet.Dimension.Columns;
                for (var rowI = 1; rowI < rowsCount; rowI++)
                {
                    if (HasEmptyCells(sheet, rowI, 4)) continue;

                    if (ZeroBasedSheet.Cell(sheet, rowI, 0).Value == null) break;
                    if (ZeroBasedSheet.Cell(sheet, rowI, 1).Value.ToString() != siteMap.Id.ToString()) continue;

                    if (!siteMapsToTranslate.Contains(siteMap.Id)) siteMapsToTranslate.Add(siteMap.Id);

                    var areaId = ZeroBasedSheet.Cell(sheet, rowI, 2).Value.ToString();
                    var groupId = ZeroBasedSheet.Cell(sheet, rowI, 3).Value.ToString();
                    var groupNode =
                        siteMapDoc.SelectSingleNode("SiteMap/Area[@Id='" + areaId + "']/Group[@Id='" + groupId + "']");
                    if (groupNode == null)
                    {
                        throw new Exception("Unable to find group with id " + groupId + " in area " + areaId);
                    }

                    var columnIndex = 5;
                    while (columnIndex < cellsCount)
                    {
                        if (ZeroBasedSheet.Cell(sheet, rowI, columnIndex).Value != null)
                        {
                            var lcid = ZeroBasedSheet.Cell(sheet, 0, columnIndex).Value.ToString();
                            var label = ZeroBasedSheet.Cell(sheet, rowI, columnIndex).Value.ToString();

                            if (ZeroBasedSheet.Cell(sheet, rowI, 4).Value.ToString() == "Title")
                            {
                                UpdateXmlNode(groupNode, "Titles", "Title", lcid, label);
                            }
                            else
                            {
                                UpdateXmlNode(groupNode, "Descriptions", "Description", lcid, label);
                            }
                        }

                        columnIndex++;
                    }
                }

                siteMap["sitemapxml"] = siteMapDoc.OuterXml;
            }
        }

        public void PrepareSubAreas(ExcelWorksheet sheet, IOrganizationService service, ConnectionDetail detail)
        {
            OnLog(new LogEventArgs($"Reading {sheet.Name}"));

            GetSiteMaps(null, service, detail);

            foreach (var siteMap in siteMaps.Entities)
            {
                var siteMapDoc = new XmlDocument();
                siteMapDoc.LoadXml(siteMap["sitemapxml"].ToString());

                var rowsCount = sheet.Dimension.Rows;
                var cellsCount = sheet.Dimension.Columns;
                for (var rowI = 1; rowI < rowsCount; rowI++)
                {
                    if (HasEmptyCells(sheet, rowI, 5)) continue;

                    if (ZeroBasedSheet.Cell(sheet, rowI, 0).Value == null) break;
                    if (ZeroBasedSheet.Cell(sheet, rowI, 1).Value.ToString() != siteMap.Id.ToString()) continue;

                    if (!siteMapsToTranslate.Contains(siteMap.Id)) siteMapsToTranslate.Add(siteMap.Id);

                    var areaId = ZeroBasedSheet.Cell(sheet, rowI, 2).Value.ToString();
                    var groupId = ZeroBasedSheet.Cell(sheet, rowI, 3).Value.ToString();
                    var subAreaId = ZeroBasedSheet.Cell(sheet, rowI, 4).Value.ToString();
                    var subAreaNode = siteMapDoc.SelectSingleNode(
                        "SiteMap/Area[@Id='" + areaId + "']/Group[@Id='" + groupId + "']/SubArea[@Id='" + subAreaId +
                        "']");
                    if (subAreaNode == null)
                    {
                        throw new Exception("Unable to find group with id " + groupId + " in area " + areaId);
                    }

                    var columnIndex = 6;
                    while (columnIndex < cellsCount)
                    {
                        if (ZeroBasedSheet.Cell(sheet, rowI, columnIndex).Value != null)
                        {
                            var lcid = ZeroBasedSheet.Cell(sheet, 0, columnIndex).Value.ToString();
                            var label = ZeroBasedSheet.Cell(sheet, rowI, columnIndex).Value.ToString();

                            if (ZeroBasedSheet.Cell(sheet, rowI, 5).Value.ToString() == "Title")
                            {
                                UpdateXmlNode(subAreaNode, "Titles", "Title", lcid, label);
                            }
                            else
                            {
                                UpdateXmlNode(subAreaNode, "Descriptions", "Description", lcid, label);
                            }
                        }

                        columnIndex++;
                    }
                }

                siteMap["sitemapxml"] = siteMapDoc.OuterXml;
            }
        }

        private void AddAreaHeader(ExcelWorksheet sheet, IEnumerable<int> languages)
        {
            var cell = 0;

            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "SiteMap Name";
            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "SiteMap Id";
            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "Area Id";
            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "Type";

            foreach (var lcid in languages)
            {
                ZeroBasedSheet.Cell(sheet, 0, cell++).Value = lcid.ToString(CultureInfo.InvariantCulture);
            }
        }

        private void AddGroupHeader(ExcelWorksheet sheet, IEnumerable<int> languages)
        {
            var cell = 0;

            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "SiteMap Name";
            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "SiteMap Id";
            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "Area Id";
            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "Group Id";
            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "Type";

            foreach (var lcid in languages)
            {
                ZeroBasedSheet.Cell(sheet, 0, cell++).Value = lcid.ToString(CultureInfo.InvariantCulture);
            }
        }

        private void AddSubAreaHeader(ExcelWorksheet sheet, IEnumerable<int> languages)
        {
            var cell = 0;

            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "SiteMap Name";
            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "SiteMap Id";
            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "Area Id";
            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "Group Id";
            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "SubArea Id";
            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "Type";

            foreach (var lcid in languages)
            {
                ZeroBasedSheet.Cell(sheet, 0, cell++).Value = lcid.ToString(CultureInfo.InvariantCulture);
            }
        }

        private void UpdateXmlNode(XmlNode node, string collectionName, string itemName, string lcid, string description)
        {
            XmlNode refNode;
            if (collectionName == "Titles" && node.FirstChild != null)
            {
                //Title should allways be first elemnt
                refNode = node.FirstChild;
            }
            else
            {
                switch (node.Name)
                {
                    case "Area":
                        refNode = node.SelectSingleNode("Group");
                        break;

                    case "Group":
                        refNode = node.SelectSingleNode("SubArea");
                        break;

                    case "SubArea":
                        refNode = node.SelectSingleNode("Privilege");
                        break;

                    default:
                        throw new Exception("Unexpected node name");
                }
            }

            var labelsNode = node.SelectSingleNode(collectionName);
            if (labelsNode == null)
            {
                labelsNode = node.OwnerDocument.CreateElement(collectionName);
                if (refNode != null)
                {
                    node.InsertBefore(labelsNode, refNode);
                }
                else
                {
                    node.AppendChild(labelsNode);
                }
            }

            var labelNode = labelsNode.SelectSingleNode(string.Format(itemName + "[@LCID='{0}']", lcid));
            if (labelNode == null)
            {
                labelNode = node.OwnerDocument.CreateElement(itemName);
                labelsNode.AppendChild(labelNode);

                var languageAttr = node.OwnerDocument.CreateAttribute("LCID");
                languageAttr.Value = lcid;
                labelNode.Attributes.Append(languageAttr);
                var descriptionAttr = node.OwnerDocument.CreateAttribute(itemName);
                labelNode.Attributes.Append(descriptionAttr);
            }

            labelNode.Attributes[itemName].Value = description;
        }
    }
}