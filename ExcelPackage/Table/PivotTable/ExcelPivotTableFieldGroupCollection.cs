/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 * 
 * All rights reserved.
 * 
 * EPPlus is an Open Source project provided under the 
 * GNU General Public License (GPL) as published by the 
 * Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA
 * 
 * See http://epplus.codeplex.com/ for details
 * 
 * The GNU General Public License can be viewed at http://www.opensource.org/licenses/gpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 * 
 * The code for this project may be used and redistributed by any means PROVIDING it is 
 * not sold for profit without the author's written consent, and providing that this notice 
 * and the author's name and all copyright notices remain intact.
 * 
 * All code and executables are provided "as is" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 * 
 * Author							Change						Date
 *******************************************************************************
 * Jan Källman		Added		21-MAR-2011
 *******************************************************************************/
using System;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Globalization;

namespace OfficeOpenXml.Table.PivotTable
{
    public enum eDateGroupBy
    {
        Years,
        Quarters,
        Months,
        Days,
        Hours,
        Minutes,
        Seconds
    }
    public class ExcelPivotTableFieldGroupCollection : IEnumerable<ExcelPivotTableFieldGroup>
    {
        ExcelPivotTableField _field;
        public ExcelPivotTableFieldGroupCollection(ExcelPivotTableField field)
        {
            _field = field;
        }
        List<ExcelPivotTableFieldGroup> _list = new List<ExcelPivotTableFieldGroup>();
        public IEnumerator<ExcelPivotTableFieldGroup> GetEnumerator()
        {
            return _list.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return _list.GetEnumerator();
        }
        public ExcelPivotTableFieldGroup AddDateGroup(eDateGroupBy GroupBy, DateTime StartDate, DateTime EndDate)
        {
            foreach (var grp in _list)
            {
                if (grp.GroupBy == GroupBy)
                {
                    throw(new ArgumentException("Grouping already exist in collection"));
                }
            }
            ExcelPivotTableFieldGroup group;
            if (_list.Count == 0)
            {
                group = new ExcelPivotTableFieldGroup(_field.NameSpaceManager, _field._cacheFieldHelper.TopNode, GroupBy);
                _field._cacheFieldHelper.SetXmlNodeBool("d:sharedItems/@containsDate",true);
                _field._cacheFieldHelper.SetXmlNodeBool("d:sharedItems/@containsSemiMixedTypes", false);

                group.TopNode.InnerXml += string.Format("<fieldGroup base=\"{0}\"><rangePr groupBy=\"{1}\" /><groupItems /></fieldGroup>", _field.Index.ToString(), GroupBy.ToString().ToLower());
                _field._cacheFieldHelper.SetXmlNodeString("d:fieldGroup/d:rangePr/@startDate", StartDate.ToString("s", CultureInfo.InvariantCulture));
                _field._cacheFieldHelper.SetXmlNodeString("d:fieldGroup/d:rangePr/@endDate", EndDate.ToString("s", CultureInfo.InvariantCulture));                
                int items=AddGroupItems(group, GroupBy, StartDate, EndDate);
                AddFieldItems(items);
            }
            else
            {
                var cacheXml = _field._table.CacheDefinition.CacheDefinitionXml;
                var fields = cacheXml.DocumentElement.SelectSingleNode("d:cacheFields\\d:cacheFields");
                var node = cacheXml.CreateElement("cacheField",ExcelPackage.schemaMain);

                node.SetAttribute("Name", GroupBy.ToString());
                node.SetAttribute("numFmtId", "0");
                node.SetAttribute("databaseField", _field.Index.ToString());
                fields.AppendChild(node);
                group = new ExcelPivotTableFieldGroup(_field.NameSpaceManager, node, GroupBy);
                group.TopNode.InnerXml += string.Format("<fieldGroup base=\"{0}\"><rangePr groupBy=\"quarters\" /></fieldGroup>", _field.Index.ToString());
            }
            _list.Add(group);
            return group;
        }

        private void AddFieldItems(int items)
        {
            XmlElement prevNode = null;
            XmlElement itemsNode = _field.TopNode.SelectSingleNode("//d:items", _field.NameSpaceManager) as XmlElement;
            for (int x = 0; x < items; x++)
            {
                var itemNode = itemsNode.OwnerDocument.CreateElement("item", ExcelPackage.schemaMain);
                itemNode.SetAttribute("x", x.ToString());
                if (prevNode == null)
                {
                    itemsNode.PrependChild(itemNode);
                }
                else
                {
                    itemsNode.InsertAfter(itemNode, prevNode);
                }
                prevNode = itemNode;
            }
            itemsNode.SetAttribute("count", (items + 1).ToString());
        }

        private int AddGroupItems(ExcelPivotTableFieldGroup group, eDateGroupBy GroupBy, DateTime StartDate, DateTime EndDate)
        {            
            XmlElement groupItems = group.TopNode.SelectSingleNode("//d:groupItems", group.NameSpaceManager) as XmlElement;
            int items = 2;
            //First date
            AddGroupItem(groupItems, "<" + StartDate.ToString("s", CultureInfo.InvariantCulture).Substring(1,10));
            
            switch (GroupBy)
            {
                case eDateGroupBy.Seconds:
                case eDateGroupBy.Minutes:
                    AddTimeSerie(60, groupItems);
                    items += 60;
                    break;
                case eDateGroupBy.Hours:
                    AddTimeSerie(24, groupItems);
                    items += 60;
                    break;
                case eDateGroupBy.Days:
                    DateTime dt = new DateTime(2008, 1, 1); //pick a year with 366 days
                    while (dt.Year == 2008) 
                    {
                        AddGroupItem(groupItems, dt.ToString("dd-MMM"));
                        dt = dt.AddDays(1);
                    }
                    items += 366;
                    break;
                case eDateGroupBy.Months:
                    AddGroupItem(groupItems, "jan");
                    AddGroupItem(groupItems, "feb");
                    AddGroupItem(groupItems, "mar");
                    AddGroupItem(groupItems, "apr");
                    AddGroupItem(groupItems, "may");
                    AddGroupItem(groupItems, "jun");
                    AddGroupItem(groupItems, "jul");
                    AddGroupItem(groupItems, "aug");
                    AddGroupItem(groupItems, "sep");
                    AddGroupItem(groupItems, "oct");
                    AddGroupItem(groupItems, "nov");
                    AddGroupItem(groupItems, "dec");
                    items += 12;
                    break;
                case eDateGroupBy.Quarters:
                    AddGroupItem(groupItems, "Qtr1");
                    AddGroupItem(groupItems, "Qtr2");
                    AddGroupItem(groupItems, "Qtr3");
                    AddGroupItem(groupItems, "Qtr4");
                    items += 4;
                    break;
                case eDateGroupBy.Years:
                    for (int year = StartDate.Year; year <= EndDate.Year; year++)
                    {
                        AddGroupItem(groupItems, year.ToString());
                    }
                    break;
                default:
                    throw (new Exception("unsupported grouping"));
            }

            //Lastdate
            AddGroupItem(groupItems, ">" + EndDate.ToString("s", CultureInfo.InvariantCulture).Substring(1,10));
            return items;
        }

        private void AddTimeSerie(int count, XmlElement groupItems)
        {
            for (int i = 0; i < count; i++)
            {
                AddGroupItem(groupItems, string.Format(":{0:00}",count));
            }
        }

        private void AddGroupItem(XmlElement groupItems, string value)
        {
            var s = groupItems.OwnerDocument.CreateElement("s", ExcelPackage.schemaMain);
            s.SetAttribute("v", value);
            groupItems.AppendChild(s);
        }
    }
}
