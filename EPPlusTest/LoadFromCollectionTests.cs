﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace EPPlusTest
{
    [TestClass]
    public class LoadFromCollectionTests
    {
        internal abstract class BaseClass
        {
            public string Id { get; set; }
            public string Name { get; set; }
        }

        internal class Implementation : BaseClass
        {
            public int Number { get; set; }
        }

        internal class Aclass
        {
            public string Id { get; set; }
            public string Name { get; set; }
            public int Number { get; set; }
        }

        [TestMethod]
        public void ShouldUseAclassProperties()
        {
            var items = new List<Aclass>()
            {
                new Aclass(){ Id = "123", Name = "Item 1", Number = 3}
            };
            using (var pck = new ExcelPackage(new MemoryStream()))
            {
                var sheet = pck.Workbook.Worksheets.Add("sheet");
                sheet.Cells["C1"].LoadFromCollection(items, true, TableStyles.Dark1);

                Assert.AreEqual("Id", sheet.Cells["C1"].Value);
            }
        }

        [TestMethod]
        public void ShouldUseBaseClassProperties()
        {
            var items = new List<BaseClass>()
            {
                new Implementation(){ Id = "123", Name = "Item 1", Number = 3}
            };
            using (var pck = new ExcelPackage(new MemoryStream()))
            {
                var sheet = pck.Workbook.Worksheets.Add("sheet");
                sheet.Cells["C1"].LoadFromCollection(items, true, TableStyles.Dark1);

                Assert.AreEqual("Id", sheet.Cells["C1"].Value);
            }
        }

        [TestMethod]
        public void ShouldUseAnonymousProperties()
        {
            var objs = new List<BaseClass>()
            {
                new Implementation(){ Id = "123", Name = "Item 1", Number = 3}
            };
            var items = objs.Select(x => new {Id = x.Id, Name = x.Name}).ToList();
            using (var pck = new ExcelPackage(new MemoryStream()))
            {
                var sheet = pck.Workbook.Worksheets.Add("sheet");
                sheet.Cells["C1"].LoadFromCollection(items, true, TableStyles.Dark1);

                Assert.AreEqual("Id", sheet.Cells["C1"].Value);
            }
        }
        [TestMethod]
        [ExpectedException(typeof(InvalidCastException))]
        public void ShouldThrowInvalidCastExceptionIf()
        {
            var objs = new List<BaseClass>()
            {
                new Implementation(){ Id = "123", Name = "Item 1", Number = 3}
            };
            var items = objs.Select(x => new { Id = x.Id, Name = x.Name }).ToList();
            using (var pck = new ExcelPackage(new MemoryStream()))
            {
                var sheet = pck.Workbook.Worksheets.Add("sheet");
                sheet.Cells["C1"].LoadFromCollection(items, true, TableStyles.Dark1, BindingFlags.Public | BindingFlags.Instance, typeof(string).GetMembers());

                Assert.AreEqual("Id", sheet.Cells["C1"].Value);
            }
        }
    }
}
