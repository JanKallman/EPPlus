/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See https://github.com/JanKallman/EPPlus for details.
 *
 * Copyright (C) 2011  Jan Källman
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.

 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "as is" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 * 
 * Author							Change						Date
 *******************************************************************************
 * Jan Källman		Added		01-01-2012
 * Jan Källman      Added compression support 27-03-2012
 * Jan Källman      Native support for compound documents 2017-04-10
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using comTypes = System.Runtime.InteropServices.ComTypes;
using System.IO;
using System.Security;

namespace OfficeOpenXml.Utils.CompundDocument
{
    internal class CompoundDocument
    {        
        internal class StoragePart
        {
            public StoragePart()
            {

            }
            internal Dictionary<string, StoragePart> SubStorage = new Dictionary<string, StoragePart>();
            internal Dictionary<string, byte[]> DataStreams = new Dictionary<string, byte[]>();
        }
        internal StoragePart Storage = null;
        internal CompoundDocument()
        {
            Storage = new StoragePart();
        }
        internal CompoundDocument(MemoryStream ms)
        {
            Read(ms);
        }
        internal CompoundDocument(FileInfo fi)
        {
            Read(fi);
        }

        internal static bool IsCompoundDocument(FileInfo fi)
        {
            return CompoundDocumentFile.IsCompoundDocument(fi);
        }
        internal static bool IsCompoundDocument(MemoryStream ms)
        {
            return CompoundDocumentFile.IsCompoundDocument(ms);
        }

        internal CompoundDocument(byte[] doc)
        {
            Read(doc);
        }
        internal void Read(FileInfo fi)
        {
            var b = File.ReadAllBytes(fi.FullName);
            Read(b);
        }
        internal void Read(byte[] doc) 
        {
            Read(new MemoryStream(doc));
        }
        internal void Read(MemoryStream ms)
        {
            using (var doc = new CompoundDocumentFile(ms))
            {
                Storage = new StoragePart();
                GetStorageAndStreams(Storage, doc.RootItem);
            }
        }

        private void GetStorageAndStreams(StoragePart storage, CompoundDocumentItem parent)
        {
            foreach(var item in parent.Children)
            {
                if(item.ObjectType==1)      //Substorage
                {
                    var part = new StoragePart();
                    storage.SubStorage.Add(item.Name, part);
                    GetStorageAndStreams(part, item);
                }
                else if(item.ObjectType==2) //Stream
                {
                    storage.DataStreams.Add(item.Name, item.Stream);
                }
            }
        }
        internal void Save(MemoryStream ms)
        {
            var doc = new CompoundDocumentFile();
            WriteStorageAndStreams(Storage, doc.RootItem);
            doc.Write(ms);
        }

        private void WriteStorageAndStreams(StoragePart storage, CompoundDocumentItem parent)
        {
            foreach(var item in storage.SubStorage)
            {
                var c = new CompoundDocumentItem() { Name = item.Key, ObjectType = 1, Stream = null, StreamSize = 0, Parent = parent };
                parent.Children.Add(c);
                WriteStorageAndStreams(item.Value, c);
            }
            foreach (var item in storage.DataStreams)
            {
                var c = new CompoundDocumentItem() { Name = item.Key, ObjectType = 2, Stream = item.Value, StreamSize = (item.Value == null ? 0 : item.Value.Length), Parent = parent };
                parent.Children.Add(c);
            }
        }
    }
}