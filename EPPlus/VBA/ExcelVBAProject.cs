﻿/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See http://www.codeplex.com/EPPlus for details.
 *
 * Copyright (C) 2011  Jan Källman
 *
 * All code and executables are provided "as is" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * If you want to understand this code have a look at the Office VBA File Format Structure Specification (MS-OVBA.PDF) or
 * http://msdn.microsoft.com/en-us/library/cc313094(v=office.12).aspx
 * 
 * * Code change notes:
 * 
 * Author							Change						Date
 *******************************************************************************
 * Jan Källman		Added		26-MAR-2012
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.IO;
using OfficeOpenXml.Utils;
using System.Security.Cryptography.Pkcs;
using System.Security.Cryptography.X509Certificates;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using OfficeOpenXml.Utils.CompundDocument;

namespace OfficeOpenXml.VBA
{
    /// <summary>
    /// Represents the VBA project part of the package
    /// </summary>
    public class ExcelVbaProject
    {
        const string schemaRelVba = "http://schemas.microsoft.com/office/2006/relationships/vbaProject";
        internal const string PartUri = @"/xl/vbaProject.bin";
        #region Classes & Enums
        /// <summary>
        /// Type of system where the VBA project was created.
        /// </summary>
        public enum eSyskind
        {
            Win16 = 0,
            Win32 = 1,
            Macintosh = 2,
            Win64 = 3
        }

        #endregion
        internal ExcelVbaProject(ExcelWorkbook wb)
        {
            _wb = wb;
            _pck = _wb._package.Package;
            References = new ExcelVbaReferenceCollection();
            Modules = new ExcelVbaModuleCollection(this);
            var rel = _wb.Part.GetRelationshipsByType(schemaRelVba).FirstOrDefault();
            if (rel != null)
            {
                Uri = UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);
                Part = _pck.GetPart(Uri);
                GetProject();                
            }
            else
            {
                Lcid = 0;
                Part = null;
            }
        }
        internal ExcelWorkbook _wb;
        internal Packaging.ZipPackage _pck;
        #region Dir Stream Properties
        /// <summary>
        /// System kind. Default Win32.
        /// </summary>
        public eSyskind SystemKind { get; set; }
        /// <summary>
        /// Name of the project
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// A description of the project
        /// </summary>
        public string Description { get; set; }
        /// <summary>
        /// A helpfile
        /// </summary>
        public string HelpFile1 { get; set; }
        /// <summary>
        /// Secondary helpfile
        /// </summary>
        public string HelpFile2 { get; set; }
        /// <summary>
        /// Context if refering the helpfile
        /// </summary>
        public int HelpContextID { get; set; }
        /// <summary>
        /// Conditional compilation constants
        /// </summary>
        public string Constants { get; set; }
        /// <summary>
        /// Codepage for encoding. Default is current regional setting.
        /// </summary>
        public int CodePage  { get; internal set; }
        internal int LibFlags { get; set; }
        internal int MajorVersion { get; set; }
        internal int MinorVersion { get; set; }
        internal int Lcid { get; set; }
        internal int LcidInvoke { get; set; }
        internal string ProjectID { get; set; }
        internal string ProjectStreamText { get; set; }
        /// <summary>
        /// Project references
        /// </summary>        
        public ExcelVbaReferenceCollection References { get; set; }
        /// <summary>
        /// Code Modules (Modules, classes, designer code)
        /// </summary>
        public ExcelVbaModuleCollection Modules { get; set; }
        ExcelVbaSignature _signature = null;
        /// <summary>
        /// The digital signature
        /// </summary>
        public ExcelVbaSignature Signature
        {
            get
            {
                if (_signature == null)
                {
                    _signature=new ExcelVbaSignature(Part);
                }
                return _signature;
            }
        }
        ExcelVbaProtection _protection=null;
        /// <summary>
        /// VBA protection 
        /// </summary>
        public ExcelVbaProtection Protection
        {
            get
            {
                if (_protection == null)
                {
                    _protection = new ExcelVbaProtection(this);
                }
                return _protection;
            }
        }
        #endregion
        #region Read Project
        private void GetProject()
        {

            var stream = Part.GetStream();
            byte[] vba;
            vba = new byte[stream.Length];
            stream.Read(vba, 0, (int)stream.Length);
            Document = new CompoundDocument(vba);

            ReadDirStream();
            ProjectStreamText = Encoding.GetEncoding(CodePage).GetString(Document.Storage.DataStreams["PROJECT"]);
            ReadModules();
            ReadProjectProperties();
        }
        private void ReadModules()
        {
            foreach (var modul in Modules)
            {
                var stream = Document.Storage.SubStorage["VBA"].DataStreams[modul.streamName];
                var byCode = VBACompression.DecompressPart(stream, (int)modul.ModuleOffset);
                string code = Encoding.GetEncoding(CodePage).GetString(byCode);
                int pos=0;
                while(pos+9<code.Length && code.Substring(pos,9)=="Attribute")
                {
                    int linePos=code.IndexOf("\r\n",pos);
                    string[] lineSplit;
                    if(linePos>0)
                    {
                        lineSplit = code.Substring(pos + 9, linePos - pos - 9).Split('=');
                    }
                    else
                    {
                        lineSplit=code.Substring(pos+9).Split(new char[]{'='},1);
                    }
                    if (lineSplit.Length > 1)
                    {
                        lineSplit[1] = lineSplit[1].Trim();
                        var attr = 
                            new ExcelVbaModuleAttribute()
                        {
                            Name = lineSplit[0].Trim(),
                            DataType = lineSplit[1].StartsWith("\"") ? eAttributeDataType.String : eAttributeDataType.NonString,
                            Value = lineSplit[1].StartsWith("\"") ? lineSplit[1].Substring(1, lineSplit[1].Length - 2) : lineSplit[1]
                        };
                        modul.Attributes._list.Add(attr);
                    }
                    pos = linePos + 2;
                }
                modul.Code=code.Substring(pos);
            }
        }

        private void ReadProjectProperties()
        {
            _protection = new ExcelVbaProtection(this);
            string prevPackage = "";
            var lines = Regex.Split(ProjectStreamText, "\r\n");
            foreach (string line in lines)
            {
                if (line.StartsWith("["))
                {

                }
                else
                {
                    var split = line.Split('=');
                    if (split.Length > 1 && split[1].Length > 1 && split[1].StartsWith("\"")) //Remove any double qouates
                    {
                        split[1] = split[1].Substring(1, split[1].Length - 2);
                    }
                    switch (split[0])
                    {
                        case "ID":
                            ProjectID = split[1];
                            break;
                        case "Document":
                            string mn = split[1].Substring(0, split[1].IndexOf("/&H"));
                            Modules[mn].Type = eModuleType.Document;
                            break;
                        case "Package":
                            prevPackage = split[1];
                            break;
                        case "BaseClass":
                            Modules[split[1]].Type = eModuleType.Designer;
                            Modules[split[1]].ClassID = prevPackage;
                            break;
                        case "Module":
                            Modules[split[1]].Type = eModuleType.Module;
                            break;
                        case "Class":
                            Modules[split[1]].Type = eModuleType.Class;
                            break;
                        case "HelpFile":
                        case "Name":
                        case "HelpContextID":
                        case "Description":
                        case "VersionCompatible32":
                            break;
                        //393222000"
                        case "CMG":
                            byte[] cmg = Decrypt(split[1]);
                            _protection.UserProtected = (cmg[0] & 1) != 0;
                            _protection.HostProtected = (cmg[0] & 2) != 0;
                            _protection.VbeProtected = (cmg[0] & 4) != 0;
                            break;
                        case "DPB":
                            byte[] dpb = Decrypt(split[1]);
                            if (dpb.Length >= 28)
                            {
                                byte reserved = dpb[0];
                                var flags = new byte[3];
                                Array.Copy(dpb, 1, flags, 0, 3);
                                var keyNoNulls = new byte[4];
                                _protection.PasswordKey = new byte[4];
                                Array.Copy(dpb, 4, keyNoNulls, 0, 4);
                                var hashNoNulls = new byte[20];
                                _protection.PasswordHash = new byte[20];
                                Array.Copy(dpb, 8, hashNoNulls, 0, 20);
                                //Handle 0x00 bitwise 2.4.4.3 
                                for (int i = 0; i < 24; i++)
                                {
                                    int bit = 128 >> (int)((i % 8));
                                    if (i < 4)
                                    {
                                        if ((int)(flags[0] & bit) == 0)
                                        {
                                            _protection.PasswordKey[i] = 0;
                                        }
                                        else
                                        {
                                            _protection.PasswordKey[i] = keyNoNulls[i];
                                        }
                                    }
                                    else
                                    {
                                        int flagIndex = (i - i % 8) / 8;
                                        if ((int)(flags[flagIndex] & bit) == 0)
                                        {
                                            _protection.PasswordHash[i - 4] = 0;
                                        }
                                        else
                                        {
                                            _protection.PasswordHash[i - 4] = hashNoNulls[i - 4];
                                        }
                                    }
                                }
                            }
                            break;
                        case "GC":
                            _protection.VisibilityState = Decrypt(split[1])[0] == 0xFF;

                            break;
                    }
                }
            }
        }

        /// <summary>
        /// 2.4.3.3 Decryption
        /// </summary>
        /// <param name="value">Byte hex string</param>
        /// <returns>The decrypted value</returns>
        private byte[] Decrypt(string value)
        {
            byte[] enc = GetByte(value);
            byte[] dec = new byte[(value.Length - 1)];
            byte seed, version, projKey, ignoredLength;
            seed = enc[0];
            dec[0] = (byte)(enc[1] ^ seed);
            dec[1] = (byte)(enc[2] ^ seed);
            for (int i = 2; i < enc.Length - 1; i++)
            {
                dec[i] = (byte)(enc[i + 1] ^ (enc[i - 1] + dec[i - 1]));
            }
            version = dec[0];
            projKey = dec[1];
            ignoredLength = (byte)((seed & 6) / 2);
            int datalength = BitConverter.ToInt32(dec, ignoredLength + 2);
            var data = new byte[datalength];
            Array.Copy(dec, 6 + ignoredLength, data, 0, datalength);
            return data;
        }
        /// <summary>
        /// 2.4.3.2 Encryption
        /// </summary>
        /// <param name="value"></param>
        /// <returns>Byte hex string</returns>
        private string Encrypt(byte[] value)
        {
            byte[] seed = new byte[1];
            var rn = RandomNumberGenerator.Create();
            rn.GetBytes(seed);
            BinaryWriter br = new BinaryWriter(new MemoryStream());
            byte[] enc = new byte[value.Length + 10];
            enc[0] = seed[0];
            enc[1] = (byte)(2 ^ seed[0]);

            byte projKey = 0;

            foreach (var c in ProjectID)
            {
                projKey += (byte)c;
            }
            enc[2] = (byte)(projKey ^ seed[0]);
            var ignoredLength = (seed[0] & 6) / 2;
            for (int i = 0; i < ignoredLength; i++)
            {
                br.Write(seed[0]);
            }
            br.Write(value.Length);
            br.Write(value);

            int pos = 3;
            byte pb = projKey;
            foreach (var b in ((MemoryStream)br.BaseStream).ToArray())
            {
                enc[pos] = (byte)(b ^ (enc[pos - 2] + pb));
                pos++;
                pb = b;
            }

            return GetString(enc, pos - 1);
        }
        private string GetString(byte[] value, int max)
        {
            string ret = "";
            for (int i = 0; i <= max; i++)
            {
                if (value[i] < 16)
                {
                    ret += "0" + value[i].ToString("x");
                }
                else
                {
                    ret += value[i].ToString("x");
                }
            }
            return ret.ToUpperInvariant();
        }
        private byte[] GetByte(string value)
        {
            byte[] ret = new byte[value.Length / 2];
            for (int i = 0; i < ret.Length; i++)
            {
                ret[i] = byte.Parse(value.Substring(i * 2, 2), System.Globalization.NumberStyles.AllowHexSpecifier);
            }
            return ret;
        }
        private void ReadDirStream()
        {
            byte[] dir = VBACompression.DecompressPart(Document.Storage.SubStorage["VBA"].DataStreams["dir"]);
            MemoryStream ms = new MemoryStream(dir);
            BinaryReader br = new BinaryReader(ms);
            ExcelVbaReference currentRef = null;
            string referenceName = "";
            ExcelVBAModule currentModule = null;
            bool terminate = false;
            while (br.BaseStream.Position < br.BaseStream.Length && terminate == false)
            {
                ushort id = br.ReadUInt16();
                uint size = br.ReadUInt32();
                switch (id)
                {
                    case 0x01:
                        SystemKind = (eSyskind)br.ReadUInt32();
                        break;
                    case 0x02:
                        Lcid = (int)br.ReadUInt32();
                        break;
                    case 0x03:
                        CodePage = (int)br.ReadUInt16();
                        break;
                    case 0x04:
                        Name = GetString(br, size);
                        break;
                    case 0x05:
                        Description = GetUnicodeString(br, size);
                        break;
                    case 0x06:
                        HelpFile1 = GetString(br, size);
                        break;
                    case 0x3D:
                        HelpFile2 = GetString(br, size);
                        break;
                    case 0x07:
                        HelpContextID = (int)br.ReadUInt32();
                        break;
                    case 0x08:
                        LibFlags = (int)br.ReadUInt32();
                        break;
                    case 0x09:
                        MajorVersion = (int)br.ReadUInt32();
                        MinorVersion = (int)br.ReadUInt16();
                        break;
                    case 0x0C:
                        Constants = GetUnicodeString(br, size);
                        break;
                    case 0x0D:
                        uint sizeLibID = br.ReadUInt32();
                        var regRef = new ExcelVbaReference();
                        regRef.Name = referenceName;
                        regRef.ReferenceRecordID = id;
                        regRef.Libid = GetString(br, sizeLibID);
                        uint reserved1 = br.ReadUInt32();
                        ushort reserved2 = br.ReadUInt16();
                        References.Add(regRef);
                        break;
                    case 0x0E:
                        var projRef = new ExcelVbaReferenceProject();
                        projRef.ReferenceRecordID = id;
                        projRef.Name = referenceName;
                        sizeLibID = br.ReadUInt32();
                        projRef.Libid = GetString(br, sizeLibID);
                        sizeLibID = br.ReadUInt32();
                        projRef.LibIdRelative = GetString(br, sizeLibID);
                        projRef.MajorVersion = br.ReadUInt32();
                        projRef.MinorVersion = br.ReadUInt16();
                        References.Add(projRef);
                        break;
                    case 0x0F:
                        ushort modualCount = br.ReadUInt16();
                        break;
                    case 0x13:
                        ushort cookie = br.ReadUInt16();
                        break;
                    case 0x14:
                        LcidInvoke = (int)br.ReadUInt32();
                        break;
                    case 0x16:
                        referenceName = GetUnicodeString(br, size);
                        break;
                    case 0x19:
                        currentModule = new ExcelVBAModule();
                        currentModule.Name = GetUnicodeString(br, size);
                        Modules.Add(currentModule);
                        break;
                    case 0x1A:
                        currentModule.streamName = GetUnicodeString(br, size);
                        break;
                    case 0x1C:
                        currentModule.Description = GetUnicodeString(br, size);
                        break;
                    case 0x1E:
                        currentModule.HelpContext = (int)br.ReadUInt32();
                        break;
                    case 0x21:
                    case 0x22:
                        break;
                    case 0x2B:      //Modul Terminator
                        break;
                    case 0x2C:
                        currentModule.Cookie = br.ReadUInt16();
                        break;
                    case 0x31:
                        currentModule.ModuleOffset = br.ReadUInt32();
                        break;
                    case 0x10:
                        terminate = true;
                        break;
                    case 0x30:
                        var extRef = (ExcelVbaReferenceControl)currentRef;
                        var sizeExt = br.ReadUInt32();
                        extRef.LibIdExternal = GetString(br, sizeExt);

                        uint reserved4 = br.ReadUInt32();
                        ushort reserved5 = br.ReadUInt16();
                        extRef.OriginalTypeLib = new Guid(br.ReadBytes(16));
                        extRef.Cookie = br.ReadUInt32();
                        break;
                    case 0x33:
                        currentRef = new ExcelVbaReferenceControl();
                        currentRef.ReferenceRecordID = id;
                        currentRef.Name = referenceName;
                        currentRef.Libid = GetString(br, size);
                        References.Add(currentRef);
                        break;
                    case 0x2F:
                        var contrRef = (ExcelVbaReferenceControl)currentRef;
                        contrRef.ReferenceRecordID = id;

                        var sizeTwiddled = br.ReadUInt32();
                        contrRef.LibIdTwiddled = GetString(br, sizeTwiddled);
                        var r1 = br.ReadUInt32();
                        var r2 = br.ReadUInt16();

                        break;
                    case 0x25:
                        currentModule.ReadOnly = true;
                        break;
                    case 0x28:
                        currentModule.Private = true;
                        break;
                    default:
                        break;
                }
            }
        }
        #endregion

        #region Save Project
        internal void Save()
        {
            if (Validate())
            {
                CompoundDocument doc = new CompoundDocument();
                doc.Storage = new CompoundDocument.StoragePart();
                var store = new CompoundDocument.StoragePart();
                doc.Storage.SubStorage.Add("VBA", store);

                store.DataStreams.Add("_VBA_PROJECT", CreateVBAProjectStream());
                store.DataStreams.Add("dir", CreateDirStream());
                foreach (var module in Modules)
                {
                    store.DataStreams.Add(module.Name, VBACompression.CompressPart(Encoding.GetEncoding(CodePage).GetBytes(module.Attributes.GetAttributeText() + module.Code)));
                }

                //Copy streams from the template, if used.
                if (Document != null)
                {
                    foreach (var ss in Document.Storage.SubStorage)
                    {
                        if (ss.Key != "VBA")
                        {
                            doc.Storage.SubStorage.Add(ss.Key, ss.Value);
                        }
                    }
                    foreach (var s in Document.Storage.DataStreams)
                    {
                        if (s.Key != "dir" && s.Key != "PROJECT" && s.Key != "PROJECTwm")
                        {
                            doc.Storage.DataStreams.Add(s.Key, s.Value);
                        }
                    }
                }

                doc.Storage.DataStreams.Add("PROJECT", CreateProjectStream());
                doc.Storage.DataStreams.Add("PROJECTwm", CreateProjectwmStream());

                if (Part == null)
                {
                    Uri = new Uri(PartUri, UriKind.Relative);
                    Part = _pck.CreatePart(Uri, ExcelPackage.schemaVBA);
                    var rel = _wb.Part.CreateRelationship(Uri, Packaging.TargetMode.Internal, schemaRelVba);
                }
                var st = Part.GetStream(FileMode.Create);
                doc.Save(st);
                st.Flush();
                //Save the digital signture
                Signature.Save(this);
            }
        }

        private bool Validate()
        {
            Description = Description ?? "";
            HelpFile1 = HelpFile1 ?? "";
            HelpFile2 = HelpFile2 ?? "";
            Constants = Constants ?? "";
            return true;
        }

        /// <summary>
        /// MS-OVBA 2.3.4.1
        /// </summary>
        /// <returns></returns>
        private byte[] CreateVBAProjectStream()
        {
            BinaryWriter bw = new BinaryWriter(new MemoryStream());
            bw.Write((ushort)0x61CC); //Reserved1
            bw.Write((ushort)0xFFFF); //Version
            bw.Write((byte)0x0); //Reserved3
            bw.Write((ushort)0x0); //Reserved4
            return ((MemoryStream)bw.BaseStream).ToArray();
        }
        /// <summary>
        /// MS-OVBA 2.3.4.1
        /// </summary>
        /// <returns></returns>
        private byte[] CreateDirStream()
        {
            BinaryWriter bw = new BinaryWriter(new MemoryStream());

            /****** PROJECTINFORMATION Record ******/
            bw.Write((ushort)1);        //ID
            bw.Write((uint)4);          //Size
            bw.Write((uint)SystemKind); //SysKind

            bw.Write((ushort)2);        //ID
            bw.Write((uint)4);          //Size
            bw.Write((uint)Lcid);       //Lcid

            bw.Write((ushort)0x14);     //ID
            bw.Write((uint)4);          //Size
            bw.Write((uint)LcidInvoke); //Lcid Invoke

            bw.Write((ushort)3);        //ID
            bw.Write((uint)2);          //Size
            bw.Write((ushort)CodePage);   //Codepage

            //ProjectName
            bw.Write((ushort)4);                                            //ID
            bw.Write((uint)Name.Length);                             //Size
            bw.Write(Encoding.GetEncoding(CodePage).GetBytes(Name)); //Project Name

            //Description
            bw.Write((ushort)5);                                            //ID
            bw.Write((uint)Description.Length);                             //Size
            bw.Write(Encoding.GetEncoding(CodePage).GetBytes(Description)); //Project Name
            bw.Write((ushort)0x40);                                           //ID
            bw.Write((uint)Description.Length*2);                           //Size
            bw.Write(Encoding.Unicode.GetBytes(Description));               //Project Description

            //Helpfiles
            bw.Write((ushort)6);                                           //ID
            bw.Write((uint)HelpFile1.Length);                              //Size
            bw.Write(Encoding.GetEncoding(CodePage).GetBytes(HelpFile1));  //HelpFile1            
            bw.Write((ushort)0x3D);                                           //ID
            bw.Write((uint)HelpFile2.Length);                              //Size
            bw.Write(Encoding.GetEncoding(CodePage).GetBytes(HelpFile2));  //HelpFile2

            //Help context id
            bw.Write((ushort)7);            //ID
            bw.Write((uint)4);              //Size
            bw.Write((uint)HelpContextID);  //Help context id

            //Libflags
            bw.Write((ushort)8);            //ID
            bw.Write((uint)4);              //Size
            bw.Write((uint)0);  //Help context id

            //Vba Version
            bw.Write((ushort)9);            //ID
            bw.Write((uint)4);              //Reserved
            bw.Write((uint)MajorVersion);   //Reserved
            bw.Write((ushort)MinorVersion); //Help context id

            //Constants
            bw.Write((ushort)0x0C);           //ID
            bw.Write((uint)Constants.Length);              //Size
            bw.Write(Encoding.GetEncoding(CodePage).GetBytes(Constants));              //Help context id
            bw.Write((ushort)0x3C);                                           //ID
            bw.Write((uint)Constants.Length/2);                              //Size
            bw.Write(Encoding.Unicode.GetBytes(Constants));  //HelpFile2

            /****** PROJECTREFERENCES Record ******/
            foreach (var reference in References)
            {
                WriteNameReference(bw, reference);

                if (reference.ReferenceRecordID == 0x2F)
                {
                    WriteControlReference(bw, reference);
                }
                else if (reference.ReferenceRecordID == 0x33)
                {
                    WriteOrginalReference(bw, reference);
                }
                else if (reference.ReferenceRecordID == 0x0D)
                {
                    WriteRegisteredReference(bw, reference);
                }
                else if (reference.ReferenceRecordID == 0x0E)
                {
                    WriteProjectReference(bw, reference);
                }
            }

            bw.Write((ushort)0x0F);
            bw.Write((uint)0x02);
            bw.Write((ushort)Modules.Count);
            bw.Write((ushort)0x13);
            bw.Write((uint)0x02);
            bw.Write((ushort)0xFFFF);

            foreach (var module in Modules)
            {
                WriteModuleRecord(bw, module);
            }
            bw.Write((ushort)0x10);             //Terminator
            bw.Write((uint)0);              

            return VBACompression.CompressPart(((MemoryStream)bw.BaseStream).ToArray());
        }

        private void WriteModuleRecord(BinaryWriter bw, ExcelVBAModule module)
        {
            bw.Write((ushort)0x19);
            bw.Write((uint)module.Name.Length);
            bw.Write(Encoding.GetEncoding(CodePage).GetBytes(module.Name));     //Name

            bw.Write((ushort)0x47);
            bw.Write((uint)module.Name.Length*2);
            bw.Write(Encoding.Unicode.GetBytes(module.Name));                   //Name

            bw.Write((ushort)0x1A);
            bw.Write((uint)module.Name.Length);
            bw.Write(Encoding.GetEncoding(CodePage).GetBytes(module.Name));     //Stream Name  

            bw.Write((ushort)0x32);
            bw.Write((uint)module.Name.Length*2);
            bw.Write(Encoding.Unicode.GetBytes(module.Name));                   //Stream Name

            module.Description = module.Description ?? "";
            bw.Write((ushort)0x1C);
            bw.Write((uint)module.Description.Length);
            bw.Write(Encoding.GetEncoding(CodePage).GetBytes(module.Description));     //Description

            bw.Write((ushort)0x48);
            bw.Write((uint)module.Description.Length*2);
            bw.Write(Encoding.Unicode.GetBytes(module.Description));                   //Description

            bw.Write((ushort)0x31);
            bw.Write((uint)4);
            bw.Write((uint)0);                              //Module Stream Offset (No PerformanceCache)

            bw.Write((ushort)0x1E);
            bw.Write((uint)4);
            bw.Write((uint)module.HelpContext);            //Help context ID

            bw.Write((ushort)0x2C);
            bw.Write((uint)2);
            bw.Write((ushort)0xFFFF);            //Help context ID

            bw.Write((ushort)(module.Type == eModuleType.Module ? 0x21 : 0x22));
            bw.Write((uint)0);

            if (module.ReadOnly)
            {
                bw.Write((ushort)0x25);
                bw.Write((uint)0);              //Readonly
            }

            if (module.Private)
            {
                bw.Write((ushort)0x28);
                bw.Write((uint)0);              //Private
            }

            bw.Write((ushort)0x2B);             //Terminator
            bw.Write((uint)0);              
        }

        private void WriteNameReference(BinaryWriter bw, ExcelVbaReference reference)
        {
            //Name record
            bw.Write((ushort)0x16);                                             //ID
            bw.Write((uint)reference.Name.Length);                              //Size
            bw.Write(Encoding.GetEncoding(CodePage).GetBytes(reference.Name));  //HelpFile1
            bw.Write((ushort)0x3E);                                             //ID
            bw.Write((uint)reference.Name.Length * 2);                            //Size
            bw.Write(Encoding.Unicode.GetBytes(reference.Name));                //HelpFile2
        }
        private void WriteControlReference(BinaryWriter bw, ExcelVbaReference reference)
        {
            WriteOrginalReference(bw, reference);

            bw.Write((ushort)0x2F);
            var controlRef=(ExcelVbaReferenceControl)reference;
            bw.Write((uint)(4 + controlRef.LibIdTwiddled.Length + 4 + 2));    // Size of SizeOfLibidTwiddled, LibidTwiddled, Reserved1, and Reserved2.
            bw.Write((uint)controlRef.LibIdTwiddled.Length);                              //Size            
            bw.Write(Encoding.GetEncoding(CodePage).GetBytes(controlRef.LibIdTwiddled));  //LibID
            bw.Write((uint)0);      //Reserved1
            bw.Write((ushort)0);    //Reserved2
            WriteNameReference(bw, reference);  //Name record again
            bw.Write((ushort)0x30); //Reserved3
            bw.Write((uint)(4 + controlRef.LibIdExternal.Length + 4 + 2 + 16 + 4));    //Size of SizeOfLibidExtended, LibidExtended, Reserved4, Reserved5, OriginalTypeLib, and Cookie
            bw.Write((uint)controlRef.LibIdExternal.Length);                              //Size            
            bw.Write(Encoding.GetEncoding(CodePage).GetBytes(controlRef.LibIdExternal));  //LibID
            bw.Write((uint)0);      //Reserved4
            bw.Write((ushort)0);    //Reserved5
            bw.Write(controlRef.OriginalTypeLib.ToByteArray());
            bw.Write((uint)controlRef.Cookie);      //Cookie
        }

        private void WriteOrginalReference(BinaryWriter bw, ExcelVbaReference reference)
        {
            bw.Write((ushort)0x33);
            bw.Write((uint)reference.Libid.Length);
            bw.Write(Encoding.GetEncoding(CodePage).GetBytes(reference.Libid));  //LibID
        }
        private void WriteProjectReference(BinaryWriter bw, ExcelVbaReference reference)
        {
            bw.Write((ushort)0x0E);
            var projRef = (ExcelVbaReferenceProject)reference;
            bw.Write((uint)(4 + projRef.Libid.Length + 4 + projRef.LibIdRelative.Length+4+2));
            bw.Write((uint)projRef.Libid.Length);
            bw.Write(Encoding.GetEncoding(CodePage).GetBytes(projRef.Libid));  //LibAbsolute
            bw.Write((uint)projRef.LibIdRelative.Length);
            bw.Write(Encoding.GetEncoding(CodePage).GetBytes(projRef.LibIdRelative));  //LibIdRelative
            bw.Write(projRef.MajorVersion);
            bw.Write(projRef.MinorVersion);
        }

        private void WriteRegisteredReference(BinaryWriter bw, ExcelVbaReference reference)
        {
            bw.Write((ushort)0x0D);
            bw.Write((uint)(4+reference.Libid.Length+4+2));
            bw.Write((uint)reference.Libid.Length);
            bw.Write(Encoding.GetEncoding(CodePage).GetBytes(reference.Libid));  //LibID            
            bw.Write((uint)0);      //Reserved1
            bw.Write((ushort)0);    //Reserved2
        }

        private byte[] CreateProjectwmStream()
        {
            BinaryWriter bw = new BinaryWriter(new MemoryStream());

            foreach (var module in Modules)
            {
                bw.Write(Encoding.GetEncoding(CodePage).GetBytes(module.Name));     //Name
                bw.Write((byte)0); //Null
                bw.Write(Encoding.Unicode.GetBytes(module.Name));                   //Name
                bw.Write((ushort)0); //Null
            }
            bw.Write((ushort)0); //Null
            return ((MemoryStream)bw.BaseStream).ToArray();
        }       
        private byte[] CreateProjectStream()
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendFormat("ID=\"{0}\"\r\n", ProjectID);
            foreach(var module in Modules)
            {
                if (module.Type == eModuleType.Document)
                {
                    sb.AppendFormat("Document={0}/&H00000000\r\n", module.Name);
                }
                else if (module.Type == eModuleType.Module)
                {
                    sb.AppendFormat("Module={0}\r\n", module.Name);
                }
                else if (module.Type == eModuleType.Class)
                {
                    sb.AppendFormat("Class={0}\r\n", module.Name);
                }
                else
                {
                    //Designer
                    sb.AppendFormat("Package={0}\r\n", module.ClassID);
                    sb.AppendFormat("BaseClass={0}\r\n", module.Name);
                }
            }
            if (HelpFile1 != "")
            {
                sb.AppendFormat("HelpFile={0}\r\n", HelpFile1);
            }
            sb.AppendFormat("Name=\"{0}\"\r\n", Name);
            sb.AppendFormat("HelpContextID={0}\r\n", HelpContextID);

            if (!string.IsNullOrEmpty(Description))
            {
                sb.AppendFormat("Description=\"{0}\"\r\n", Description);
            }
            sb.AppendFormat("VersionCompatible32=\"393222000\"\r\n");

            sb.AppendFormat("CMG=\"{0}\"\r\n", WriteProtectionStat());
            sb.AppendFormat("DPB=\"{0}\"\r\n", WritePassword());
            sb.AppendFormat("GC=\"{0}\"\r\n\r\n", WriteVisibilityState());

            sb.Append("[Host Extender Info]\r\n");
            sb.Append("&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000\r\n");
            sb.Append("\r\n");
            sb.Append("[Workspace]\r\n");
            foreach(var module in Modules)
            {
                sb.AppendFormat("{0}=0, 0, 0, 0, C \r\n",module.Name);              
            }
            string s = sb.ToString();
            return Encoding.GetEncoding(CodePage).GetBytes(s);
        }
        private string WriteProtectionStat()
        {
            int stat=(_protection.UserProtected ? 1:0) |  
                     (_protection.HostProtected ? 2:0) |
                     (_protection.VbeProtected ? 4:0);

            return Encrypt(BitConverter.GetBytes(stat));    
        }
        private string WritePassword()
        {
            byte[] nullBits=new byte[3];
            byte[] nullKey = new byte[4];
            byte[] nullHash = new byte[20];
            if (Protection.PasswordKey == null)
            {
                return Encrypt(new byte[] { 0 });
            }
            else
            {
                Array.Copy(Protection.PasswordKey, nullKey, 4);
                Array.Copy(Protection.PasswordHash, nullHash, 20);

                //Set Null bits
                for (int i = 0; i < 24; i++)
                {
                    byte bit = (byte)(128 >> (int)((i % 8)));
                    if (i < 4)
                    {
                        if (nullKey[i] == 0)
                        {
                            nullKey[i] = 1;
                        }
                        else
                        {
                            nullBits[0] |= bit;
                        }
                    }
                    else
                    {
                        if (nullHash[i - 4] == 0)
                        {
                            nullHash[i - 4] = 1;
                        }
                        else
                        {
                            int byteIndex = (i - i % 8) / 8;
                            nullBits[byteIndex] |= bit;
                        }
                    }
                }
                //Write the Password Hash Data Structure (2.4.4.1)
                BinaryWriter bw = new BinaryWriter(new MemoryStream());
                bw.Write((byte)0xFF);
                bw.Write(nullBits);
                bw.Write(nullKey);
                bw.Write(nullHash);
                bw.Write((byte)0);
                return Encrypt(((MemoryStream)bw.BaseStream).ToArray());
            }
        }
        private string WriteVisibilityState()
        {
            return Encrypt(new byte[] { (byte)(Protection.VisibilityState ? 0xFF : 0) }); 
        }
        #endregion
        private string GetString(BinaryReader br, uint size)
        {
            return GetString(br, size, System.Text.Encoding.GetEncoding(CodePage));
        }
        private string GetString(BinaryReader br, uint size, Encoding enc)
        {
            if (size > 0)
            {
                byte[] byteTemp = new byte[size];
                byteTemp = br.ReadBytes((int)size);
                return enc.GetString(byteTemp);
            }
            else
            {
                return "";
            }
        }
        private string GetUnicodeString(BinaryReader br, uint size)
        {
            string s = GetString(br, size);
            int reserved = br.ReadUInt16();
            uint sizeUC = br.ReadUInt32();
            string sUC = GetString(br, sizeUC, System.Text.Encoding.Unicode);
            return sUC.Length == 0 ? s : sUC;
        }
        internal CompoundDocument Document { get; set; }
        internal Packaging.ZipPackagePart Part { get; set; }
        internal Uri Uri { get; private set; }
        /// <summary>
        /// Create a new VBA Project
        /// </summary>
        internal void Create()
        {
            if(Lcid>0)
            {
                throw (new InvalidOperationException("Package already contains a VBAProject"));
            }
            ProjectID = "{5DD90D76-4904-47A2-AF0D-D69B4673604E}";
            Name = "VBAProject";
            SystemKind = eSyskind.Win32;            //Default
            Lcid = 1033;                            //English - United States
            LcidInvoke = 1033;                      //English - United States
            CodePage = Encoding.GetEncoding(0).CodePage;    //Switched from Default to make it work in Core
            MajorVersion = 1361024421;
            MinorVersion = 6;
            HelpContextID = 0;
            Modules.Add(new ExcelVBAModule(_wb.CodeNameChange) { Name = "ThisWorkbook", Code = "", Attributes=GetDocumentAttributes("ThisWorkbook", "0{00020819-0000-0000-C000-000000000046}"), Type = eModuleType.Document, HelpContext = 0 });
            foreach (var sheet in _wb.Worksheets)
            {
                var name = GetModuleNameFromWorksheet(sheet);
                if (!Modules.Exists(name))
                {
                    Modules.Add(new ExcelVBAModule(sheet.CodeNameChange) { Name = name, Code = "", Attributes = GetDocumentAttributes(sheet.Name, "0{00020820-0000-0000-C000-000000000046}"), Type = eModuleType.Document, HelpContext = 0 });
                }
            }
            _protection = new ExcelVbaProtection(this) { UserProtected = false, HostProtected = false, VbeProtected = false, VisibilityState = true };
        }

        internal string GetModuleNameFromWorksheet(ExcelWorksheet sheet)
        {
            var name = sheet.Name;
            name = name.Substring(0, name.Length < 31 ? name.Length : 31);  //Maximum 31 charachters
            if (this.Modules[name] != null || !Regex.IsMatch(name, "^[a-zA-Z][a-zA-Z0-9_ ]*$")) //Check for valid chars, if not valid, set to sheetX.
            {
                int i = sheet.PositionID;
                name = "Sheet" + i.ToString();
                while (this.Modules[name] != null)
                {
                    name = "Sheet" + (++i).ToString(); ;
                }
            }            
            return name;
        }
        internal ExcelVbaModuleAttributesCollection GetDocumentAttributes(string name, string clsid)
        {
            var attr = new ExcelVbaModuleAttributesCollection();
            attr._list.Add(new ExcelVbaModuleAttribute() { Name = "VB_Name", Value = name, DataType = eAttributeDataType.String });
            attr._list.Add(new ExcelVbaModuleAttribute() { Name = "VB_Base", Value = clsid, DataType = eAttributeDataType.String });
            attr._list.Add(new ExcelVbaModuleAttribute() { Name = "VB_GlobalNameSpace", Value = "False", DataType = eAttributeDataType.NonString });
            attr._list.Add(new ExcelVbaModuleAttribute() { Name = "VB_Creatable", Value = "False", DataType = eAttributeDataType.NonString });
            attr._list.Add(new ExcelVbaModuleAttribute() { Name = "VB_PredeclaredId", Value = "True", DataType = eAttributeDataType.NonString });
            attr._list.Add(new ExcelVbaModuleAttribute() { Name = "VB_Exposed", Value = "False", DataType = eAttributeDataType.NonString });
            attr._list.Add(new ExcelVbaModuleAttribute() { Name = "VB_TemplateDerived", Value = "False", DataType = eAttributeDataType.NonString });
            attr._list.Add(new ExcelVbaModuleAttribute() { Name = "VB_Customizable", Value = "True", DataType = eAttributeDataType.NonString });

            return attr;
        }
        /// <summary>
        /// Remove the project from the package
        /// </summary>
        public void Remove()
        {
            if (Part == null) return;

            foreach (var rel in Part.GetRelationships())
            {
                _pck.DeleteRelationship(rel.Id);
            }
            if (_pck.PartExists(Uri))
            {
                _pck.DeletePart(Uri);
            }
            Part = null;
            Modules.Clear();
            References.Clear();
            Lcid = 0;
            LcidInvoke = 0;
            CodePage = 0;
            MajorVersion = 0;
            MinorVersion = 0;
            HelpContextID = 0;
        }
        public override string ToString()
        {
            return Name;
        }
    }
}
