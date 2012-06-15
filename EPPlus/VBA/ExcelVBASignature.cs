/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See http://www.codeplex.com/EPPlus for details.
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
 * Jan Källman		Added		26-MAR-2012
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Security.Cryptography.X509Certificates;
using System.Security.Cryptography.Pkcs;
using OfficeOpenXml.Utils;
using System.IO.Packaging;
using System.IO;

namespace OfficeOpenXml.VBA
{
    public class ExcelVbaSignature
    {
        const string schemaRelVbaSignature = "http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature";
        PackagePart _vbaPart = null;
        internal ExcelVbaSignature(PackagePart vbaPart)
        {
            _vbaPart = vbaPart;
            GetSignature();
        }
        private void GetSignature()
        {
            if (_vbaPart == null) return;
            var rel = _vbaPart.GetRelationshipsByType(schemaRelVbaSignature).FirstOrDefault();
            if (rel != null)
            {
                Uri = PackUriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);
                Part = _vbaPart.Package.GetPart(Uri);

                var stream = Part.GetStream();
                BinaryReader br = new BinaryReader(stream);
                uint cbSignature = br.ReadUInt32();        
                uint signatureOffset = br.ReadUInt32();     //44 ??
                uint cbSigningCertStore = br.ReadUInt32();  
                uint certStoreOffset = br.ReadUInt32();     
                uint cbProjectName = br.ReadUInt32();       
                uint projectNameOffset = br.ReadUInt32();   
                uint fTimestamp = br.ReadUInt32();
                uint cbTimestampUrl = br.ReadUInt32();
                uint timestampUrlOffset = br.ReadUInt32();  
                byte[] signature = br.ReadBytes((int)cbSignature);
                uint version = br.ReadUInt32();             
                uint fileType = br.ReadUInt32();            

                uint id = br.ReadUInt32();
                while (id != 0)
                {
                    uint encodingType = br.ReadUInt32();
                    uint length = br.ReadUInt32();
                    if (length > 0)
                    {
                        byte[] value = br.ReadBytes((int)length);
                        for (int i = 0; i < value.Length; i++) { System.Diagnostics.Debug.Write(",0x"+value[i].ToString("x")); }
                        switch (id)
                        {
                            //Add property values here...
                            case 0x20:
                                Certificate = new X509Certificate2(value);
                                break;
                            default:
                                break;
                        }
                    }
                    id = br.ReadUInt32();
                }
                uint endel1 = br.ReadUInt32();  //0
                uint endel2 = br.ReadUInt32();  //0
                ushort rgchProjectNameBuffer = br.ReadUInt16();
                ushort rgchTimestampBuffer = br.ReadUInt16();
                Verifier = new SignedCms();
                Verifier.Decode(signature);
            }
            else
            {
                Certificate = null;
                Verifier = null;
            }
        }
        //Create Oid from a bytearray
        //private string ReadHash(byte[] content)
        //{
        //    StringBuilder builder = new StringBuilder();
        //    int offset = 0x6;
        //    if (0 < (content.Length))
        //    {
        //        byte num = content[offset];
        //        byte num2 = (byte)(num / 40);
        //        builder.Append(num2.ToString(null, null));
        //        builder.Append(".");
        //        num2 = (byte)(num % 40);
        //        builder.Append(num2.ToString(null, null));
        //        ulong num3 = 0L;
        //        for (int i = offset + 1; i < content.Length; i++)
        //        {
        //            num2 = content[i];
        //            num3 = (ulong)(ulong)(num3 << 7) + ((byte)(num2 & 0x7f));
        //            if ((num2 & 0x80) == 0)
        //            {
        //                builder.Append(".");
        //                builder.Append(num3.ToString(null, null));
        //                num3 = 0L;
        //            }
        //            //1.2.840.113549.2.5
        //        }
        //    }


        //    string oId = builder.ToString();

        //    return oId;
        //}
        internal void Save(ExcelVbaProject proj)
        {
            if (Certificate == null || Certificate.HasPrivateKey==false)    //No signature. Remove any Signature part
            {
                if (Part != null)
                {
                    foreach (var r in Part.GetRelationships())
                    {
                        Part.DeleteRelationship(r.Id);
                    }
                    Part.Package.DeletePart(Part.Uri);
                }
                return;
            }
            var ms = new MemoryStream();
            var bw = new BinaryWriter(ms);

            byte[] certStore = GetCertStore();

            byte[] cert = SignProject(proj);
            bw.Write((uint)cert.Length);
            bw.Write((uint)44);                  //?? 36 ref inside cert ??
            bw.Write((uint)certStore.Length);    //cbSigningCertStore
            bw.Write((uint)(cert.Length + 44));  //certStoreOffset
            bw.Write((uint)0);                   //cbProjectName
            bw.Write((uint)(cert.Length + certStore.Length + 44));    //projectNameOffset
            bw.Write((uint)0);    //fTimestamp
            bw.Write((uint)0);    //cbTimestampUrl
            bw.Write((uint)(cert.Length + certStore.Length + 44 + 2));    //timestampUrlOffset
            bw.Write(cert);
            bw.Write(certStore);
            bw.Write((ushort)0);//rgchProjectNameBuffer
            bw.Write((ushort)0);//rgchTimestampBuffer
            bw.Write((ushort)0);
            bw.Flush();

            var rel = proj.Part.GetRelationshipsByType(schemaRelVbaSignature).FirstOrDefault();
            if (Part == null)
            {

                if (rel != null)
                {
                    Uri = rel.TargetUri;
                    Part = proj._pck.GetPart(rel.TargetUri);
                }
                else
                {
                    Uri = new Uri("/xl/vbaProjectSignature.bin", UriKind.Relative);
                    Part = proj._pck.CreatePart(Uri, ExcelPackage.schemaVBASignature);
                }
            }
            if (rel == null)
            {
                proj.Part.CreateRelationship(PackUriHelper.ResolvePartUri(proj.Uri, Uri), TargetMode.Internal, schemaRelVbaSignature);                
            }
            var b = ms.ToArray();
            Part.GetStream(FileMode.Create).Write(b, 0, b.Length);            
        }

        private byte[] GetCertStore()
        {
            var ms = new MemoryStream();
            var bw = new BinaryWriter(ms);

            bw.Write((uint)0); //Version
            bw.Write((uint)0x54524543); //fileType

            //SerializedCertificateEntry
            var certData = Certificate.RawData;
            bw.Write((uint)0x20);
            bw.Write((uint)1);
            bw.Write((uint)certData.Length);
            bw.Write(certData);

            //EndElementMarkerEntry
            bw.Write((uint)0);
            bw.Write((ulong)0);

            bw.Flush();
            return ms.ToArray();
        }

        private void WriteProp(BinaryWriter bw, int id, byte[] data)
        {
            bw.Write((uint)id);
            bw.Write((uint)1);
            bw.Write((uint)data.Length);
            bw.Write(data);
        }
        internal byte[] SignProject(ExcelVbaProject proj)
        {
            if (!Certificate.HasPrivateKey)
            {
                //throw (new InvalidOperationException("The certificate doesn't have a private key"));
                Certificate = null;
                return null;
            }
            var hash = GetContentHash(proj);

            BinaryWriter bw = new BinaryWriter(new MemoryStream());
            bw.Write((byte)0x30); //Constructed Type 
            bw.Write((byte)0x32); //Total length
            bw.Write((byte)0x30); //Constructed Type 
            bw.Write((byte)0x0E); //Length SpcIndirectDataContent
            bw.Write((byte)0x06); //Oid Tag Indentifier 
            bw.Write((byte)0x0A); //Lenght OId
            bw.Write(new byte[] { 0x2B, 0x06, 0x01, 0x04, 0x01, 0x82, 0x37, 0x02, 0x01, 0x1D }); //Encoded Oid 1.3.6.1.4.1.311.2.1.29
            bw.Write((byte)0x04);   //Octet String Tag Identifier
            bw.Write((byte)0x00);   //Zero length

            bw.Write((byte)0x30); //Constructed Type (DigestInfo)
            bw.Write((byte)0x20); //Length DigestInfo
            bw.Write((byte)0x30); //Constructed Type (Algorithm)
            bw.Write((byte)0x0C); //length AlgorithmIdentifier
            bw.Write((byte)0x06); //Oid Tag Indentifier 
            bw.Write((byte)0x08); //Lenght OId
            bw.Write(new byte[] { 0x2A, 0x86, 0x48, 0x86, 0xF7, 0x0D, 0x02, 0x05 }); //Encoded Oid for 1.2.840.113549.2.5 (AlgorithmIdentifier MD5)
            bw.Write((byte)0x05);   //Null type identifier
            bw.Write((byte)0x00);   //Null length
            bw.Write((byte)0x04);   //Octet String Identifier
            bw.Write((byte)hash.Length);   //Hash length
            bw.Write(hash);                //Content hash

            ContentInfo contentInfo = new ContentInfo(((MemoryStream)bw.BaseStream).ToArray());
            contentInfo.ContentType.Value = "1.3.6.1.4.1.311.2.1.4";
            Verifier = new SignedCms(contentInfo);
            var signer = new CmsSigner(Certificate);
            Verifier.ComputeSignature(signer, false);
            return Verifier.Encode();
        }

        private byte[] GetContentHash(ExcelVbaProject proj)
        {
            //MS-OVBA 2.4.2
            var enc = System.Text.Encoding.GetEncoding(proj.CodePage);
            BinaryWriter bw = new BinaryWriter(new MemoryStream());
            bw.Write(enc.GetBytes(proj.Name));
            bw.Write(enc.GetBytes(proj.Constants));
            foreach (var reference in proj.References)
            {
                if (reference.ReferenceRecordID == 0x0D)
                {
                    bw.Write((byte)0x7B);
                }
                if (reference.ReferenceRecordID == 0x0E)
                {
                    //var r = (ExcelVbaReferenceProject)reference;
                    //BinaryWriter bwTemp = new BinaryWriter(new MemoryStream());
                    //bwTemp.Write((uint)r.Libid.Length);
                    //bwTemp.Write(enc.GetBytes(r.Libid));              
                    //bwTemp.Write((uint)r.LibIdRelative.Length);
                    //bwTemp.Write(enc.GetBytes(r.LibIdRelative));
                    //bwTemp.Write(r.MajorVersion);
                    //bwTemp.Write(r.MinorVersion);
                    foreach (byte b in BitConverter.GetBytes((uint)reference.Libid.Length))  //Length will never be an UInt with 4 bytes that aren't 0 (> 0x00FFFFFF), so no need for the rest of the properties.
                    {
                        if (b != 0)
                        {
                            bw.Write(b);
                        }
                        else
                        {
                            break;
                        }
                    }
                }
            }
            foreach (var module in proj.Modules)
            {
                var lines = module.Code.Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var line in lines)
                {
                    if (!line.StartsWith("attribute", true, null))
                    {
                        bw.Write(enc.GetBytes(line));
                    }
                }
            }
            var buffer = (bw.BaseStream as MemoryStream).ToArray();
            var hp = System.Security.Cryptography.MD5CryptoServiceProvider.Create();
            return hp.ComputeHash(buffer);
        }
        /// <summary>
        /// The certificate to sign the VBA project.
        /// <remarks>
        /// This certificate must have a private key.
        /// There is no validation that the certificate is valid for codesigning, so make sure it's valid to sign Excel files with (Excel 2010 is more strict that prior versions).
        /// </remarks>
        /// </summary>
        public X509Certificate2 Certificate { get; set; }
        /// <summary>
        /// The verifier
        /// </summary>
        public SignedCms Verifier { get; internal set; }
        internal CompoundDocument Signature { get; set; }
        internal PackagePart Part { get; set; }
        internal Uri Uri { get; private set; }
    }
}
