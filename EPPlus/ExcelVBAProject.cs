using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO.Packaging;
using System.IO;
using OfficeOpenXml.Utils;
using System.Security.Cryptography.Pkcs;
using System.Security.Cryptography.X509Certificates;

namespace OfficeOpenXml
{
    public class ExcelVBAProject
    {
        const string schemaRelVba = "http://schemas.microsoft.com/office/2006/relationships/vbaProject";
        const string schemaRelVbaSignature = "http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature";
        #region Classes & Enums
        public class VbaReference
        {
            public string Name { get; set; }
            public string TypeLibraryID { get; set; }
            public string LibidOriginal { get; set; }
            public string LibidTwiddled { get; set; }
            public override string ToString()
            {
                return Name;
            }
        }
        public class VBAModual
        {
            public string Name { get; set; }
            public string Description { get; set; }
            public string Code { get; set; }
            public int HelpContext { get; set; }

            internal string streamName { get; set; }
            public ushort Cookie { get; set; }
            public uint ModuleOffset { get; set; }
            public override string ToString()
            {
                return Name;
            }
        }
        public class ExcelVBAModuleName
        {
            public ExcelVBAModuleName(byte[] part)
            {
                uint sizeOfName = BitConverter.ToUInt32(part, 2);

                byte[] byName = new byte[sizeOfName];
                string name = Encoding.UTF8.GetString(byName);
            }

        }
        public class ExcelVBASignature
        {
            PackagePart _vbaPart=null;
            public ExcelVBASignature(PackagePart vbaPart)
	        {
                _vbaPart = vbaPart;
                GetSignature();
	        }
            private void GetSignature()
            {
                var rel = _vbaPart.GetRelationshipsByType(schemaRelVbaSignature).FirstOrDefault();
                if (rel != null)
                {
                    var uri = PackUriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);
                    var part = _vbaPart.Package.GetPart(uri);

                    var stream = part.GetStream();
                    BinaryReader br = new BinaryReader(stream);
                    uint cbSignature = br.ReadUInt32();
                    uint signatureOffset = br.ReadUInt32();
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
                            switch (id)
                            {
                                //Add property values here...
                                case 0x20:
                                    Certificate = new System.Security.Cryptography.X509Certificates.X509Certificate2(value);
                                    break;
                                default:
                                    break;
                            }
                        }
                        id = br.ReadUInt32();
                    }
                    uint endel1 = br.ReadUInt32();
                    uint endel2 = br.ReadUInt32();
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
            internal void WriteSignature()
            {
                var ms=new MemoryStream();
                var bw=new BinaryWriter(ms);

                byte[] certStore = GetCertStore();

                byte[] cert = SignProject();
                bw.Write((UInt32)cert.Length);
                bw.Write((UInt32)36);
                bw.Write((UInt32)certStore.Length);    //cbSigningCertStore
                bw.Write((UInt32)cert.Length+36);    //certStoreOffset
                bw.Write((UInt32)0);    //cbProjectName
                bw.Write((UInt32)cert.Length+certStore.Length+36);    //projectNameOffset
                bw.Write((UInt32)0);    //fTimestamp
                bw.Write((UInt32)0);    //cbTimestampUrl
                bw.Write((UInt32)cert.Length + certStore.Length + 36);    //timestampUrlOffset
                bw.Write(cert);
                bw.Write(certStore);
                bw.Write((short)0);//rgchProjectNameBuffer
                bw.Write((short)0);//rgchTimestampBuffer

                bw.Write(Verifier.Encode());
                bw.Flush();

                if (Part == null)
                {
                    Uri=new Uri("/xl/vbaProjectSignature.bin");
                    Part = _vbaPart.Package.CreatePart(Uri, ExcelPackage.schemaVBASignature);
                }
                var rel = Part.GetRelationshipsByType(schemaRelVbaSignature).FirstOrDefault();
                if (rel != null)
                {
                    _vbaPart.CreateRelationship(PackUriHelper.ResolvePartUri(_vbaPart.Uri, Uri), TargetMode.Internal, schemaRelVbaSignature)                     ;
                }
                Part.GetStream().Write(ms.ToArray(), 0, 0);
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
                bw.Write((UInt64)0);

                bw.Flush();
                return ms.ToArray();
            }
            internal byte[] SignProject()
            {
                if (!Certificate.HasPrivateKey)
                {
                    throw (new InvalidOperationException("The certificate don't have a private key"));
                }
                var stream=_vbaPart.GetStream();
                byte[] cont=new byte[stream.Length];
                stream.Read(cont,0,(int)stream.Length);
                ContentInfo contentInfo = new ContentInfo(cont);
                contentInfo.ContentType.Value = "1.3.6.1.4.1.311.2.1.4";
                Verifier = new SignedCms(contentInfo);
                var signer = new CmsSigner(Certificate);                
                Verifier.ComputeSignature(signer, true);
                return Verifier.Encode();            
            }
            public X509Certificate2 Certificate {get;set;}
            public SignedCms Verifier {get;set;}
            internal CompoundDocument Signature { get; set; }
            internal PackagePart Part { get; set; }
            internal Uri Uri { get; private set; }        
        }
        public enum eSyskind
        {
            Win16 = 0,
            Win32 = 1,
            Macintosh = 2,
            Win64 = 3
        }

        #endregion
        public ExcelVBAProject(ExcelWorkbook wb)
        {
            _wb = wb;
            _pck = _wb._package.Package;
            var rel = _wb.Part.GetRelationshipsByType(schemaRelVba).FirstOrDefault();
            if (rel != null)
            {
                Uri = PackUriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);
                Part = _pck.GetPart(Uri);
                GetProject();                
            }
            else
            {
                Part = null;
            }
        }
        ExcelWorkbook _wb;
        Package _pck;
        #region Dir Stream Properties
        public eSyskind SystemKind { get; set; }
        public string ProjectName { get; set; }
        public string Description { get; set; }
        public string HelpFile1 { get; set; }
        public string HelpFile2 { get; set; }
        public int HelpContextID { get; set; }
        public string Constants { get; set; }
        internal int LibFlags { get; set; }
        internal int MajorVersion { get; set; }
        internal int MinorVersion { get; set; }
        internal int Lcid { get; set; }
        internal int LcidInvoke { get; set; }
        internal int CodePage { get; set; }
        internal string ProjectStreamText { get; set; }
        //internal string ProjectlkText { get; set; }
        //internal string ProjectwmText { get; set; }
        public List<VbaReference> References { get; set; }
        public List<VBAModual> Moduls { get; set; }
        ExcelVBASignature _signature = null;
        public ExcelVBASignature Signature
        {
            get
            {
                if (_signature == null)
                {
                    _signature=new ExcelVBASignature(Part);
                }
                return _signature;
            }
        }
        #endregion
        private void GetProject()
        {

            var stream = Part.GetStream();
            byte[] vba;
            vba = new byte[stream.Length];
            stream.Read(vba, 0, (int)stream.Length);

            Document = new CompoundDocument(vba);
            ReadDirStream();
            ProjectStreamText = Encoding.GetEncoding(CodePage).GetString(Document.Storage.DataStreams["PROJECT"]);
            ReadModuls();
            foreach (var key in Document.Storage.SubStorage.Keys)
            {
                if (key != "VBA")
                {
                    var st = Document.Storage.SubStorage[key];
                    string vbFrame = Encoding.GetEncoding(CodePage).GetString(st.DataStreams["\x3VBFrame"]);
                }
            }
        }
        private void ReadModuls()
        {
            foreach (var modul in Moduls)
            {
                var stream = Document.Storage.SubStorage["VBA"].DataStreams[modul.streamName];
                var byCode = CompoundDocument.DecompressPart(stream, (int)modul.ModuleOffset);
                modul.Code = Encoding.GetEncoding(CodePage).GetString(byCode);
            }
        }
        private void ReadDirStream()
        {
            byte[] dir = CompoundDocument.DecompressPart(Document.Storage.SubStorage["VBA"].DataStreams["dir"]);
            MemoryStream ms = new MemoryStream(dir);
            BinaryReader br = new BinaryReader(ms);
            VbaReference currentRef = null;
            VBAModual currentModual = null;
            References = new List<VbaReference>();
            Moduls = new List<VBAModual>();
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
                        ProjectName = GetString(br, size);
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
                        currentRef.TypeLibraryID = GetString(br, sizeLibID);
                        uint reserved1 = br.ReadUInt32();
                        ushort reserved2 = br.ReadUInt16();
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
                        currentRef = new VbaReference();
                        currentRef.Name = GetUnicodeString(br, size);
                        References.Add(currentRef);
                        break;
                    case 0x19:
                        currentModual = new VBAModual();
                        currentModual.Name = GetUnicodeString(br, size);
                        Moduls.Add(currentModual);
                        break;
                    case 0x1A:
                        currentModual.streamName = GetUnicodeString(br, size);
                        break;
                    case 0x1C:
                        currentModual.Description = GetUnicodeString(br, size);
                        break;
                    case 0x1E:
                        currentModual.HelpContext = (int)br.ReadUInt32();
                        break;
                    case 0x21:
                    case 0x22:
                        bool isProcedural = (id == 22);
                        break;
                    case 0x2B:      //Modul Terminator
                        break;
                    case 0x2C:
                        currentModual.Cookie = br.ReadUInt16();
                        break;
                    case 0x31:
                        currentModual.ModuleOffset = br.ReadUInt32();
                        break;
                    case 0x10:
                        terminate = true;
                        break;
                    case 0x30:
                        uint sizeExt = br.ReadUInt32();
                        currentRef.LibidTwiddled = GetString(br, sizeExt);
                        uint reserved3 = br.ReadUInt32();
                        ushort reserved4 = br.ReadUInt16();
                        byte[] guid = br.ReadBytes(16);
                        uint extCookie = br.ReadUInt32();
                        break;
                    case 0x33:
                        currentRef.LibidOriginal = GetString(br, size);
                        break;
                    case 0x2F:
                        uint sizetLibID = br.ReadUInt32();
                        currentRef.TypeLibraryID = GetString(br, sizetLibID);
                        reserved1 = br.ReadUInt32();
                        reserved2 = br.ReadUInt16();

                        break;

                    default:
                        break;
                }
            }
        }
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
        internal PackagePart Part { get; set; }
        internal Uri Uri { get; private set; }        
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
        }
    }
}
