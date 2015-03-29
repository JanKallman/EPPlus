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
 * ******************************************************************************
 * Jan Källman		    Added       		        2013-01-05
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Encryption
{
    [Flags]
    internal enum Flags
    {
        Reserved1 = 1,   // (1 bit): MUST be set to zero, and MUST be ignored.
        Reserved2 = 2,   // (1 bit): MUST be set to zero, and MUST be ignored.
        fCryptoAPI = 4,   // (1 bit): A flag that specifies whether CryptoAPI RC4 or [ECMA-376] encryption is used. It MUST be set to 1 unless fExternal is 1. If fExternal is set to 1, it MUST be set to zero.        
        fDocProps = 8,   // (1 bit): MUST be set to zero if document properties are encrypted. Otherwise, it MUST be set to 1. Encryption of document properties is specified in section 2.3.5.4.
        fExternal = 16,  // (1 bit): If extensible encryption is used, it MUST be set to 1. Otherwise, it MUST be set to zero. If this field is set to 1, all other fields in this structure MUST be set to zero.
        fAES = 32   //(1 bit): If the protected content is an [ECMA-376] document, it MUST be set to 1. Otherwise, it MUST be set to zero. If the fAES bit is set to 1, the fCryptoAPI bit MUST also be set to 1
    }
    internal enum AlgorithmID
    {
        Flags = 0x00000000,   // Determined by Flags
        RC4 = 0x00006801,   // RC4
        AES128 = 0x0000660E,   // 128-bit AES
        AES192 = 0x0000660F,   // 192-bit AES
        AES256 = 0x00006610    // 256-bit AES
    }
    internal enum AlgorithmHashID
    {
        App = 0x00000000,
        SHA1 = 0x00008004,
    }
    internal enum ProviderType
    {
        Flags = 0x00000000,//Determined by Flags
        RC4 = 0x00000001,
        AES = 0x00000018,
    }
    /// <summary>
    /// Encryption Header inside the EncryptionInfo stream
    /// </summary>
    internal class EncryptionHeader
    {
        internal Flags Flags;
        internal int SizeExtra;             //MUST be 0x00000000.
        internal AlgorithmID AlgID;         //MUST be 0x0000660E (AES-128), 0x0000660F (AES-192), or 0x00006610 (AES-256).
        internal AlgorithmHashID AlgIDHash; //MUST be 0x00008004 (SHA-1).
        internal int KeySize;               //MUST be 0x00000080 (AES-128), 0x000000C0 (AES-192), or 0x00000100 (AES-256).
        internal ProviderType ProviderType; //SHOULD<10> be 0x00000018 (AES).
        internal int Reserved1;             //Undefined and MUST be ignored.
        internal int Reserved2;             //MUST be 0x00000000 and MUST be ignored.
        internal string CSPName;            //SHOULD<11> be set to either "Microsoft Enhanced RSA and AES Cryptographic Provider" or "Microsoft Enhanced RSA and AES Cryptographic Provider (Prototype)" as a null-terminated Unicode string.
        internal byte[] WriteBinary()
        {
            MemoryStream ms = new MemoryStream();
            BinaryWriter bw = new BinaryWriter(ms);

            bw.Write((int)Flags);
            bw.Write(SizeExtra);
            bw.Write((int)AlgID);
            bw.Write((int)AlgIDHash);
            bw.Write((int)KeySize);
            bw.Write((int)ProviderType);
            bw.Write(Reserved1);
            bw.Write(Reserved2);
            bw.Write(Encoding.Unicode.GetBytes(CSPName));

            bw.Flush();
            return ms.ToArray();
        }
    }
}
