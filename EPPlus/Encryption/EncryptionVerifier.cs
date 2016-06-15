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
    /// <summary>
    /// Encryption verifier inside the EncryptionInfo stream
    /// </summary>
    internal class EncryptionVerifier
    {
        internal uint SaltSize;              // An unsigned integer that specifies the size of the Salt field. It MUST be 0x00000010.
        internal byte[] Salt;                //(16 bytes): An array of bytes that specifies the salt value used during password hash generation. It MUST NOT be the same data used for the verifier stored encrypted in the EncryptedVerifier field.
        internal byte[] EncryptedVerifier;   //(16 bytes): MUST be the randomly generated Verifier value encrypted using the algorithm chosen by the implementation.
        internal uint VerifierHashSize;      //(4 bytes): An unsigned integer that specifies the number of bytes needed to contain the hash of the data used to generate the EncryptedVerifier field.
        internal byte[] EncryptedVerifierHash; //(variable): An array of bytes that contains the encrypted form of the hash of the randomly generated Verifier value. The length of the array MUST be the size of the encryption block size multiplied by the number of blocks needed to encrypt the hash of the Verifier. If the encryption algorithm is RC4, the length MUST be 20 bytes. If the encryption algorithm is AES, the length MUST be 32 bytes.
        internal byte[] WriteBinary()
        {
            MemoryStream ms = new MemoryStream();
            BinaryWriter bw = new BinaryWriter(ms);

            bw.Write(SaltSize);
            bw.Write(Salt);
            bw.Write(EncryptedVerifier);
            bw.Write(0x14);                 //Sha1 is 20 bytes  (Encrypted is 32)
            bw.Write(EncryptedVerifierHash);

            bw.Flush();
            return ms.ToArray();
        }
    }
}
