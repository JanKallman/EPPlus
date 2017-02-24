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
 **************************************************************************************
 * This class is created with the help of the MS-OFFCRYPTO PDF documentation... http://msdn.microsoft.com/en-us/library/cc313071(office.12).aspx
 * Decryption library for Office Open XML files(Lyquidity) and Sminks very nice example 
 * on "Reading compound documents in c#" on Stackoverflow. Many thanks!
 ***************************************************************************************
 * Code change notes:
 * 
 * Author							Change						Date
 *******************************************************************************
 * Jan Källman		Added		10-AUG-2010
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices; 
using comTypes=System.Runtime.InteropServices.ComTypes;
using System.IO;
using System.Security.Cryptography;
using System.Xml;
namespace OfficeOpenXml.Utils
{
    [ComImport] 
    [Guid("0000000d-0000-0000-C000-000000000046")] 
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)] 
    internal interface IEnumSTATSTG 
    { 
        // The user needs to allocate an STATSTG array whose size is celt. 
        [PreserveSig] 
        uint Next( 
            uint celt, 
            [MarshalAs(UnmanagedType.LPArray), Out] 
            System.Runtime.InteropServices.ComTypes.STATSTG[] rgelt, 
            out uint pceltFetched 
        ); 
 
        void Skip(uint celt); 
 
        void Reset(); 
 
        [return: MarshalAs(UnmanagedType.Interface)] 
        IEnumSTATSTG Clone(); 
    } 
 
    [ComImport] 
    [Guid("0000000b-0000-0000-C000-000000000046")] 
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)] 
    interface IStorage 
    { 
        void CreateStream( 
            /* [string][in] */ string pwcsName, 
            /* [in] */ uint grfMode, 
            /* [in] */ uint reserved1, 
            /* [in] */ uint reserved2, 
            /* [out] */ out comTypes.IStream ppstm); 
        
        void OpenStream( 
            /* [string][in] */ string pwcsName, 
            /* [unique][in] */ IntPtr reserved1, 
            /* [in] */ uint grfMode, 
            /* [in] */ uint reserved2,
            /* [out] */ out comTypes.IStream ppstm); 
 
        void CreateStorage( 
            /* [string][in] */ string pwcsName, 
            /* [in] */ uint grfMode, 
            /* [in] */ uint reserved1, 
            /* [in] */ uint reserved2, 
            /* [out] */ out IStorage ppstg); 
 
        void OpenStorage( 
            /* [string][unique][in] */ string pwcsName, 
            /* [unique][in] */ IStorage pstgPriority, 
            /* [in] */ uint grfMode, 
            /* [unique][in] */ IntPtr snbExclude, 
            /* [in] */ uint reserved, 
            /* [out] */ out IStorage ppstg);

        void CopyTo(
            [InAttribute] uint ciidExclude,
            [InAttribute] Guid[] rgiidExclude,
            [InAttribute] IntPtr snbExclude,
            [InAttribute] IStorage pstgDest
        );
 
        void MoveElementTo( 
            /* [string][in] */ string pwcsName, 
            /* [unique][in] */ IStorage pstgDest, 
            /* [string][in] */ string pwcsNewName, 
            /* [in] */ uint grfFlags); 
 
        void Commit( 
            /* [in] */ uint grfCommitFlags); 
 
        void Revert(); 
 
        void EnumElements( 
            /* [in] */ uint reserved1, 
            /* [size_is][unique][in] */ IntPtr reserved2, 
            /* [in] */ uint reserved3, 
            /* [out] */ out IEnumSTATSTG ppenum); 
 
        void DestroyElement( 
            /* [string][in] */ string pwcsName); 
 
        void RenameElement( 
            /* [string][in] */ string pwcsOldName, 
            /* [string][in] */ string pwcsNewName); 
 
        void SetElementTimes( 
            /* [string][unique][in] */ string pwcsName, 
            /* [unique][in] */ System.Runtime.InteropServices.ComTypes.FILETIME pctime, 
            /* [unique][in] */ System.Runtime.InteropServices.ComTypes.FILETIME patime, 
            /* [unique][in] */ System.Runtime.InteropServices.ComTypes.FILETIME pmtime); 
 
        void SetClass( 
            /* [in] */ Guid clsid); 
 
        void SetStateBits( 
            /* [in] */ uint grfStateBits, 
            /* [in] */ uint grfMask); 
 
        void Stat( 
            /* [out] */ out System.Runtime.InteropServices.ComTypes.STATSTG pstatstg, 
            /* [in] */ uint grfStatFlag); 
 
    }
    [ComVisible(false)]
    [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("0000000A-0000-0000-C000-000000000046")]
    internal interface ILockBytes
    {
        void ReadAt(long ulOffset, System.IntPtr pv, int cb, out UIntPtr pcbRead);
        void WriteAt(long ulOffset, System.IntPtr pv, int cb, out UIntPtr pcbWritten);
        void Flush();
        void SetSize(long cb);
        void LockRegion(long libOffset, long cb, int dwLockType);
        void UnlockRegion(long libOffset, long cb, int dwLockType);
        void Stat(out System.Runtime.InteropServices.ComTypes.STATSTG pstatstg, int grfStatFlag);
    }
    [Flags] 
    internal enum STGM : int 
    { 
        DIRECT = 0x00000000, 
        TRANSACTED = 0x00010000, 
        SIMPLE = 0x08000000, 
        READ = 0x00000000, 
        WRITE = 0x00000001, 
        READWRITE = 0x00000002, 
        SHARE_DENY_NONE = 0x00000040, 
        SHARE_DENY_READ = 0x00000030, 
        SHARE_DENY_WRITE = 0x00000020, 
        SHARE_EXCLUSIVE = 0x00000010, 
        PRIORITY = 0x00040000, 
        DELETEONRELEASE = 0x04000000, 
        NOSCRATCH = 0x00100000, 
        CREATE = 0x00001000, 
        CONVERT = 0x00020000, 
        FAILIFTHERE = 0x00000000, 
        NOSNAPSHOT = 0x00200000, 
        DIRECT_SWMR = 0x00400000, 
    } 
 
    internal enum STATFLAG : uint 
    { 
        STATFLAG_DEFAULT = 0, 
        STATFLAG_NONAME = 1, 
        STATFLAG_NOOPEN = 2 
    } 
 
    internal enum STGTY : int 
    { 
        STGTY_STORAGE = 1, 
        STGTY_STREAM = 2, 
        STGTY_LOCKBYTES = 3, 
        STGTY_PROPERTY = 4 
    }
}
