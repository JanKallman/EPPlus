// ComHelper.cs
// ------------------------------------------------------------------
//
// Copyright (c) 2009 Dino Chiesa.
// All rights reserved.
//
// This code module is part of DotNetZip, a zipfile class library.
//
// ------------------------------------------------------------------
//
// This code is licensed under the Microsoft Public License.
// See the file License.txt for the license details.
// More info on: http://dotnetzip.codeplex.com
//
// ------------------------------------------------------------------
//
// last saved (in emacs):
// Time-stamp: <2011-June-13 17:04:06>
//
// ------------------------------------------------------------------
//
// This module defines a COM Helper class.
//
// Created: Tue, 08 Sep 2009  22:03
//

using Interop=System.Runtime.InteropServices;

namespace Ionic.Zip
{
    /// <summary>
    /// This class exposes a set of COM-accessible wrappers for static
    /// methods available on the ZipFile class.  You don't need this
    /// class unless you are using DotNetZip from a COM environment.
    /// </summary>
    [System.Runtime.InteropServices.GuidAttribute("ebc25cf6-9120-4283-b972-0e5520d0000F")]
    [System.Runtime.InteropServices.ComVisible(true)]
#if !NETCF
    [System.Runtime.InteropServices.ClassInterface(System.Runtime.InteropServices.ClassInterfaceType.AutoDispatch)]
#endif

    internal class ComHelper
    {
        /// <summary>
        ///  A wrapper for <see cref="ZipFile.IsZipFile(string)">ZipFile.IsZipFile(string)</see>
        /// </summary>
        /// <param name="filename">The filename to of the zip file to check.</param>
        /// <returns>true if the file contains a valid zip file.</returns>
        public bool IsZipFile(string filename)
        {
            return ZipFile.IsZipFile(filename);
        }

        /// <summary>
        ///  A wrapper for <see cref="ZipFile.IsZipFile(string, bool)">ZipFile.IsZipFile(string, bool)</see>
        /// </summary>
        /// <remarks>
        /// We cannot use "overloaded" Method names in COM interop.
        /// So, here, we use a unique name.
        /// </remarks>
        /// <param name="filename">The filename to of the zip file to check.</param>
        /// <returns>true if the file contains a valid zip file.</returns>
        public bool IsZipFileWithExtract(string filename)
        {
            return ZipFile.IsZipFile(filename, true);
        }

#if !NETCF
        /// <summary>
        ///  A wrapper for <see cref="ZipFile.CheckZip(string)">ZipFile.CheckZip(string)</see>
        /// </summary>
        /// <param name="filename">The filename to of the zip file to check.</param>
        ///
        /// <returns>true if the named zip file checks OK. Otherwise, false. </returns>
        public bool CheckZip(string filename)
        {
            return ZipFile.CheckZip(filename);
        }

        /// <summary>
        ///  A COM-friendly wrapper for the static method <see cref="ZipFile.CheckZipPassword(string,string)"/>.
        /// </summary>
        ///
        /// <param name="filename">The filename to of the zip file to check.</param>
        ///
        /// <param name="password">The password to check.</param>
        ///
        /// <returns>true if the named zip file checks OK. Otherwise, false. </returns>
        public bool CheckZipPassword(string filename, string password)
        {
            return ZipFile.CheckZipPassword(filename, password);
        }

        /// <summary>
        ///  A wrapper for <see cref="ZipFile.FixZipDirectory(string)">ZipFile.FixZipDirectory(string)</see>
        /// </summary>
        /// <param name="filename">The filename to of the zip file to fix.</param>
        public void FixZipDirectory(string filename)
        {
            ZipFile.FixZipDirectory(filename);
        }
#endif

        /// <summary>
        ///  A wrapper for <see cref="ZipFile.LibraryVersion">ZipFile.LibraryVersion</see>
        /// </summary>
        /// <returns>
        ///  the version number on the DotNetZip assembly, formatted as a string.
        /// </returns>
        public string GetZipLibraryVersion()
        {
            return ZipFile.LibraryVersion.ToString();
        }

    }
}