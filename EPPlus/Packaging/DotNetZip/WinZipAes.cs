//#define Trace

// WinZipAes.cs
// ------------------------------------------------------------------
//
// Copyright (c) 2009-2011 Dino Chiesa.
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
// Time-stamp: <2011-July-12 13:42:06>
//
// ------------------------------------------------------------------
//
// This module defines the classes for dealing with WinZip's AES encryption,
// according to the specifications for the format available on WinZip's website.
//
// Created: January 2009
//
// ------------------------------------------------------------------

using System;
using System.IO;
using System.Collections.Generic;
using System.Security.Cryptography;

#if AESCRYPTO
namespace OfficeOpenXml.Packaging.Ionic.Zip
{
    /// <summary>
    ///   This is a helper class supporting WinZip AES encryption.
    ///   This class is intended for use only by the DotNetZip library.
    /// </summary>
    ///
    /// <remarks>
    ///   Most uses of the DotNetZip library will not involve direct calls into
    ///   the WinZipAesCrypto class.  Instead, the WinZipAesCrypto class is
    ///   instantiated and used by the ZipEntry() class when WinZip AES
    ///   encryption or decryption on an entry is employed.
    /// </remarks>
    internal class WinZipAesCrypto
    {
        internal byte[] _Salt;
        internal byte[] _providedPv;
        internal byte[] _generatedPv;
        internal int _KeyStrengthInBits;
        private byte[] _MacInitializationVector;
        private byte[] _StoredMac;
        private byte[] _keyBytes;
        private Int16 PasswordVerificationStored;
        private Int16 PasswordVerificationGenerated;
        private int Rfc2898KeygenIterations = 1000;
        private string _Password;
        private bool _cryptoGenerated ;

        private WinZipAesCrypto(string password, int KeyStrengthInBits)
        {
            _Password = password;
            _KeyStrengthInBits = KeyStrengthInBits;
        }

        public static WinZipAesCrypto Generate(string password, int KeyStrengthInBits)
        {
            WinZipAesCrypto c = new WinZipAesCrypto(password, KeyStrengthInBits);

            int saltSizeInBytes = c._KeyStrengthInBytes / 2;
            c._Salt = new byte[saltSizeInBytes];
            Random rnd = new Random();
            rnd.NextBytes(c._Salt);
            return c;
        }



        public static WinZipAesCrypto ReadFromStream(string password, int KeyStrengthInBits, Stream s)
        {
            // from http://www.winzip.com/aes_info.htm
            //
            // Size(bytes)   Content
            // -----------------------------------
            // Variable      Salt value
            // 2             Password verification value
            // Variable      Encrypted file data
            // 10            Authentication code
            //
            // ZipEntry.CompressedSize represents the size of all of those elements.

            // salt size varies with key length:
            //    128 bit key => 8 bytes salt
            //    192 bits => 12 bytes salt
            //    256 bits => 16 bytes salt

            WinZipAesCrypto c = new WinZipAesCrypto(password, KeyStrengthInBits);

            int saltSizeInBytes = c._KeyStrengthInBytes / 2;
            c._Salt = new byte[saltSizeInBytes];
            c._providedPv = new byte[2];

            s.Read(c._Salt, 0, c._Salt.Length);
            s.Read(c._providedPv, 0, c._providedPv.Length);

            c.PasswordVerificationStored = (Int16)(c._providedPv[0] + c._providedPv[1] * 256);
            if (password != null)
            {
                c.PasswordVerificationGenerated = (Int16)(c.GeneratedPV[0] + c.GeneratedPV[1] * 256);
                if (c.PasswordVerificationGenerated != c.PasswordVerificationStored)
                    throw new BadPasswordException("bad password");
            }

            return c;
        }

        public byte[] GeneratedPV
        {
            get
            {
                if (!_cryptoGenerated) _GenerateCryptoBytes();
                return _generatedPv;
            }
        }


        public byte[] Salt
        {
            get
            {
                return _Salt;
            }
        }


        private int _KeyStrengthInBytes
        {
            get
            {
                return _KeyStrengthInBits / 8;

            }
        }

        public int SizeOfEncryptionMetadata
        {
            get
            {
                // 10 bytes after, (n-10) before the compressed data
                return _KeyStrengthInBytes / 2 + 10 + 2;
            }
        }

        public string Password
        {
            set
            {
                _Password = value;
                if (_Password != null)
                {
                    PasswordVerificationGenerated = (Int16)(GeneratedPV[0] + GeneratedPV[1] * 256);
                    if (PasswordVerificationGenerated != PasswordVerificationStored)
                        throw new Ionic.Zip.BadPasswordException();
                }
            }
            private get
            {
                return _Password;
            }
        }


        private void _GenerateCryptoBytes()
        {
            //Console.WriteLine(" provided password: '{0}'", _Password);

            System.Security.Cryptography.Rfc2898DeriveBytes rfc2898 =
                new System.Security.Cryptography.Rfc2898DeriveBytes(_Password, Salt, Rfc2898KeygenIterations);

            _keyBytes = rfc2898.GetBytes(_KeyStrengthInBytes); // 16 or 24 or 32 ???
            _MacInitializationVector = rfc2898.GetBytes(_KeyStrengthInBytes);
            _generatedPv = rfc2898.GetBytes(2);

            _cryptoGenerated = true;
        }


        public byte[] KeyBytes
        {
            get
            {
                if (!_cryptoGenerated) _GenerateCryptoBytes();
                return _keyBytes;
            }
        }


        public byte[] MacIv
        {
            get
            {
                if (!_cryptoGenerated) _GenerateCryptoBytes();
                return _MacInitializationVector;
            }
        }

        public byte[] CalculatedMac;


        public void ReadAndVerifyMac(System.IO.Stream s)
        {
            bool invalid = false;

            // read integrityCheckVector.
            // caller must ensure that the file pointer is in the right spot!
            _StoredMac = new byte[10];  // aka "authentication code"
            s.Read(_StoredMac, 0, _StoredMac.Length);

            if (_StoredMac.Length != CalculatedMac.Length)
                invalid = true;

            if (!invalid)
            {
                for (int i = 0; i < _StoredMac.Length; i++)
                {
                    if (_StoredMac[i] != CalculatedMac[i])
                        invalid = true;
                }
            }

            if (invalid)
                throw new Ionic.Zip.BadStateException("The MAC does not match.");
        }

    }


    #region DONT_COMPILE_BUT_KEEP_FOR_POTENTIAL_FUTURE_USE
#if NO
    internal class Util
    {
        private static void _Format(System.Text.StringBuilder sb1,
                                    byte[] b,
                                    int offset,
                                    int length)
        {

            System.Text.StringBuilder sb2 = new System.Text.StringBuilder();
            sb1.Append("0000    ");
            int i;
            for (i = 0; i < length; i++)
            {
                int x = offset+i;
                if (i != 0 && i % 16 == 0)
                {
                    sb1.Append("    ")
                        .Append(sb2)
                        .Append("\n")
                        .Append(String.Format("{0:X4}    ", i));
                    sb2.Remove(0,sb2.Length);
                }
                sb1.Append(System.String.Format("{0:X2} ", b[x]));
                if (b[x] >=32 && b[x] <= 126)
                    sb2.Append((char)b[x]);
                else
                    sb2.Append(".");
            }
            if (sb2.Length > 0)
            {
                sb1.Append(new String(' ', ((16 - i%16) * 3) + 4))
                    .Append(sb2);
            }
        }



        internal static string FormatByteArray(byte[] b, int limit)
        {
            System.Text.StringBuilder sb1 = new System.Text.StringBuilder();

            if ((limit * 2 > b.Length) || limit == 0)
            {
                _Format(sb1, b, 0, b.Length);
            }
            else
            {
                // first N bytes of the buffer
                _Format(sb1, b, 0, limit);

                if (b.Length > limit * 2)
                    sb1.Append(String.Format("\n   ...({0} other bytes here)....\n", b.Length - limit * 2));

                // last N bytes of the buffer
                _Format(sb1, b, b.Length - limit, limit);
            }

            return sb1.ToString();
        }


        internal static string FormatByteArray(byte[] b)
        {
            return FormatByteArray(b, 0);
        }
    }

#endif
    #endregion




    /// <summary>
    ///   A stream that encrypts as it writes, or decrypts as it reads.  The
    ///   Crypto is AES in CTR (counter) mode, which is compatible with the AES
    ///   encryption employed by WinZip 12.0.
    /// </summary>
    /// <remarks>
    ///   <para>
    ///     The AES/CTR encryption protocol used by WinZip works like this:
    ///
    ///       - start with a counter, initialized to zero.
    ///
    ///       - to encrypt, take the data by 16-byte blocks. For each block:
    ///         - apply the transform to the counter
    ///         - increement the counter
    ///         - XOR the result of the transform with the plaintext to
    ///           get the ciphertext.
    ///         - compute the mac on the encrypted bytes
    ///       - when finished with all blocks, store the computed MAC.
    ///
    ///       - to decrypt, take the data by 16-byte blocks. For each block:
    ///         - compute the mac on the encrypted bytes,
    ///         - apply the transform to the counter
    ///         - increement the counter
    ///         - XOR the result of the transform with the ciphertext to
    ///           get the plaintext.
    ///       - when finished with all blocks, compare the computed MAC against
    ///         the stored MAC
    ///
    ///   </para>
    /// </remarks>
    //
    internal class WinZipAesCipherStream : Stream
    {
        private WinZipAesCrypto _params;
        private System.IO.Stream _s;
        private CryptoMode _mode;
        private int _nonce;
        private bool _finalBlock;

        internal HMACSHA1 _mac;

        // Use RijndaelManaged from .NET 2.0.
        // AesManaged came in .NET 3.5, but we want to limit
        // dependency to .NET 2.0.  AES is just a restricted form
        // of Rijndael (fixed block size of 128, some crypto modes not supported).

        internal RijndaelManaged _aesCipher;
        internal ICryptoTransform _xform;

        private const int BLOCK_SIZE_IN_BYTES = 16;

        private byte[] counter = new byte[BLOCK_SIZE_IN_BYTES];
        private byte[] counterOut = new byte[BLOCK_SIZE_IN_BYTES];

        // I've had a problem when wrapping a WinZipAesCipherStream inside
        // a DeflateStream. Calling Read() on the DeflateStream results in
        // a Read() on the WinZipAesCipherStream, but the buffer is larger
        // than the total size of the encrypted data, and larger than the
        // initial Read() on the DeflateStream!  When the encrypted
        // bytestream is embedded within a larger stream (As in a zip
        // archive), the Read() doesn't fail with EOF.  This causes bad
        // data to be returned, and it messes up the MAC.

        // This field is used to provide a hard-stop to the size of
        // data that can be read from the stream.  In Read(), if the buffer or
        // read request goes beyond the stop, we truncate it.

        private long _length;
        private long _totalBytesXferred;
        private byte[] _PendingWriteBlock;
        private int _pendingCount;
        private byte[] _iobuf;

        /// <summary>
        /// The constructor.
        /// </summary>
        /// <param name="s">The underlying stream</param>
        /// <param name="mode">To either encrypt or decrypt.</param>
        /// <param name="cryptoParams">The pre-initialized WinZipAesCrypto object.</param>
        /// <param name="length">The maximum number of bytes to read from the stream.</param>
        internal WinZipAesCipherStream(System.IO.Stream s, WinZipAesCrypto cryptoParams, long length, CryptoMode mode)
            : this(s, cryptoParams, mode)
        {
            // don't read beyond this limit!
            _length = length;
            //Console.WriteLine("max length of AES stream: {0}", _length);
        }


#if WANT_TRACE
            Stream untransformed;
        String traceFileUntransformed;
        Stream transformed;
        String traceFileTransformed;
#endif


        internal WinZipAesCipherStream(System.IO.Stream s, WinZipAesCrypto cryptoParams, CryptoMode mode)
            : base()
        {
            TraceOutput("-------------------------------------------------------");
            TraceOutput("Create {0:X8}", this.GetHashCode());

            _params = cryptoParams;
            _s = s;
            _mode = mode;
            _nonce = 1;

            if (_params == null)
                throw new BadPasswordException("Supply a password to use AES encryption.");

            int keySizeInBits = _params.KeyBytes.Length * 8;
            if (keySizeInBits != 256 && keySizeInBits != 128 && keySizeInBits != 192)
                throw new ArgumentOutOfRangeException("keysize",
                                                      "size of key must be 128, 192, or 256");

            _mac = new HMACSHA1(_params.MacIv);

            _aesCipher = new System.Security.Cryptography.RijndaelManaged();
            _aesCipher.BlockSize = 128;
            _aesCipher.KeySize = keySizeInBits;  // 128, 192, 256
            _aesCipher.Mode = CipherMode.ECB;
            _aesCipher.Padding = PaddingMode.None;

            byte[] iv = new byte[BLOCK_SIZE_IN_BYTES]; // all zeroes

            // Create an ENCRYPTOR, regardless whether doing decryption or encryption.
            // It is reflexive.
            _xform = _aesCipher.CreateEncryptor(_params.KeyBytes, iv);

            if (_mode == CryptoMode.Encrypt)
            {
                _iobuf = new byte[2048];
                _PendingWriteBlock = new byte[BLOCK_SIZE_IN_BYTES];
            }


#if WANT_TRACE
                traceFileUntransformed = "unpack\\WinZipAesCipherStream.trace.untransformed.out";
            traceFileTransformed = "unpack\\WinZipAesCipherStream.trace.transformed.out";

            untransformed = System.IO.File.Create(traceFileUntransformed);
            transformed = System.IO.File.Create(traceFileTransformed);
#endif
        }

        private void XorInPlace(byte[] buffer, int offset, int count)
        {
            for (int i = 0; i < count; i++)
            {
                buffer[offset + i] = (byte)(counterOut[i] ^ buffer[offset + i]);
            }
        }

        private void WriteTransformOneBlock(byte[] buffer, int offset)
        {
            System.Array.Copy(BitConverter.GetBytes(_nonce++), 0, counter, 0, 4);
            _xform.TransformBlock(counter,
                                  0,
                                  BLOCK_SIZE_IN_BYTES,
                                  counterOut,
                                  0);
            XorInPlace(buffer, offset, BLOCK_SIZE_IN_BYTES);
            _mac.TransformBlock(buffer, offset, BLOCK_SIZE_IN_BYTES, null, 0);
        }


        private void WriteTransformBlocks(byte[] buffer, int offset, int count)
        {
            int posn = offset;
            int last = count + offset;

            while (posn < buffer.Length && posn < last)
            {
                WriteTransformOneBlock (buffer, posn);
                posn += BLOCK_SIZE_IN_BYTES;
            }
        }


        private void WriteTransformFinalBlock()
        {
            if (_pendingCount == 0)
                throw new InvalidOperationException("No bytes available.");

            if (_finalBlock)
                throw new InvalidOperationException("The final block has already been transformed.");

            System.Array.Copy(BitConverter.GetBytes(_nonce++), 0, counter, 0, 4);
            counterOut = _xform.TransformFinalBlock(counter,
                                                    0,
                                                    BLOCK_SIZE_IN_BYTES);
            XorInPlace(_PendingWriteBlock, 0, _pendingCount);
            _mac.TransformFinalBlock(_PendingWriteBlock, 0, _pendingCount);
            _finalBlock = true;
        }





        private int ReadTransformOneBlock(byte[] buffer, int offset, int last)
        {
            if (_finalBlock)
                throw new NotSupportedException();

            int bytesRemaining = last - offset;
            int bytesToRead = (bytesRemaining > BLOCK_SIZE_IN_BYTES)
                ? BLOCK_SIZE_IN_BYTES
                : bytesRemaining;

            // update the counter
            System.Array.Copy(BitConverter.GetBytes(_nonce++), 0, counter, 0, 4);

            // Determine if this is the final block
            if ((bytesToRead == bytesRemaining) &&
                (_length > 0) &&
                (_totalBytesXferred + last == _length))
            {
                _mac.TransformFinalBlock(buffer, offset, bytesToRead);
                counterOut = _xform.TransformFinalBlock(counter,
                                                        0,
                                                        BLOCK_SIZE_IN_BYTES);
                _finalBlock = true;
            }
            else
            {
                _mac.TransformBlock(buffer, offset, bytesToRead, null, 0);
                _xform.TransformBlock(counter,
                                      0, // offset
                                      BLOCK_SIZE_IN_BYTES,
                                      counterOut,
                                      0);  // offset
            }

            XorInPlace(buffer, offset, bytesToRead);
            return bytesToRead;
        }



        private void ReadTransformBlocks(byte[] buffer, int offset, int count)
        {
            int posn = offset;
            int last = count + offset;

            while (posn < buffer.Length && posn < last )
            {
                int n = ReadTransformOneBlock (buffer, posn, last);
                posn += n;
            }
        }



        public override int Read(byte[] buffer, int offset, int count)
        {
            if (_mode == CryptoMode.Encrypt)
                throw new NotSupportedException();

            if (buffer == null)
                throw new ArgumentNullException("buffer");

            if (offset < 0)
                throw new ArgumentOutOfRangeException("offset",
                                                      "Must not be less than zero.");
            if (count < 0)
                throw new ArgumentOutOfRangeException("count",
                                                      "Must not be less than zero.");

            if (buffer.Length < offset + count)
                throw new ArgumentException("The buffer is too small");

            // When I wrap a WinZipAesStream in a DeflateStream, the
            // DeflateStream asks its captive to read 4k blocks, even if the
            // encrypted bytestream is smaller than that.  This is a way to
            // limit the number of bytes read.

            int bytesToRead = count;

            if (_totalBytesXferred >= _length)
            {
                return 0; // EOF
            }

            long bytesRemaining = _length - _totalBytesXferred;
            if (bytesRemaining < count) bytesToRead = (int)bytesRemaining;

            int n = _s.Read(buffer, offset, bytesToRead);


#if WANT_TRACE
                untransformed.Write(buffer, offset, bytesToRead);
#endif

            ReadTransformBlocks(buffer, offset, bytesToRead);

#if WANT_TRACE
                transformed.Write(buffer, offset, bytesToRead);
#endif
            _totalBytesXferred += n;
            return n;
        }



        /// <summary>
        /// Returns the final HMAC-SHA1-80 for the data that was encrypted.
        /// </summary>
        public byte[] FinalAuthentication
        {
            get
            {
                if (!_finalBlock)
                {
                    // special-case zero-byte files
                    if ( _totalBytesXferred != 0)
                        throw new BadStateException("The final hash has not been computed.");

                    // Must call ComputeHash on an empty byte array when no data
                    // has run through the MAC.

                    byte[] b = {  };
                    _mac.ComputeHash(b);
                    // fall through
                }
                byte[] macBytes10 = new byte[10];
                System.Array.Copy(_mac.Hash, 0, macBytes10, 0, 10);
                return macBytes10;
            }
        }


        public override void Write(byte[] buffer, int offset, int count)
        {
            if (_finalBlock)
                throw new InvalidOperationException("The final block has already been transformed.");

            if (_mode == CryptoMode.Decrypt)
                throw new NotSupportedException();

            if (buffer == null)
                throw new ArgumentNullException("buffer");

            if (offset < 0)
                throw new ArgumentOutOfRangeException("offset",
                                                      "Must not be less than zero.");
            if (count < 0)
                throw new ArgumentOutOfRangeException("count",
                                                      "Must not be less than zero.");
            if (buffer.Length < offset + count)
                throw new ArgumentException("The offset and count are too large");

            if (count == 0)
                return;

            TraceOutput("Write off({0}) count({1})", offset, count);

#if WANT_TRACE
            untransformed.Write(buffer, offset, count);
#endif

            // For proper AES encryption, an AES encryptor application calls
            // TransformBlock repeatedly for all 16-byte blocks except the
            // last. For the last block, it then calls TransformFinalBlock().
            //
            // This class is a stream that encrypts via Write().  But, it's not
            // possible to recognize which are the "last" bytes from within the call
            // to Write(). The caller can call Write() several times in succession,
            // with varying buffers. This class only "knows" that the last bytes
            // have been written when the app calls Close().
            //
            // Therefore, this class buffers writes: After completion every Write(),
            // a 16-byte "pending" block (_PendingWriteBlock) must hold between 1
            // and 16 bytes, which will be used in TransformFinalBlock if the app
            // calls Close() immediately thereafter. Also, every write must
            // transform any pending bytes, before transforming the data passed in
            // to the current call.
            //
            // In operation, after the first call to Write() and before the call to
            // Close(), one full or partial block of bytes is always available,
            // pending.  At time of Close(), this class calls
            // WriteTransformFinalBlock() to flush the pending bytes.
            //
            // This approach works whether the caller writes in odd-sized batches,
            // for example 5000 bytes, or in batches that are neat multiples of the
            // blocksize (16).
            //
            // Logicaly, what we do is this:
            //
            //  1. if there are fewer than 16 bytes (pending + current), then
            //     just copy them into th pending buffer and return.
            //
            //  2. there are more than 16 bytes to write. So, take the leading slice
            //     of bytes from the current buffer, enough to fill the pending
            //     buffer. Transform the pending block, and write it out.
            //
            //  3. Take the trailing slice of bytes (a full block or a partial block),
            //     and copy it to the pending block for next time.
            //
            //  4. transform and write all the other blocks, the middle slice.
            //

            // There are 16 or fewer bytes, so just buffer the bytes.
            if (count + _pendingCount <= BLOCK_SIZE_IN_BYTES)
            {
                Buffer.BlockCopy(buffer,
                                 offset,
                                 _PendingWriteBlock,
                                 _pendingCount,
                                 count);
                _pendingCount += count;

                // At this point, _PendingWriteBlock contains up to
                // BLOCK_SIZE_IN_BYTES bytes, and _pendingCount ranges from 0 to
                // BLOCK_SIZE_IN_BYTES. We don't want to xform+write them yet,
                // because this may have been the last block.  The last block gets
                // written at Close().
                return;
            }

            // We know there are at least 17 bytes, counting those in the current
            // buffer, along with the (possibly empty) pending block.

            int bytesRemaining = count;
            int curOffset = offset;

            // workitem 12815
            //
            // xform chunkwise ... Cannot transform in place using the original
            // buffer because that is user-maintained.

            if (_pendingCount != 0)
            {
                // We have more than one block of data to write, therefore it is safe
                // to xform+write.
                int fillCount = BLOCK_SIZE_IN_BYTES - _pendingCount;

                // fillCount is possibly zero here. That happens when the pending
                // buffer held 16 bytes (one complete block) before this call to
                // Write.
                if (fillCount > 0)
                {
                    Buffer.BlockCopy(buffer,
                                     offset,
                                     _PendingWriteBlock,
                                     _pendingCount,
                                     fillCount);

                    // adjust counts:
                    bytesRemaining -= fillCount;
                    curOffset += fillCount;
                }

                // xform and write:
                WriteTransformOneBlock(_PendingWriteBlock, 0);
                _s.Write(_PendingWriteBlock, 0, BLOCK_SIZE_IN_BYTES);
                _totalBytesXferred += BLOCK_SIZE_IN_BYTES;
                _pendingCount = 0;
            }

            // At this point _PendingWriteBlock is empty, and bytesRemaining is
            // always greater than 0.

            // Now, xform N blocks, where N = floor((bytesRemaining-1)/16).  If
            // writing 32 bytes, then xform 1 block, and stage the remaining 16.  If
            // writing 10037 bytes, xform 627 blocks of 16 bytes, then stage the
            // remaining 5 bytes.

            int blocksToXform = (bytesRemaining-1)/BLOCK_SIZE_IN_BYTES;
            _pendingCount = bytesRemaining - (blocksToXform * BLOCK_SIZE_IN_BYTES);

            // _pendingCount is ALWAYS between 1 and 16.
            // Put the last _pendingCount bytes into the pending block.
            Buffer.BlockCopy(buffer,
                             curOffset + bytesRemaining - _pendingCount,
                             _PendingWriteBlock,
                             0,
                             _pendingCount);
            bytesRemaining -= _pendingCount;
            _totalBytesXferred += bytesRemaining; // will be true after the loop

            // now, transform all the full blocks preceding that.
            // bytesRemaining is always a multiple of 16 .
            if (blocksToXform > 0)
            {
                do
                {
                    int c = _iobuf.Length;
                    if (c > bytesRemaining) c = bytesRemaining;
                    Buffer.BlockCopy(buffer,
                                     curOffset,
                                     _iobuf,
                                     0,
                                     c);

                    WriteTransformBlocks(_iobuf, 0, c);
                    _s.Write(_iobuf, 0, c);
                    bytesRemaining -= c;
                    curOffset += c;
                } while(bytesRemaining > 0);
            }
        }



        /// <summary>
        ///   Close the stream.
        /// </summary>
        public override void Close()
        {
            TraceOutput("Close {0:X8}", this.GetHashCode());

            // In the degenerate case, no bytes have been written to the
            // stream at all.  Need to check here, and NOT emit the
            // final block if Write has not been called.
            if (_pendingCount > 0)
            {
                WriteTransformFinalBlock();
                _s.Write(_PendingWriteBlock, 0, _pendingCount);
                _totalBytesXferred += _pendingCount;
                _pendingCount = 0;
            }
            _s.Close();

#if WANT_TRACE
            untransformed.Close();
            transformed.Close();
            Console.WriteLine("\nuntransformed bytestream is in  {0}", traceFileUntransformed);
            Console.WriteLine("\ntransformed bytestream is in  {0}", traceFileTransformed);
#endif
            TraceOutput("-------------------------------------------------------");
        }


        /// <summary>
        /// Returns true if the stream can be read.
        /// </summary>
        public override bool CanRead
        {
            get
            {
                if (_mode != CryptoMode.Decrypt) return false;
                return true;
            }
        }


        /// <summary>
        /// Always returns false.
        /// </summary>
        public override bool CanSeek
        {
            get { return false; }
        }

        /// <summary>
        /// Returns true if the CryptoMode is Encrypt.
        /// </summary>
        public override bool CanWrite
        {
            get { return (_mode == CryptoMode.Encrypt); }
        }

        /// <summary>
        /// Flush the content in the stream.
        /// </summary>
        public override void Flush()
        {
            _s.Flush();
        }

        /// <summary>
        /// Getting this property throws a NotImplementedException.
        /// </summary>
        public override long Length
        {
            get { throw new NotImplementedException(); }
        }

        /// <summary>
        /// Getting or Setting this property throws a NotImplementedException.
        /// </summary>
        public override long Position
        {
            get { throw new NotImplementedException(); }
            set { throw new NotImplementedException(); }
        }

        /// <summary>
        /// This method throws a NotImplementedException.
        /// </summary>
        public override long Seek(long offset, System.IO.SeekOrigin origin)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// This method throws a NotImplementedException.
        /// </summary>
        public override void SetLength(long value)
        {
            throw new NotImplementedException();
        }



        [System.Diagnostics.ConditionalAttribute("Trace")]
        private void TraceOutput(string format, params object[] varParams)
        {
            lock(_outputLock)
            {
                int tid = System.Threading.Thread.CurrentThread.GetHashCode();
                Console.ForegroundColor = (ConsoleColor) (tid % 8 + 8);
                Console.Write("{0:000} WZACS ", tid);
                Console.WriteLine(format, varParams);
                Console.ResetColor();
            }
        }

        private object _outputLock = new Object();
    }
}
#endif
