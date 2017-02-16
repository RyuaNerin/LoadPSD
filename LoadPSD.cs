// LoadPSD
// Made by RyuaNerin
// https://ryuanerin.kr/post/2016-04-06-loadpsd
//
// *** VERSION
// rev. 1 (2016-04-06)
//
// *** SUPPORT
// Net Framework 2.0 or newer
// (Grayscale, Indexed, RGB, CMYK, MultiChannel, Duotone, Lab) + (8/16/32 bit color depth) + alpha channel
//
// *** LICENSE
// BSD 3-clause "New" or "Revised" License
//
// Copyright (c) 2016, RyuaNerin
// All rights reserved.
//
// Redistribution and use in source and binary forms, with or without
// modification, are permitted provided that the following conditions are met:
//
// * Redistributions of source code must retain the above copyright notice, this
//   list of conditions and the following disclaimer.
//
// * Redistributions in binary form must reproduce the above copyright notice,
//   this list of conditions and the following disclaimer in the documentation
//   and/or other materials provided with the distribution.
//
// * Neither the name of LoadPSD nor the names of its
//   contributors may be used to endorse or promote products derived from
//   this software without specific prior written permission.
//
// THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
// AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
// IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
// DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE
// FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
// DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
// SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
// CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY,
// OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
// OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

// if do not want to use unsafe code, annotate or delete next line.
#define USE_UNSAFE_CODE

using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
#if !USE_UNSAFE_CODE
using System.Runtime.InteropServices;
#endif

namespace RyuaNerin.Drawing
{
    public static class LoadPSD
    {
        private class PSDReader
        {
            public PSDReader(Stream stream)
            {
                this.m_baseStream = stream;
            }

            private Stream m_baseStream;

            public long Position
            {
                get { return this.m_baseStream.Position; }
                set { this.m_baseStream.Position = value; }
            }

            public byte ReadByte()
            {
                return (byte)this.m_baseStream.ReadByte();
            }
            public byte[] ReadBytes(int count)
            {
                var buff = new byte[count];
                if (count == 0)
                    return buff;

                int r = this.m_baseStream.Read(buff, 0, count);
                if (r != count)
                    throw new IndexOutOfRangeException();

                return buff;
            }
            public short ReadInt16()
            {
                return (short)(this.m_baseStream.ReadByte() << 8 |
                               this.m_baseStream.ReadByte());
            }
            public int ReadInt32()
            {
                return this.m_baseStream.ReadByte() << 32 |
                       this.m_baseStream.ReadByte() << 16 |
                       this.m_baseStream.ReadByte() << 8  |
                       this.m_baseStream.ReadByte();
            }

            public static int ToInteger(byte[] array, int index, int length)
            {
                int shift = 8 * (length - 1);
                int value = 0;
                for (int i = 0; i < length; ++i)
                {
                    value |= array[index + i] << shift;
                    shift -= 8;
                }

                return value;
            }

            public static float ToFloat(byte[] array, int index = 0)
            {
                var buff = new byte[4];
                buff[0] = array[index + 3];
                buff[1] = array[index + 2];
                buff[2] = array[index + 1];
                buff[3] = array[index + 0];

                return BitConverter.ToSingle(buff, 0);
            }
        }

        public static Image Load(string path)
        {
            using (var file = File.OpenRead(path))
                return Load(file);
        }
        public static Image Load(Stream stream)
        {
            if (stream.CanSeek)
                stream.Position = 0;

            var reader = new PSDReader(stream);

            PsdInfo info = new PsdInfo();

            ReadHeader(ref info, reader);
            ReadColorData(ref info, reader);
            ReadImageResources(ref info, reader);
            ReadLayerAndMaskInfomation(ref info, reader);
            ReadImageData(ref info, reader);

            return CreateBitmap(ref info);
        }

        private enum ColorMods
        {
            Bitmap = 0,
            Grayscale = 1,
            Indexed = 2,
            RGB = 3,
            CMYK = 4,
            Multichannel = 7,
            Duotone = 8,
            Lab = 9
        }
        public enum CompressionModes
        {
            Raw = 0,
            RLE = 1,
            ZipWithoutPrediction = 2,
            ZipWithPrediction = 3
        }

        private struct PsdInfo
        {
            // File Header Section
            public int Channels;
            public int Height;
            public int Width;
            public int ColorDepth;
            public ColorMods ColorMode;

            // Color Mode Data Section
            public byte[] ColorData;

            // Image Resources Section
            public int DpiX;
            public int DpiY;
            public short TransparencyIndex;

            // Image Data
            public CompressionModes CompressionMode;
            public byte[][] ImageData; // [Channels][Data]

            public int Pixcels;
            public int ColorDepthBytes;
            public int SizePerChannel;
            public int Scan0Stride;
            public int Scan0bpp;
        }
        
        /* File Header Section
         * Len  Description         Value
         *  4   Signature           "8BPS"
         *  2   Version             1
         *  6   Reserved            not used.
         *  2   Channels            1 ~ 56
         *  4   Height              1 ~ 30,000
         *  4   Width               1 ~ 30,000
         *  2   Color Depth         1, 8, 16, 32
         *  2   Color Mode          1 ~ 9
        */
        private static void ReadHeader(ref PsdInfo info, PSDReader reader)
        {
            // Signature
            if (reader.ReadByte() != 0x38 ||
                reader.ReadByte() != 0x42 ||
                reader.ReadByte() != 0x50 ||
                reader.ReadByte() != 0x53)
                throw new NotSupportedException();

            // Version
            var version = reader.ReadInt16();
            if (version != 1)
                throw new NotSupportedException();

            // Reserved
            reader.Position += 6;

            // Channels
            info.Channels = reader.ReadInt16();

            // Height
            info.Height = reader.ReadInt32();

            // Width
            info.Width = reader.ReadInt32();

            // Color Depth
            info.ColorDepth = reader.ReadInt16();
            info.ColorDepthBytes = info.ColorDepth / 8;
            if (info.ColorDepthBytes != 1 &&
                info.ColorDepthBytes != 2 &&
                info.ColorDepthBytes != 4)
                throw new NotSupportedException();

            // Color Mode
            info.ColorMode = (ColorMods)reader.ReadInt16();
            if (info.ColorMode == ColorMods.Bitmap)
                throw new NotSupportedException();

            info.Pixcels = info.Width * info.Height;
            info.SizePerChannel = info.Pixcels * info.ColorDepthBytes;
        }

        /* Color Mode Data Section
         * Len  Description         Value
         *  4   Length
         *  n   Color Data
        */
        private static void ReadColorData(ref PsdInfo info, PSDReader reader)
        {
            var len = reader.ReadInt32();
            info.ColorData = reader.ReadBytes(len);
        }
        
        /* Image Resources Section
         * Len  Description         Value
         *  4   Length
         *  n   Resource Block Data
        */
        /* Image Resource Blocks
         * Len  Description         Value
         *  4   Signature           '8BIM'
         *  2   Resource Id
         *  n   Name
         *  4   Resource Data Length
         *  n   Resource Data
        */
        private static void ReadImageResources(ref PsdInfo info, PSDReader reader)
        {
            var length = reader.ReadInt32();

            short id;
            int len;

            info.TransparencyIndex = -1;

            long pos = reader.Position + length;
            while (reader.Position < pos)
            {
                // Signature
                if (reader.ReadByte() != 0x38 ||
                    reader.ReadByte() != 0x42 ||
                    reader.ReadByte() != 0x49 ||
                    reader.ReadByte() != 0x4D)
                    throw new NotSupportedException();

                // Resource Id
                id = reader.ReadInt16();

                // Name
                len = reader.ReadByte();
                if (len > 0)
                {
                    if((len % 2) != 0)
                        len = reader.ReadByte();
                    reader.Position += len;
                }
                reader.Position += 1;
                
                // Data Length
                len = reader.ReadInt32();
                if (len % 2 != 0)
                    len++;

                switch (id)
                {
                case 1005: // ResolutionInfo structure
                    info.DpiX = reader.ReadInt16();
                    reader.Position += 6;
                    info.DpiY = reader.ReadInt16();
                    reader.Position += 6;
                    break;

                case 1047: // Transparency Index. 2 bytes for the index of transparent color, if any.
                    info.TransparencyIndex = reader.ReadInt16();
                    break;

                default:
                    reader.Position += len;
                    break;
                }
            }
        }
        
        /* Layer and Mask Information Section
         * Len  Description
         *  4   Section Length
         *  v   Layer info
         *  v   Global layer mask info
         *  v   Additional Layer Information
        */
        private static void ReadLayerAndMaskInfomation(ref PsdInfo info, PSDReader reader)
        {
            var length = reader.ReadInt32();
            reader.Position += length;
        }

        /* Image Data Section
         * Len  Description         Value
         *  2   Compression Mode    CompressionModes
         *  n   Image Data
         */
        private static void ReadImageData(ref PsdInfo info, PSDReader reader)
        {
            info.CompressionMode = (CompressionModes)reader.ReadInt16();
            int i;

            info.ImageData = new byte[info.Channels][];

            switch (info.CompressionMode)
            {
            case CompressionModes.Raw:
                for (i = 0; i < info.Channels; ++i)
                    info.ImageData[i] = reader.ReadBytes(info.SizePerChannel);
                break;

            case CompressionModes.RLE:
                reader.Position += info.Height * info.Channels * 2;
                for (i = 0; i < info.Channels; ++i)
                {
                    info.ImageData[i] = new byte[info.SizePerChannel];
                    RLEDecompress(reader, info.ImageData[i], info.SizePerChannel);
                }
                break;

            case CompressionModes.ZipWithoutPrediction:
            case CompressionModes.ZipWithPrediction:
                throw new NotSupportedException();
            }
        }
        // By RyuaNerin
        private static void RLEDecompress(PSDReader reader, byte[] output, int count)
        {
            // https://en.wikipedia.org/wiki/PackBits
            int ind = 0;
            int len;
            byte val;

            while (ind < count)
            {
                len = reader.ReadByte();

                // (1 + n) literal bytes of data
                // bin       sb   ub
                // 0000 0000 0    128
                // 0111 1111 127  127
                if (len < 128) // 0 ~ 127
                {
                    while (len-- >= 0)
                        output[ind++] = reader.ReadByte();
                }
                // One byte of data, repeated (1 – n) times in the decompressed output
                // bin       sb   ub    not         +2           1-n
                // 1111 1111 -1   256   0000 0000   0000 0010    2
                // 1000 0001 -127 129   0111 1110   1000 0000    128
                else if (128 < len)
                {
                    val = reader.ReadByte();
                    len = (len ^ 0xFF) + 2;
                    while (len-- > 0)
                        output[ind++] = val;
                }
            }
        }
        
        private static Image CreateBitmap(ref PsdInfo info)
        {
            BitmapData lockBits = null;
            Bitmap bitmap = null;

            var format = GetPixelFormat(ref info);

            try
            {
                bitmap = new Bitmap(info.Width, info.Height, format);
                bitmap.SetResolution(info.DpiX, info.DpiY);

                try
                {
                    lockBits = bitmap.LockBits(new Rectangle(0, 0, info.Width, info.Height), ImageLockMode.WriteOnly, format);
                    info.Scan0Stride = lockBits.Stride;

#if USE_UNSAFE_CODE
                    unsafe
                    {
                        CreateBitmap(ref info, (byte*)lockBits.Scan0);
                    }
#else
                    CreateBitmap(ref info, lockBits.Scan0);
#endif
                }
                catch (Exception e)
                {
                    throw e;
                }
                finally
                {
                    bitmap.UnlockBits(lockBits);
                }
            }
            catch (Exception e)
            {
                if (bitmap != null)
                    bitmap.Dispose();

                throw e;
            }

            return bitmap;
        }
        private static PixelFormat GetPixelFormat(ref PsdInfo info)
        {
            bool containsAlpha = false;

            switch (info.ColorMode)
            {
            case ColorMods.Bitmap:
                throw new NotSupportedException();

            case ColorMods.Grayscale:
            case ColorMods.Duotone:
                containsAlpha = info.Channels >= 2;
                break;

            case ColorMods.Indexed:
                containsAlpha = info.TransparencyIndex != -1;
                break;

            case ColorMods.RGB:
            case ColorMods.Lab:
                containsAlpha = info.Channels >= 4;
                break;

            case ColorMods.CMYK:
            case ColorMods.Multichannel:
                containsAlpha = info.Channels >= 5;
                break;
            }

            info.Scan0bpp = containsAlpha ? 4 : 3;
            return containsAlpha ? PixelFormat.Format32bppArgb : PixelFormat.Format24bppRgb;
        }
#if USE_UNSAFE_CODE
        private static unsafe void CreateBitmap(ref PsdInfo info, byte* scan0)
#else
        private static void CreateBitmap(ref PsdInfo info, IntPtr scan0)
#endif
        {
            int i0, i1, i2, i3, alpha = 255;
            double d0, d1, d2, d3;

            var buffer = new byte[64];
            int pos;

            switch (info.ColorMode)
            {
            case ColorMods.Bitmap:
                throw new NotSupportedException();

            case ColorMods.Grayscale:
            case ColorMods.Duotone:
                #region
                for (pos = 0; pos < info.Pixcels; ++pos)
                {
                    i0 = ReadColor(ref info, 0, pos);

                    if (info.Channels >= 2)
                        alpha = ReadColor(ref info, 1, pos);

                    SetRGB(ref info, scan0, pos, alpha, i0);
                }
                break;
                #endregion

            case ColorMods.Indexed:
                #region
                {
                    int tr, tg, tb;
                    if (info.TransparencyIndex != -1)
                    {
                        tr = info.ColorData[info.TransparencyIndex];
                        tg = info.ColorData[info.TransparencyIndex + 256];
                        tb = info.ColorData[info.TransparencyIndex + 512];
                    }
                    else
                    {
                        tr = tg = tb = -1;
                    }

                    for (pos = 0; pos < info.Pixcels; ++pos)
                    {
                        i0 = info.ImageData[0][pos];

                        i1 = info.ColorData[i0];
                        i2 = info.ColorData[i0 + 256];
                        i3 = info.ColorData[i0 + 512];

                        if (i1 == tr && i2 == tg && i3 == tb)
                            alpha = 0;
                        else
                            alpha = 255;

                        SetRGB(ref info, scan0, pos, alpha, i1, i2, i3);
                    }
                }
                break;
                #endregion

            case ColorMods.RGB:
                #region
                for (pos = 0; pos < info.Pixcels; ++pos)
                {
                    i0 = ReadColor(ref info, 0, pos);
                    i1 = ReadColor(ref info, 1, pos);
                    i2 = ReadColor(ref info, 2, pos);

                    if (info.Channels >= 4)
                        alpha = ReadColor(ref info, 3, pos);

                    SetRGB(ref info, scan0, pos, alpha, i0, i1, i2);
                }
                break;
                #endregion

            case ColorMods.CMYK:
                #region
                for (pos = 0; pos < info.Pixcels; ++pos)
                {
                    d0 = (1.0 - ReadColor(ref info, 0, pos) / 255d);
                    d1 = (1.0 - ReadColor(ref info, 1, pos) / 255d);
                    d2 = (1.0 - ReadColor(ref info, 2, pos) / 255d);
                    d3 = (1.0 - ReadColor(ref info, 3, pos) / 255d);

                    if (info.Channels >= 5)
                        alpha = ReadColor(ref info, 4, pos);

                    SetCMYK(ref info, scan0, pos, alpha, d0, d1, d2, d3);
                }
                break;
                #endregion

            case ColorMods.Multichannel:
                #region
                d3 = 0;
                for (pos = 0; pos < info.Pixcels; ++pos)
                {
                    d0 = (1.0 - ReadColor(ref info, 0, pos) / 255d);
                    d1 = (1.0 - ReadColor(ref info, 1, pos) / 255d);
                    d2 = (1.0 - ReadColor(ref info, 2, pos) / 255d);

                    if (info.Channels >= 4)
                        d3 = (1.0 - ReadColor(ref info, 3, pos) / 255d);

                    if (info.Channels >= 5)
                        alpha = ReadColor(ref info, 4, pos);

                    SetCMYK(ref info, scan0, pos, alpha, d0, d1, d2, d3);
                }
                break;
                #endregion

            case ColorMods.Lab:
                #region
                for (pos = 0; pos < info.Pixcels; ++pos)
                {
                    d0 = ReadColor(ref info, 0, pos) / 255d * 100d; // L : 0 ~ 100
                    d1 = ReadColor(ref info, 1, pos) - 128.0;       // a : -128 ~ 127
                    d2 = ReadColor(ref info, 2, pos) - 128.0;       // b : -128 ~ 127

                    if (info.Channels >= 4)
                        alpha = ReadColor(ref info, 3, pos);

                    SetLab(ref info, scan0, pos, alpha, d0, d1, d2);
                }
                break;
                #endregion
            }
        }

        private static int ReadColor(ref PsdInfo info, int c, int pos)
        {
            int v;

            pos *= info.ColorDepthBytes;

            if (info.ColorDepthBytes == 1)
                v = PSDReader.ToInteger(info.ImageData[c], pos, info.ColorDepthBytes);

            else if (info.ColorDepthBytes == 2)
                v = (int)(PSDReader.ToInteger(info.ImageData[c], pos, info.ColorDepthBytes) / 257d); // (65536 - 1) * (256 - 1) = 257

            else
                v = (int)(255 * Math.Pow(PSDReader.ToFloat(info.ImageData[c], pos), 0.45470693));

            return v;
        }

#if USE_UNSAFE_CODE
        private static unsafe void SetRGB(ref PsdInfo info, byte*  scan0, int pos, int a, int r, int g, int b)
#else
        private static        void SetRGB(ref PsdInfo info, IntPtr scan0, int pos, int a, int r, int g, int b)
#endif
        {
            if (a < 0) a = 0; else if (255 < a) a = 255;
            if (r < 0) r = 0; else if (255 < r) r = 255;
            if (g < 0) g = 0; else if (255 < g) g = 255;
            if (b < 0) b = 0; else if (255 < b) b = 255;

            if (info.Scan0bpp == 4)
            {
                pos = pos * 4;
#if USE_UNSAFE_CODE
                scan0[pos + 0] = (byte)b;
                scan0[pos + 1] = (byte)g;
                scan0[pos + 2] = (byte)r;
                scan0[pos + 3] = (byte)a;
#else
                Marshal.WriteByte(scan0 + pos + 0, (byte)b);
                Marshal.WriteByte(scan0 + pos + 1, (byte)g);
                Marshal.WriteByte(scan0 + pos + 2, (byte)r);
                Marshal.WriteByte(scan0 + pos + 3, (byte)a);
#endif
            }
            else
            {
                pos = (pos / info.Width) * info.Scan0Stride + (pos % info.Width) * 3;

#if USE_UNSAFE_CODE
                scan0[pos + 0] = (byte)b;
                scan0[pos + 1] = (byte)g;
                scan0[pos + 2] = (byte)r;
#else
                Marshal.WriteByte(scan0 + pos + 0, (byte)b);
                Marshal.WriteByte(scan0 + pos + 1, (byte)g);
                Marshal.WriteByte(scan0 + pos + 2, (byte)r);
#endif
            }
        }

#if USE_UNSAFE_CODE
        private static unsafe void SetRGB(ref PsdInfo info, byte*  scan0, int pos, int a, int g)
#else
        private static        void SetRGB(ref PsdInfo info, IntPtr scan0, int pos, int a, int g)
#endif
        {
            SetRGB(ref info, scan0, pos, a, g, g, g);
        }
        
#if USE_UNSAFE_CODE
        private static unsafe void SetLab(ref PsdInfo info, byte*  scan0, int pos, int alpha, double l, double a, double b)
#else
        private static        void SetLab(ref PsdInfo info, IntPtr scan0, int pos, int alpha, double l, double a, double b)
#endif
        {
            // https://en.wikipedia.org/wiki/Lab_color_space#Reverse_transformation
            
            // l a b
            // L a b (Lab)
            // y x z (XYZ)
            l = (l + 16) / 116.0;
            a = l + a / 500.0;
            b = l - b / 200.0;

            // t = theta = 6 / 29 = 0.2068966
            // 3 * t * t = 0.1284186
            // 3 * t * t * 4 / 29 = 0.0177129
            if (a > 0.2068966) a = a * a * a; else a = 0.1284186 * a - 0.0177129; //3tt(y - 4/29);
            if (l > 0.2068966) l = l * l * l; else l = 0.1284186 * l - 0.0177129;
            if (b > 0.2068966) b = b * b * b; else b = 0.1284186 * b - 0.0177129;

            // 6,500 K (D65)
            // https://en.wikipedia.org/wiki/Illuminant_D65
            // Normalizing for relative luminance, the XYZ tristimulus values are X=95.047, Y=100.00, Z=108.883
            a = 0.95047 * a;
            l = 1.00000 * l;
            b = 1.08883 * b;
            
            // CIE XYZ to sRGB
            // https://en.wikipedia.org/wiki/SRGB#The_forward_transformation_.28CIE_xyY_or_CIE_XYZ_to_sRGB.29
            // https://www.w3.org/Graphics/Color/sRGB
            double rr = a *  3.2404542 + l * -1.5371385 + b * -0.4985314;
            double rg = a * -0.9692660 + l *  1.8760108 + b * -0.0415560;
            double rb = a *  0.0556434 + l * -0.2040259 + b *  1.0572252;

            // a = 0.055;
            // 1 / 2.4 = 0.41666667
            if (rr <= 0.0031308) rr = 12.92 * rr; else rr = 1.055 * (Math.Pow(rr, 0.41666667)) - 0.055;
            if (rg <= 0.0031308) rg = 12.92 * rg; else rg = 1.055 * (Math.Pow(rg, 0.41666667)) - 0.055;
            if (rb <= 0.0031308) rb = 12.92 * rb; else rb = 1.055 * (Math.Pow(rb, 0.41666667)) - 0.055;
            
            rr *= 256;
            rg *= 256;
            rb *= 256;

            SetRGB(ref info, scan0, pos, alpha, (int)rr, (int)rg, (int)rb);
        }
        
#if USE_UNSAFE_CODE
        private static unsafe void SetCMYK(ref PsdInfo info, byte*  scan0, int pos, int alpha, double C, double M, double Y, double K)
#else
        private static        void SetCMYK(ref PsdInfo info, IntPtr scan0, int pos, int alpha, double C, double M, double Y, double K)
#endif
        {
            int r = (int)(255 * (1 - C) * (1 - K));
            int g = (int)(255 * (1 - M) * (1 - K));
            int b = (int)(255 * (1 - Y) * (1 - K));

            SetRGB(ref info, scan0, pos, alpha, r, g, b);
        }
    }
}
