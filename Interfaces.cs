using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace DumpThumbnail.Interop
{
    /// <summary>
    /// Exposes a method that initializes a handler, such as a property handler, thumbnail handler, or preview handler, with a stream.
    /// </summary>
    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("b824b49d-22ac-4161-ac8a-9916e8fa3f7f")]
    public interface IInitializeWithStream
    {
        /// <summary>
        /// Initializes a handler with a stream.
        /// </summary>
        /// <param name="pstream">A pointer to an IStream interface that represents the stream source.</param>
        /// <param name="grfMode">One of the following STGM values that indicates the access mode for pstream. STGM_READ or STGM_READWRITE.</param>
        [PreserveSig]
        HResult Initialize(IStream pstream, uint grfMode);
    }

    /// <summary>
    /// Exposes a method for getting a thumbnail image.
    /// </summary>
    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("e357fccd-a995-4576-b01f-234630154e96")]
    public interface IThumbnailProvider
    {
        /// <summary>
        /// Gets a thumbnail image and alpha type.
        /// </summary>
        /// <param name="cx">The maximum thumbnail size, in pixels. The Shell draws the returned bitmap at this size or smaller. The returned bitmap should fit into a square of width and height cx, though it does not need to be a square image. The Shell scales the bitmap to render at lower sizes. For example, if the image has a 6:4 aspect ratio, then the returned bitmap should also have a 6:4 aspect ratio.</param>
        /// <param name="phbmp">When this method returns, contains a pointer to the thumbnail image handle. The image must be a DIB section and 32 bits per pixel. The Shell scales down the bitmap if its width or height is larger than the size specified by cx. The Shell always respects the aspect ratio and never scales a bitmap larger than its original size.</param>
        /// <param name="pdwAlpha">When this method returns, contains a pointer to one of the following values from the WTS_ALPHATYPE enumeration.</param>
        /// <returns>If this method succeeds, it returns S_OK. Otherwise, it returns an HRESULT error code.</returns>
        [PreserveSig]
        HResult GetThumbnail(UInt32 cx, out IntPtr phbmp, out WTS_ALPHATYPE pdwAlpha);
    }

    /// <summary>
    /// Alpha channel type information.
    /// </summary>
    public enum WTS_ALPHATYPE
    {
        /// <summary>
        /// The bitmap is an unknown format. The Shell tries nonetheless to detect whether the image has an alpha channel.
        /// </summary>
        WTSAT_UNKNOWN = 0x0,

        /// <summary>
        /// The bitmap is an RGB image without alpha. The alpha channel is invalid and the Shell ignores it.
        /// </summary>
        WTSAT_RGB = 0x1,

        /// <summary>
        /// The bitmap is an ARGB image with a valid alpha channel.
        /// </summary>
        WTSAT_ARGB = 0x2
    }

    public enum HResult : int
    {
        S_OK = 0,
        S_FALSE = 1,
        E_ABORT = -2147467260,
        E_ACCESSDENIED = -2147024891,
        E_FAIL = -2147467259,
        E_HANDLE = -2147024890,
        E_INVALIDARG = -2147024809,
        E_NOINTERFACE = -2147467262,
        E_NOTIMPL = -2147467263,
        E_OUTOFMEMORY = -2147024882,
        E_POINTER = -2147467261,
        E_UNEXPECTED = -2147418113,
        CO_E_SERVER_EXEC_FAILURE = -2146959355,
        REGDB_E_CLASSNOTREG = -2147221164
    }
}
