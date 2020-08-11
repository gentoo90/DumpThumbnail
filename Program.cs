using DumpThumbnail.Interop;
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;

namespace DumpThumbnail
{
    class Program
    {
        static int Main(string[] args)
        {
            if (args.Length < 4)
            {
                Console.WriteLine($"Usage:\n\t{System.AppDomain.CurrentDomain.FriendlyName} <CLSID> <input file> <output file> <size>");
                return 1;
            }

            var guid = new Guid(args[0]);
            var input = Path.GetFullPath(args[1]);
            var output = Path.GetFullPath(args[2]);
            var size = uint.Parse(args[3]);
            using (Bitmap bmp = CallThumbnailer(guid, input, size))
            {
                bmp.Save(output, ImageFormat.Png);
            }

            return 0;
        }

        public static Bitmap CallThumbnailer(Guid guid, string filePath, uint size)
        {
            using (FileStream fs = File.OpenRead(filePath))
            {
                using (var istm = new ManagedIStream(fs))
                {
                    var thumbType = Type.GetTypeFromCLSID(guid);
                    object thumbHandler = Activator.CreateInstance(thumbType);
                    if (thumbHandler is IInitializeWithStream)
                    {
                        var init_res = ((IInitializeWithStream)thumbHandler).Initialize(istm, 0);
                        if (thumbHandler is IThumbnailProvider)
                        {
                            IntPtr pHbitmap;
                            WTS_ALPHATYPE alphaType;
                            var thumb_res = ((IThumbnailProvider)thumbHandler).GetThumbnail(size, out pHbitmap, out alphaType);

                            var img = Image.FromHbitmap(pHbitmap);
                            return img;
                        }
                        else
                        {
                            throw new NotSupportedException();
                        }
                    }
                    else
                    {
                        throw new NotSupportedException();
                    }
                }
            }
        }
    }
}
