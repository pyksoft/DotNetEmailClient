using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing.Imaging;
using System.Linq;
using System.Net.Mail;
using System.Net.Mime;
using System.Text;

namespace DM.Core.Net
{
    public class EmailImage : IDisposable
    {
        public enum MediaTypeName
        {
            Jpeg, Gif, Tiff
        }

        LinkedResource _LinkedResource;
        [Browsable(false)]
        public LinkedResource LinkedResource
        {
            get { return _LinkedResource; }
        }

        string _Filename = "";
        [Browsable(false)]
        public string Filename
        {
            get { return _Filename; }
            set { _Filename = value; }
        }

        ImageFormat _MediaType = ImageFormat.Jpeg;
        [Browsable(false)]
        public ImageFormat MediaType
        {
            get { return _MediaType; }
            set { _MediaType = value; }
        }

        string _ContentID = "";
        [Browsable(false)]
        public string ContentID
        {
            get { return _ContentID; }
            set { _ContentID = value; }
        }


        public EmailImage(string filePath, ImageFormat mediaType, string contentID)
        {
            Filename = filePath;
            MediaType = mediaType;
            ContentID = contentID;

            _LinkedResource = new LinkedResource(filePath, string.Concat("image/", mediaType.ToString().ToLower()));
            _LinkedResource.ContentId = contentID;
            _LinkedResource.TransferEncoding = TransferEncoding.Base64;

        }


        public void Dispose()
        {
            _LinkedResource.Dispose();
            _LinkedResource = null;
        }
    }
}
