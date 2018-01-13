using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.PowerPoint;
using System.Drawing;

namespace SharedPowerpointFavoritesPlugin.model
{
    class ShapeFavorite
    {

        private Image cachedImage;

        public string FilePath
        {
            get;
        }

        public Shape Shape
        {
            get;
        }

        internal Image Thumbnail
        {
            get
            {
                if (this.cachedImage == null)
                {
                    var thumbnail = ShapePersistence.INSTANCE.GetThumbnail(this);
                    using (Image loadedImage = Image.FromFile(thumbnail, true))
                    {
                        var bmp = new Bitmap(loadedImage);
                        this.cachedImage = bmp;
                    }
                }
                return this.cachedImage;
            }
        }

        public ShapeFavorite(string filePath, Shape shape)
        {
            this.FilePath = filePath;
            this.Shape = shape;
        }

        public ShapeFavorite(string filePath, Shape shape, Image thumbnail) : this(filePath, shape)
        {
            if (thumbnail == null)
            {
                throw new ArgumentNullException();
            }
            this.cachedImage = thumbnail;
        }
    }
}
