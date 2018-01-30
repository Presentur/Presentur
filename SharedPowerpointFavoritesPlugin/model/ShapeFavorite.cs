/*
Presentur - Creating presentations in corporate design easily using your set of designed elements, icons and shapes.
Copyright (C) 2018 Christopher Rudoll, Eduard Hajek

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program.  If not, see <https://www.gnu.org/licenses/>.
*/

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
