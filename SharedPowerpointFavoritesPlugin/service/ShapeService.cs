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
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using SharedPowerpointFavoritesPlugin.model;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;
using Core = Microsoft.Office.Core;

namespace SharedPowerpointFavoritesPlugin
{
    class ShapeService
    {

        public static ShapeService INSTANCE = new ShapeService();
        private ShapePersistence shapePersistance = ShapePersistence.INSTANCE;
        private static readonly DebugLogger logger = DebugLogger.GetLogger(typeof(ShapeService).Name);
        private ShapeService()
        {
            //singleton
        }

        private void PasteToCurrentPresentation()
        {
            Globals.ThisAddIn.Application.ActiveWindow.View.Slide.Shapes.Paste();
        }

        public List<ShapeFavorite> GetShapesByType(Core.MsoShapeType? type)
        {
            var allShapes = this.shapePersistance.GetShapes();
            if (type == null)
            {
                return allShapes;
            }
            return allShapes.FindAll(shapeFavorite => shapeFavorite.Shape.Type == type);
        }

        public void PasteToCurrentPresentation(ShapeFavorite shape)
        {
            SetCenterLocation(shape);
            shape.Shape.Copy();
            this.PasteToCurrentPresentation();
        }

        private void SetCenterLocation(ShapeFavorite shape)
        {
            PowerPoint.Presentation currentPresentation = Globals.ThisAddIn.Application.ActivePresentation;
            var slideHeight = currentPresentation.PageSetup.SlideHeight;
            var slideWidth = currentPresentation.PageSetup.SlideWidth;
            var centerLeft = 0F;
            var centerTop = 0F;
            if (shape.Shape.Type != Core.MsoShapeType.msoTextBox)
            {
                var shapeHeight = shape.Shape.Height;
                var shapeWidth = shape.Shape.Width;
                centerLeft = slideWidth / 2 - (shapeWidth / 2);
                centerTop = slideHeight / 2 - (shapeHeight / 2);
            }
            else
            {
                centerLeft = slideWidth / 2;
                centerTop = slideHeight / 2;
            }
            shape.Shape.Left = centerLeft;
            shape.Shape.Top = centerTop;
        }

        internal void DeleteShape(ShapeFavorite shapeFavorite)
        {
            this.shapePersistance.DeleteShape(shapeFavorite);
        }
    }
}
