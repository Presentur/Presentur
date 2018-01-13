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

        private ShapeService()
        {
            //singleton
        }

        private void PasteToCurrentPresentation()
        {
            Globals.ThisAddIn.Application.ActiveWindow.View.Slide.Shapes.Paste();
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
            var shapeHeight = shape.Shape.Height;
            var shapeWidth = shape.Shape.Width;
            var centerLeft = slideWidth / 2 - (shapeWidth / 2);
            var centerTop = slideHeight / 2 - (shapeHeight / 2);
            shape.Shape.Left = centerLeft;
            shape.Shape.Top = centerTop;
        }
    }
}
