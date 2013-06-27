namespace ScreenCapture
{
    using System;
    using System.Drawing;
    using System.IO;
    using System.Windows;
    using System.Windows.Automation;

    using Point = System.Drawing.Point;
    using Size = System.Drawing.Size;

    public class AutomationElementCapture
    {
        public string CaptureFolder { get; set; }

        public void CaptureElementByName(string elementName, string fileName)
        {
            if (string.IsNullOrWhiteSpace(this.CaptureFolder))
            {
                this.CaptureFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures);
            }

            CaptureAutomationElementByName(elementName, this.CaptureFolder, fileName);
        }

        private static void CaptureAutomationElementByName(string elementName, string path, string fileName)
        {
            AutomationElement element = FindElement(elementName);
            CaptureAutomationElement(element, path, fileName);
        }

        private static void CaptureAutomationElement(AutomationElement element, string path, string fileName)
        {
            Rect elementRect = (Rect)element.Current.BoundingRectangle;

            Bitmap dumpBitmap = new Bitmap(
                Convert.ToInt32(elementRect.Width),
                Convert.ToInt32(elementRect.Height));

            using (Graphics targetGraphics = Graphics.FromImage(dumpBitmap))
            {
                Point captureTopLeft = new Point(
                    Convert.ToInt32(elementRect.TopLeft.X),
                    Convert.ToInt32(elementRect.TopLeft.Y));

                Size captureSize = new Size(dumpBitmap.Width, dumpBitmap.Height);
                targetGraphics.CopyFromScreen(captureTopLeft, new Point(0, 0), captureSize);

                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }

                dumpBitmap.Save(string.Format("{0}\\{1}.bmp", path, fileName));
            }
        }

        private static AutomationElement FindElement(string elementName)
        {
            AutomationElement root = AutomationElement.RootElement;

            AutomationElement elementNode = TreeWalker.ControlViewWalker.GetFirstChild(root);

            while (elementNode != null)
            {
                if (elementNode.Current.Name.Contains(elementName))
                {
                    return elementNode;
                }

                elementNode = TreeWalker.ControlViewWalker.GetNextSibling(elementNode);
            }

            return null;
        }
    }
}