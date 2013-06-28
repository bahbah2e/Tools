namespace ScreenCapture
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
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

        public void CaptureElementAutomationId(string elementAutomationId, string fileName)
        {
            if (string.IsNullOrWhiteSpace(this.CaptureFolder))
            {
                this.CaptureFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures);
            }

            CaptureAutomationElementByAutomationId(elementAutomationId, this.CaptureFolder, fileName);
        }

        public void CaptureElementAutomationId(string elementAutomationId, string rootAutomationId, string fileName)
        {
            if (string.IsNullOrWhiteSpace(this.CaptureFolder))
            {
                this.CaptureFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures);
            }

            CaptureAutomationElementByAutomationId(elementAutomationId, rootAutomationId, this.CaptureFolder, fileName);
        }

        private static void CaptureAutomationElementByName(string elementName, string path, string fileName)
        {
            AutomationElement element = FindElement(elementName);
            CaptureAutomationElement(element, path, fileName);
        }

        private static void CaptureAutomationElementByAutomationId(string elementAutomationId, string path, string fileName)
        {
            AutomationElement element = FindElementById(elementAutomationId);
            CaptureAutomationElement(element, path, fileName);
        }

        private static void CaptureAutomationElementByAutomationId(string elementAutomationId, string rootAutomationId, string path, string fileName)
        {
            AutomationElement element = FindElementById(elementAutomationId, rootAutomationId);
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

        private static AutomationElement FindElementById(string elementAutomationId)
        {
            AutomationElement root = AutomationElement.RootElement;

            return FindElementById(elementAutomationId, root);
        }

        private static AutomationElement FindElementById(string elementAutomationId, string rootAutomationId)
        {
            AutomationElement root = FindElementById(rootAutomationId);

            return FindElementById(elementAutomationId, root);
        }

        private static AutomationElement FindElementById(string elementAutomationId, AutomationElement root)
        {
            List<AutomationElement> children = FindElementChildren(root);

            foreach(AutomationElement element in children)
            {
                Trace.WriteLine(element.Current.AutomationId);
                if (element.Current.AutomationId == elementAutomationId ||
                    element.Current.Name == elementAutomationId)
                {
                    return element;
                }
            }

            foreach (AutomationElement element in children)
            {
                return FindElementById(elementAutomationId, element);

            }

            return null;
        }

        private static List<AutomationElement> FindElementChildren(AutomationElement root)
        {
            List<AutomationElement> children = new List<AutomationElement>();

            AutomationElement elementNode = TreeWalker.ControlViewWalker.GetFirstChild(root);

            while (elementNode != null)
            {
                children.Add(elementNode);
                elementNode = TreeWalker.ControlViewWalker.GetNextSibling(elementNode);
            }

            return children;
        }
    }
}