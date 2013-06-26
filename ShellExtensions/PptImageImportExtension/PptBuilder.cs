namespace PptImageImportExtension
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;

    using Microsoft.Office.Core;
    using Microsoft.Office.Interop.PowerPoint;

    internal class PptBuilder
    {
        private readonly Application ppt = new Application();

        private readonly Presentation presentation;

        private readonly Slides slides;

        public PptBuilder()
        {
            Presentations presentations = this.ppt.Presentations;
            this.presentation = presentations.Add();
            this.presentation.PageSetup.SlideSize = PpSlideSizeType.ppSlideSizeOnScreen16x10;
            this.slides = this.presentation.Slides;
        }

        public void AddToPpt(FileInfo[] unsortedPngFiles)
        {
            int pageNumber = 0;
            FileInfo[] sortedPngFiles = this.Sort(unsortedPngFiles);

            foreach (FileInfo capture in sortedPngFiles)
            {
                Console.WriteLine("Adding {0}", capture.Name);
                _Slide step = this.slides.Add(++pageNumber, PpSlideLayout.ppLayoutBlank);

                step.Shapes.AddPicture(
                    string.Format(capture.FullName),
                    MsoTriState.msoFalse,
                    MsoTriState.msoTrue,
                    1.4454f,
                    1.0f,
                    717f,
                    448f);
            }

            if (sortedPngFiles.Length > 0)
            {
                string pngOne = sortedPngFiles[0].FullName.Replace("png", "ppt");
                this.presentation.SaveAs(pngOne);
            }
        }

        public FileInfo[] Sort(string directory)
        {
            var localDirectory = new DirectoryInfo(directory);
            FileInfo[] pngFiles = localDirectory.GetFiles("*.png");
            var sortedFiles = new Dictionary<int, FileInfo>();

            foreach (FileInfo capture in localDirectory.GetFiles("*.png"))
            {
                int index = this.GetFileIndex(capture.Name);
                sortedFiles.Add(index, capture);
            }

            IEnumerable<FileInfo> sortedPngFiles = sortedFiles.OrderBy(kv => kv.Key).Select(kv => kv.Value);
            return sortedPngFiles as FileInfo[];
        }

        public FileInfo[] Sort(FileInfo[] unsorted)
        {
            var sortedFilesHash = new Dictionary<int, FileInfo>();

            foreach (FileInfo capture in unsorted)
            {
                int index = this.GetFileIndex(capture.Name);
                sortedFilesHash.Add(index, capture);
            }
            return sortedFilesHash.OrderBy(kv => kv.Key).Select(kv => kv.Value).ToArray();
        }

        private int GetFileIndex(string fileName)
        {
            int indexSeparator = fileName.IndexOf('.');
            return Int32.Parse(fileName.Substring(0, indexSeparator));
        }
    }
}