namespace ScreenCapture.Test
{
    using System.Diagnostics;
    using System.Threading;

    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    public class TestNotepadCapture
    {
        [TestMethod]
        public void test_capture_notepad()
        {
            Process.Start("notepad.exe");
            
            // Sleep long enough to allow the application to start
            Thread.Sleep(2000);

            AutomationElementCapture capture = new AutomationElementCapture();
            capture.CaptureElementByName("Untitled - Notepad", "notepad.exe");
        }
    }
}