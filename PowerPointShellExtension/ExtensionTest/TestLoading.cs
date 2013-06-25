namespace ExtensionTest
{
    using System;

    using ContextMenuExtension;

    using Microsoft.VisualStudio.TestTools.UnitTesting;

    using SampleExtension;

    [TestClass]
    public class TestLoading
    {
        [TestMethod]
        public void TestActivation()
        {
            var extension = Activator.CreateInstance(typeof(PstContextMenu)) as ContextMenuExt;

            Assert.AreEqual(extension.RegistryFriendlyName, "PST Destroyer");
        }
    }
}
