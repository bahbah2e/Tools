namespace PowerButton
{
    using System.Windows.Input;

    using Microsoft.Test;

    internal class Program
    {
        private static void Main(string[] args)
        {
            Keyboard.Press(Key.LWin);
            Keyboard.Press(Key.I);
            Keyboard.Release(Key.I);
            Keyboard.Release(Key.LWin);
        }
    }
}