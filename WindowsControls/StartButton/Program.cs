namespace StartButton
{
    using System.Windows.Input;

    using Microsoft.Test;

    internal class Program
    {
        private static void Main(string[] args)
        {
            Keyboard.Press(Key.LWin);
            Keyboard.Release(Key.LWin);
        }
    }
}