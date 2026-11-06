using System.Runtime.InteropServices;

namespace WinFormsApp1
{
    public partial class Form1 : Form
    {
        [DllImport(@"C:\Users\T00235\source\repos\DLL\ASAudioDLL\ASAudioDLL\x64\Debug\ASAudioDLL.dll", EntryPoint = "MacroTest", CallingConvention = CallingConvention.Cdecl)]
        [return: MarshalAs(UnmanagedType.BStr)]
        public static extern string MacroTest(IntPtr ownerHwnd, string filePath);
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            textBox1.Text = MacroTest(this.Handle, "C:\\Users\\T00235\\source\\repos\\DLL\\ASAudioDLL\\ASAudioDLL\\audio.ini");
        }
    }
}
