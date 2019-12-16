using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace MisClasesFuncionesC
{
    public class MiComun
    {
        public enum MessageBeepType
        {
            Default = -1,
            Ok = 0x00000000,
            Error = 0x00000010,
            Question = 0x00000020,
            Warning = 0x00000030,
            Information = 0x00000040
        }

        [DllImport("user32.dll", SetLastError = true, EntryPoint = "MessageBeep")]
        public static extern bool MessageBeep(MessageBeepType type);

        [DllImport("kernel32.dll")]
        public static extern bool Beep(int frequency, int duration);
    }
}
