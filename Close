using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using System.Management;

class Program { 
[DllImport("winmm.dll")] protected static extern int mciSendString(string Cmd, StringBuilder StrReturn, int ReturnLength, IntPtr HwndCallback);

  static void Main(string[] args) {

  ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT MediaLoaded FROM Win32_CDROMDrive");
  ManagementObjectCollection moc = searcher.Get();
  var enumerator = moc.GetEnumerator();
  if (!enumerator.MoveNext()) throw new Exception("No elements");
  ManagementObject obj = (ManagementObject) enumerator.Current;
  bool status = (bool) obj["MediaLoaded"];

  if (!status) {
    MessageBox.Show("The drive is either open or empty", "Optical Drive Status");
    mciSendString("set cdaudio door closed", null, 0, IntPtr.Zero);
  }
  else MessageBox.Show("The drive is closed and contains an optical media", "Optical Drive Status");
 }
}
