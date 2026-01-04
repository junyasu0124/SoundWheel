using Microsoft.Win32;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace SoundWheel;

internal class Program
{
  private const int WH_MOUSE_LL = 14;
  private const int WM_XBUTTONDOWN = 0x020B;
  private const int WM_XBUTTONUP = 0x020C;
  private const int WM_MBUTTONDOWN = 0x0207; // 中央ボタン
  private const int WM_MBUTTONUP = 0x0208;
  private const int WM_MOUSEWHEEL = 0x020A;

  private const int VK_VOLUME_UP = 0xAF;
  private const int VK_VOLUME_DOWN = 0xAE;
  private const uint KEYEVENTF_KEYDOWN = 0x0000;
  private const uint KEYEVENTF_KEYUP = 0x0002;

  private const float SMALL_STEP = 0.01f;

  private static readonly LowLevelMouseProc proc = HookCallback;
  private static IntPtr hookID = IntPtr.Zero;

  private static CoreAudio? coreAudio = null;

  private static bool isSmallStep = true;
  private static bool isBackButtonHeld = false;
  private static bool suppressBackButton = false;

  static void Main()
  {
    if (!IsStartupEnabled())
      SetStartupEnabled();

    hookID = SetHook(proc);
    try
    {
      Application.Run();
    }
    finally
    {
      ReleaseHook();
      if (coreAudio != null)
      {
        coreAudio.Dispose();
        coreAudio = null;
      }
    }
  }

  private static IntPtr SetHook(LowLevelMouseProc proc)
  {
    using Process curProcess = Process.GetCurrentProcess();
    if (curProcess != null && curProcess.MainModule != null)
    {
      using ProcessModule curModule = curProcess.MainModule;
      return SetWindowsHookEx(WH_MOUSE_LL, proc, GetModuleHandle(curModule.ModuleName), 0);
    }
    else
    {
      throw new InvalidOperationException("Failed to get current process or module.");
    }
  }

  private static void ReleaseHook()
  {
    if (hookID != IntPtr.Zero)
    {
      UnhookWindowsHookEx(hookID);
      hookID = IntPtr.Zero;
    }
  }

  private delegate IntPtr LowLevelMouseProc(int nCode, IntPtr wParam, IntPtr lParam);

  private static IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
  {
    try
    {
      if (nCode >= 0)
      {
        var hookStruct = Marshal.PtrToStructure<MSLLHOOKSTRUCT>(lParam);

        if (hookStruct.dwExtraInfo == IntPtr.Zero)
        {
          if (wParam == WM_XBUTTONDOWN)
          {
            ushort button = (ushort)(hookStruct.mouseData >> 16);
            if (button == 1) // XBUTTON1 (戻る)
            {
              isBackButtonHeld = true;

              return 1;
            }
          }
          else if (wParam == WM_XBUTTONUP)
          {
            ushort button = (ushort)(hookStruct.mouseData >> 16);
            if (button == 1)
            {
              isBackButtonHeld = false;

              if (suppressBackButton)
              {
                suppressBackButton = false;
                if (coreAudio != null)
                {
                  coreAudio.Dispose();
                  coreAudio = null;
                }

                return 1;
              }
            }
          }
          else if (wParam == WM_MBUTTONDOWN)
          {
            if (isBackButtonHeld)
            {
              return 1;
            }
          }
          else if (wParam == WM_MBUTTONUP)
          {
            if (isBackButtonHeld)
            {
              isSmallStep = !isSmallStep;
              suppressBackButton = true;
              return 1;
            }
          }
          else if (wParam == WM_MOUSEWHEEL)
          {
            if (isBackButtonHeld)
            {
              short delta = (short)((hookStruct.mouseData >> 16) & 0xffff);
              if (delta > 0)
              {
                if (isSmallStep)
                {
                  var currentVolume = (coreAudio ??= new CoreAudio()).GetMasterVolume();
                  if (currentVolume.HasValue)
                  {
                    var newVolume = (float)Math.Round(currentVolume.Value + SMALL_STEP, 2);
                    if (newVolume > 0.9975f)
                      newVolume = 1.0f;
                    coreAudio!.SetMasterVolume(newVolume);
                  }
                }
                else
                {
                  SendVolumeKey(VK_VOLUME_UP);
                }
              }
              else if (delta < 0)
              {
                if (isSmallStep)
                {
                  var currentVolume = (coreAudio ??= new CoreAudio()).GetMasterVolume();
                  if (currentVolume.HasValue)
                  {
                    var newVolume = (float)Math.Round(currentVolume.Value - SMALL_STEP, 2);
                    if (newVolume < 0.0025f)
                      newVolume = 0.0f;
                    coreAudio!.SetMasterVolume(newVolume);
                  }
                }
                else
                {
                  SendVolumeKey(VK_VOLUME_DOWN);
                }
              }

              suppressBackButton = true;

              return 1;
            }
          }
        }
      }

      return CallNextHookEx(hookID, nCode, wParam, lParam);
    }
    catch
    {
      if (coreAudio != null)
      {
        coreAudio.Dispose();
        coreAudio = null;
      }
      return CallNextHookEx(hookID, nCode, wParam, lParam);
    }
  }

  private static void SendVolumeKey(byte keyCode)
  {
    keybd_event(keyCode, 0, KEYEVENTF_KEYDOWN, 0);
    keybd_event(keyCode, 0, KEYEVENTF_KEYUP, 0);
  }

  private static void SetStartupEnabled()
  {
    using var key = Registry.CurrentUser.CreateSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", true);
    key.SetValue(Application.ProductName, Application.ExecutablePath);
  }
  private static bool IsStartupEnabled()
  {
    using var key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", true);
    return key?.GetValue(Application.ProductName) != null;
  }

  [StructLayout(LayoutKind.Sequential)]
  private struct POINT
  {
    public int x;
    public int y;
  }

  [StructLayout(LayoutKind.Sequential)]
  private struct MSLLHOOKSTRUCT
  {
    public POINT pt;
    public uint mouseData;
    public uint flags;
    public uint time;
    public IntPtr dwExtraInfo;
  }

  [DllImport("user32.dll")]
  private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, uint dwExtraInfo);

  [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
  private static extern IntPtr SetWindowsHookEx(int idHook,
      LowLevelMouseProc lpfn, IntPtr hMod, uint dwThreadId);

  [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
  private static extern bool UnhookWindowsHookEx(IntPtr hhk);

  [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
  private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

  [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
  private static extern IntPtr GetModuleHandle(string lpModuleName);
}
