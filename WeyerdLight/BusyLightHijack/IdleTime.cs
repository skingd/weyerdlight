using System;
using System.Runtime.InteropServices;

namespace BusyLightHijack {	

	/// <summary>
	/// This class was borrowed from the interwebz. Not currently used,
	/// as user idle status is pulled from skype.
	/// </summary>
	public class IdleTime {
		public static class IdleTimeDetector {
			[DllImport("user32.dll")]
			static extern bool GetLastInputInfo(ref LASTINPUTINFO plii);

			public static IdleTimeInfo GetIdleTimeInfo() {
				int systemUptime = Environment.TickCount,
					lastInputTicks = 0,
					idleTicks = 0;

				LASTINPUTINFO lastInputInfo = new LASTINPUTINFO();
				lastInputInfo.cbSize = (uint)Marshal.SizeOf(lastInputInfo);
				lastInputInfo.dwTime = 0;

				if (GetLastInputInfo(ref lastInputInfo)) {
					lastInputTicks = (int)lastInputInfo.dwTime;

					idleTicks = systemUptime - lastInputTicks;
				}

				return new IdleTimeInfo {
					LastInputTime = DateTime.Now.AddMilliseconds(-1 * idleTicks),
					IdleTime = new TimeSpan(0, 0, 0, 0, idleTicks),
					SystemUptimeMilliseconds = systemUptime,
				};
			}
		}

		public class IdleTimeInfo {
			public DateTime LastInputTime { get; internal set; }

			public TimeSpan IdleTime { get; internal set; }

			public int SystemUptimeMilliseconds { get; internal set; }
		}

		internal struct LASTINPUTINFO {
			public uint cbSize;
			public uint dwTime;
		}
	}
}
