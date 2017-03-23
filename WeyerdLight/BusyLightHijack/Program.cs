
using Busylight;
using System.Threading;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Lync;
using Microsoft.Lync.Model;

namespace BusyLightHijack {

	/// <summary>
	/// This program controls Kuando busy lights. 
	/// <author>Stephen King
	/// <version>3.21.2017.01
	/// <comments> I wrote this program during down periods at work. I'm not certain
	/// how long or how well I will maintain this. It can detect when the system leaves
	/// hybernate and can launch on startup if placed in shell:startup folder.
	/// 
	/// <functionality> Currently just uninstall the BusyLight app and put this in startup.
	/// <limitations>The busylight DLL only works with Skype for Business.
	/// <TODO>
	///	Detect missed call 
	///	Custom color window
	/// Dock to systray
	/// Exit point
	/// More robust holiday system (detect when holiday is on a weekend)
	/// Holiday toggle
	/// </summary>
	class MainClass {
		static LyncClient skypeClient = null;

		[DllImport("kernel32.dll")]
		static extern IntPtr GetConsoleWindow();

		[DllImport("user32.dll")]
		static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

		const int SW_HIDE = 0;
		const int SW_SHOW = 5;



		//instantiate SDK
		public static SDK busylight = new Busylight.SDK();
		public static BusylightColor turquoise = new BusylightColor() { RedRgbValue = 0, GreenRgbValue = 206, BlueRgbValue = 209 };
		public static BusylightColor white = new BusylightColor() { RedRgbValue = 255, GreenRgbValue = 255, BlueRgbValue = 255 };
		public static BusylightColor deepRed = new BusylightColor() { RedRgbValue = 153, GreenRgbValue = 0, BlueRgbValue = 0 };
		public static BusylightColor violet = new BusylightColor() { RedRgbValue = 255, GreenRgbValue = 0, BlueRgbValue = 255 };
		public static BusylightColor orange = new BusylightColor() { RedRgbValue = 255, GreenRgbValue = 30, BlueRgbValue = 0 };

		/// <summary>
		/// Main method
		/// </summary>
		/// <param name="args">The command-line arguments.</param>
		public static void Main(string[] args) {

			Console.WriteLine("Starting Program");

			//Check status every half second
			while (true) {
					
					if (skypeClient == null) {
						Console.WriteLine("Trying to find skype");
						SetLyncClient();
					} else if(GetClientState()) {

						switch (GetStatus()) {
							case 3500: SetStatusAvailable(); break;
							case 6500: SetStatusBusy(); break;
							case 9500: SetStatusDoNotDisturb(); break;
							case 15500: SetStatusAway(); break;

						}
					}
				Thread.Sleep(500);
				}
			}


		/// <summary>
		/// Cops the strobe.
		/// </summary>
		public static void CopStrobe() {
			//busylight.Color(BusylightColor.Blue, BusylightColor.Red);
			busylight.Light(BusylightColor.Red);
			Thread.Sleep(500);
			busylight.Light(BusylightColor.Blue);
			Thread.Sleep(500);
		}

		/// <summary>
		/// Happy St Patrick's day
		/// </summary>
		public static void stPatricksDay() {
			busylight.Light(BusylightColor.Green);
			Thread.Sleep(1000);
			busylight.Light(white);
			Thread.Sleep(1000);
			busylight.Light(orange);
			Thread.Sleep(1000);
		}

		/// <summary>
		/// Independances the day.
		/// </summary>
		public static void IndependenceDay() {
			busylight.Light(BusylightColor.Red);
			Thread.Sleep(1000);
			busylight.Light(white);
			Thread.Sleep(1000);
			busylight.Light(BusylightColor.Blue);
			Thread.Sleep(1000);
		}

		/// <summary>
		/// Blues it.
		/// </summary>
		public static void SetStatusAvailable() {
			
			if (CheckDate("March 17")) {
				stPatricksDay();
			} else if (CheckDate("July 4")) {
				IndependenceDay();
			} else if (CheckDate("February 14")) {
				valentinesDay();
			} else {
				StatusAvailable();
			}
		}

		public static void StatusAvailable(){
				busylight.Light(turquoise);
		}

		/// <summary>
		/// Sets the status busy.
		/// </summary>
		public static void SetStatusBusy() {

			Busylight.PulseSequence pulseSequence = Busylight.PulseSequence.Standard;
			pulseSequence.Color = Busylight.BusylightColor.Red;
			busylight.Pulse(pulseSequence);

		}

		/// <summary>
		/// Sets the status away.
		/// </summary>
		public static void SetStatusAway() {
			busylight.Light(orange);
		}

		/// <summary>
		/// Sets the status do not disturb.
		/// </summary>
		public static void SetStatusDoNotDisturb() {
			busylight.Light(BusylightColor.Red);
		}

		/// <summary>
		/// Valentineses the day.
		/// </summary>
		public static void valentinesDay() {
			busylight.Light(violet);
			//Thread.Sleep(500);
		}

		/// <summary>
		/// Checks the date.
		/// </summary>
		/// <returns><c>true</c>, if date was checked, <c>false</c> otherwise.</returns>
		/// <param name="testDate">Test date.</param>
		public static bool CheckDate(string testDate) {

			DateTime date = DateTime.Today;
			if (date.ToString("M") == testDate) {
				return true;
			} else {
				return false;
			}

		}

		/// <summary>
		/// Gets the status.
		/// </summary>
		/// <returns>The status.</returns>
		public static int GetStatus() {
			
			//Console.WriteLine("Client State " + skypeClient.Self.Contact.GetContactInformation(ContactInformationType.Availability));
			return (int)skypeClient.Self.Contact.GetContactInformation(ContactInformationType.Availability);

		}


		/// <summary>
		/// Gets the state of the client.
		/// </summary>
		/// <returns><c>true</c>, if client state was gotten, <c>false</c> otherwise.</returns>
		public static bool GetClientState() {
			if (skypeClient.State == ClientState.SignedIn) {
				return true;
			} else {
				return false;
			}
		}

		/// <summary>
		/// Sets the lync client.
		/// </summary>
		public static void SetLyncClient(){
			try{
				skypeClient = LyncClient.GetClient();
			}catch(Exception e){
				Console.WriteLine("The program experienced an error: " + e);
			}
		}

	//	public static bool CheckForMissedMessages(){

	//		skypeClient.ConversationManager.
	//	}

		/*	public static void SetStatus() {
		var idleTime = IdleTime.IdleTimeDetector.GetIdleTimeInfo();
		if (idleTime.IdleTime.TotalSeconds >= 10) {
			SetClientStateIdle();
		} else {
			SetClientStateAvailable();
		}
	}		

	public static void SetClientStateIdle() {
					if (GetClientState()) {
						if (GetStatus() == 3500) {
							Dictionary<PublishableContactInformationType, object> status = new Dictionary<PublishableContactInformationType, object>();
							status.Add(PublishableContactInformationType.Availability, ContactAvailability.Away);
							skypeClient.Self.BeginPublishContactInformation(status, null, null);
						}
					}
				}

				public static void SetClientStateAvailable() {
					if (GetClientState()) {
						if (GetStatus() == 15500) {
							Dictionary<PublishableContactInformationType, object> status = new Dictionary<PublishableContactInformationType, object>();
							status.Add(PublishableContactInformationType.Availability, ContactAvailability.Free);
							skypeClient.Self.BeginPublishContactInformation(status, null, null);
						}
					}
				}*/
	}
}