//tabs=4
// --------------------------------------------------------------------------------
// ASCOM Dome driver for Banderita
//
// Description:	Handles the Arduino based Dome controller for La Banderita 
//				observatory main dome.
//
// Implements:	ASCOM Dome interface version: 6.2
// Author:		(PIR) Emilio Primucci <eprimucci@gmail.com>
//
// Edit Log:
//
// Date			Who	Vers	Description
// -----------	---	-----	-------------------------------------------------------
// 10-08-2016	PIR	6.0.0	Initial edit, created from ASCOM driver template
// 16-03-2017   PIR 6.2.5   Reviews and FindHome implemented
// --------------------------------------------------------------------------------
//

#define Dome

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using System.Runtime.InteropServices;

using ASCOM;
using ASCOM.Astrometry;
using ASCOM.Astrometry.AstroUtils;
using ASCOM.Utilities;
using ASCOM.DeviceInterface;
using System.Globalization;
using System.Collections;
using System.Linq;

namespace ASCOM.Banderita {
    //
    // Your driver's DeviceID is ASCOM.Banderita.Dome
    //
    /// <summary>
    /// ASCOM Dome Driver for Banderita.
    /// </summary>
    [Guid("0612f073-aefa-4a0b-ac12-77f9e8774425")]
    [ClassInterface(ClassInterfaceType.None)]
    public class Dome : IDomeV2 {
        /// <summary>
        /// ASCOM DeviceID (COM ProgID) for this driver.
        /// The DeviceID is used by ASCOM applications to load the driver at runtime.
        /// </summary>
        internal static string driverID = "ASCOM.Banderita.Dome";
        /// <summary>
        /// Driver para Domo Observatorio La Banderita con Inverter Adlee MS2-IPM.
        /// </summary>
        private static string driverDescription = "ASCOM Dome Driver para La Banderita.";

        internal static string comPortProfileName = "COM Port"; // Constants used for Profile persistence
        internal static string comPortDefault = "COM14";
        internal static string traceStateProfileName = "Trace Level";
        internal static string traceStateDefault = "false";

        internal static string comPort; // Variables to hold the currrent device configuration

        /// <summary>
        /// Private variable to hold the connected state
        /// </summary>
        private bool connectedState;

        /// <summary>
        /// Private variable to hold an ASCOM Utilities object
        /// </summary>
        private Util utilities;

        /// <summary>
        /// Private variable to hold an ASCOM AstroUtilities object to provide the Range method
        /// </summary>
        private AstroUtils astroUtilities;

        /// <summary>
        /// Variable to hold the trace logger object (creates a diagnostic log file with information that you specify)
        /// </summary>
        internal static TraceLogger tl;


        public static string s_csDriverID = "ASCOM.Banderita.Dome";
        public static string s_csDriverDescription = "La Banderita groso Domo";

        private ArduinoSerial SerialConnection;

        private Util HC = new Util();
        private Config Config = new Config();
        private ArrayList supportedActions = new ArrayList();

        /// <summary>
        /// Initializes a new instance of the <see cref="Banderita"/> class.
        /// Must be public for COM registration.
        /// </summary>
        public Dome() {
            tl = new TraceLogger("", "Banderita");
            ReadProfile(); // Read device configuration from the ASCOM Profile store

            tl.LogMessage("Dome", "Starting initialisation");

            connectedState = false; // Initialise connected to false
            utilities = new Util(); //Initialise util object
            astroUtilities = new AstroUtils(); // Initialise astro utilities object

            // Implemented methods
            supportedActions.Add("AbortSlew");
            supportedActions.Add("CloseShutter");
            // implementedMethods.Add("CommandBlind");
            // implementedMethods.Add("CommandBool");
            // implementedMethods.Add("CommandString");
            supportedActions.Add("Dispose");
            supportedActions.Add("FindHome");
            supportedActions.Add("OpenShutter");
            supportedActions.Add("Park");
            supportedActions.Add("SetPark");
            //supportedActions.Add("SetupDialog");
            //supportedActions.Add("SlewToAltitude");
            supportedActions.Add("SlewToAzimuth");
            supportedActions.Add("SyncToAzimuth");


            tl.LogMessage("Dome", "Completed initialisation. Supported: "+ string.Join(",", supportedActions.ToArray().Select(o => o.ToString()).ToArray()));
        }


        //
        // PUBLIC COM INTERFACE IDomeV2 IMPLEMENTATION
        //

        #region Common properties and methods.

        /// <summary>
        /// Displays the Setup Dialog form.
        /// If the user clicks the OK button to dismiss the form, then
        /// the new settings are saved, otherwise the old values are reloaded.
        /// THIS IS THE ONLY PLACE WHERE SHOWING USER INTERFACE IS ALLOWED!
        /// </summary>
        public void SetupDialog() {
            // consider only showing the setup dialog if not connected
            // or call a different dialog if connected
            if (IsConnected)
                System.Windows.Forms.MessageBox.Show("Ya está conectado, presione OK");

            using (SetupDialogForm F = new SetupDialogForm()) {
                var result = F.ShowDialog();
                if (result == System.Windows.Forms.DialogResult.OK) {
                    WriteProfile(); // Persist device configuration values to the ASCOM Profile store
                }
            }
        }
        
        public ArrayList SupportedActions {
            get {
                tl.LogMessage("SupportedActions Get", "Returning implemented methods: "+ 
                    string.Join(",", supportedActions.ToArray().Select(o => o.ToString()).ToArray())
                    );
                return supportedActions;
            }
        }

        public string Action(string actionName, string actionParameters) {
            LogMessage("", "Action {0}, parameters {1} not implemented", actionName, actionParameters);
            throw new ASCOM.ActionNotImplementedException("Action " + actionName + " no está implementada por este driver");
        }

        public void CommandBlind(string command, bool raw) {
            CheckConnected("CommandBlind");
            // Call CommandString and return as soon as it finishes
            //this.CommandString(command, raw);
            // or
            throw new ASCOM.MethodNotImplementedException("CommandBlind");
            // DO NOT have both these sections!  One or the other
        }

        public bool CommandBool(string command, bool raw) {
            CheckConnected("CommandBool");
            //string ret = CommandString(command, raw);
            // TODO decode the return string and return true or false
            // or
            throw new ASCOM.MethodNotImplementedException("CommandBool");
            // DO NOT have both these sections!  One or the other
        }

        public string CommandString(string command, bool raw) {
            CheckConnected("CommandString");
            // it's a good idea to put all the low level communication with the device here,
            // then all communication calls this function
            // you need something to ensure that only one command is in progress at a time

            throw new ASCOM.MethodNotImplementedException("CommandString");
        }

        public void Dispose() {
            // Clean up the tracelogger and util objects
            tl.Enabled = false;
            tl.Dispose();
            tl = null;
            utilities.Dispose();
            utilities = null;
            astroUtilities.Dispose();
            astroUtilities = null;
        }



        public string Description {
            // TODO customise this device description
            get {
                tl.LogMessage("Description Get", driverDescription);
                return driverDescription;
            }
        }

        public string DriverInfo {
            get {
                Version version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
                // TODO customise this driver description
                string driverInfo = "Information about the driver itself. Version: " + String.Format(CultureInfo.InvariantCulture, "{0}.{1}", version.Major, version.Minor);
                tl.LogMessage("DriverInfo Get", driverInfo);
                return driverInfo;
            }
        }

        public string DriverVersion {
            get {
                Version version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
                string driverVersion = String.Format(CultureInfo.InvariantCulture, "{0}.{1}", version.Major, version.Minor);
                tl.LogMessage("DriverVersion Get", driverVersion);
                return driverVersion;
            }
        }

        public short InterfaceVersion {
            // set by the driver wizard
            get {
                LogMessage("InterfaceVersion Get", "2");
                return Convert.ToInt16("2");
            }
        }

        public string Name {
            get {
                string name = "Banderita dome driver";
                tl.LogMessage("Name Get", name);
                return name;
            }
        }

        #endregion

        #region IDome Implementation

        private bool domeShutterState = false; // Variable to hold the open/closed status of the shutter, true = Open


        public void AbortSlew() {
            SerialConnection.SendCommand(ArduinoSerial.SerialCommand.Abort);
            this.Slaved = false;
            tl.LogMessage("AbortSlew", "Completed");
        }

        public double Altitude {
            get {
                tl.LogMessage("Altitude Get", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("Altitude", false);
            }
        }

        public bool AtHome {
            get { return this.Config.AtHome; }
        }

        public bool AtPark {
            get { return this.Config.Parked; }
        }

        public double Azimuth {
            get { return this.Config.Azimuth; }
        }

        public bool CanFindHome {
            get { return true; }
        }

        public bool CanPark {
            get { return true; }
        }

        public bool CanSetAltitude {
            get { return false; }
        }

        public bool CanSetAzimuth {
            get { return true; }
        }

        public bool CanSetPark {
            get { return true; }
        }

        public bool CanSetShutter {
            get { return true; }
        }

        public bool CanSlave {
            get { return true; }
        }

        public bool CanSyncAzimuth {
            get { return true; }
        }

        public void CloseShutter() {
            this.Config.ShutterStatus = ShutterState.shutterClosing;
            SerialConnection.SendCommand(ArduinoSerial.SerialCommand.CloseShutter);

            while (this.Config.ShutterStatus == ShutterState.shutterClosed)
                HC.WaitForMilliseconds(100);

            domeShutterState = false;
        }


        public bool Connected {
            get { return this.Config.Link; }
            set {
                switch (value) {
                    case true:
                        this.Config.Link = this.ConnectDome();
                        connectedState = true;
                        break;
                    case false:
                        this.Config.Link = !this.DisconnectDome();
                        connectedState = false;
                        break;
                }
            }
        }

        private bool DisconnectDome() {
            SerialConnection.Close();
            return true;
        }

        private bool ConnectDome() {
            SerialConnection = new ArduinoSerial();
            SerialConnection.CommandQueueReady += new ArduinoSerial.CommandQueueReadyEventHandler(SerialConnection_CommandQueueReady);
            HC.WaitForMilliseconds(2000);

            return true;
        }


        public void FindHome() {
            if (this.Slaved) {
                tl.LogMessage("FindHome", "Unable. I am Slaved!");
                throw new ASCOM.SlavedException("El domo está en Slave. No puedo encontrar HOME así!");
            }
            this.Config.Parked = false;
            this.Config.IsSlewing = true;
            this.Config.AtHome = false;
            SerialConnection.SendCommand(ArduinoSerial.SerialCommand.FindHome);
            tl.LogMessage("FindHome", "Just started...."+ ArduinoSerial.SerialCommand.FindHome);

            while (!this.Config.AtHome)
                HC.WaitForMilliseconds(100);
        }


        void SerialConnection_CommandQueueReady(object sender, EventArgs e) {
            while (SerialConnection.CommandQueue.Count > 0) {
                string[] com_args = ((string)SerialConnection.CommandQueue.Pop()).Split(' ');

                string command = com_args[0];

                switch (command) {
                    case "HOMED":
                        this.Config.AtHome = true;
                        this.Config.Azimuth = Int32.Parse(com_args[1]);
                        this.Config.IsSlewing = false;
                        this.Config.HomePosition = this.Config.Azimuth;
                        tl.LogMessage("FindHome", "Completed");
                        break;
                    case "P":
                        this.Config.Azimuth = Int32.Parse(com_args[1]);
                        this.Config.IsSlewing = false;
                        if(this.Config.HomePosition == this.Config.Azimuth) {
                            this.Config.AtHome = true;
                        }
                        else {
                            this.Config.AtHome = false;
                        }
                        break;
                    case "SHUTTER":
                        this.Config.ShutterStatus = (com_args[1] == "OPEN") ? ShutterState.shutterOpen : ShutterState.shutterClosed;
                        break;
                    case "SYNCED":
                        this.Config.Synced = true;
                        break;
                    case "PARKED":
                        this.Config.Parked = true;
                        break;
                    default:
                        break;
                }
            }
        }

        public void OpenShutter() {
            this.Config.ShutterStatus = ShutterState.shutterOpening;
            SerialConnection.SendCommand(ArduinoSerial.SerialCommand.OpenShutter);

            while (this.Config.ShutterStatus == ShutterState.shutterOpening)
                HC.WaitForMilliseconds(100);

            domeShutterState = true;
        }

        public void Park() {
            SerialConnection.SendCommand(ArduinoSerial.SerialCommand.Park, this.Config.ParkPosition);

            while (!this.Config.Parked)
                HC.WaitForMilliseconds(100);
        }

        public void SetPark() {
            this.Config.ParkPosition = this.Config.Azimuth;
        }

        public ShutterState ShutterStatus {
            get { return this.Config.ShutterStatus; }
        }

        public bool Slaved {
            get { return this.Config.Slaved; }
            set { this.Config.Slaved = value; }
        }

        public void SlewToAltitude(double Altitude) {
            tl.LogMessage("SlewToAltitude", "Not implemented");
            throw new ASCOM.MethodNotImplementedException("SlewToAltitude");
        }

        public void SlewToAzimuth(double Azimuth) {
            if (Azimuth > 360 || Azimuth < 0)
                throw new ASCOM.InvalidValueException("Azimuth fuera de rango!");
            this.Config.IsSlewing = true;
            SerialConnection.SendCommand(ArduinoSerial.SerialCommand.Slew, Azimuth);

            while (this.Config.IsSlewing)
                HC.WaitForMilliseconds(100);
        }

        public bool Slewing {
            get { return this.Config.IsSlewing; }
        }


        public void SyncToAzimuth(double Azimuth) {
            this.Config.Synced = false;
            if (Azimuth > 360 || Azimuth < 0)
                throw new ASCOM.InvalidValueException("Azimuth fuera de rango!");
            SerialConnection.SendCommand(ArduinoSerial.SerialCommand.SyncToAzimuth, Azimuth);

            while (!this.Config.Synced)
                HC.WaitForMilliseconds(100);
        }

        #endregion

        #region Private properties and methods
        // here are some useful properties and methods that can be used as required
        // to help with driver development

        #region ASCOM Registration

        // Register or unregister driver for ASCOM. This is harmless if already
        // registered or unregistered. 
        //
        /// <summary>
        /// Register or unregister the driver with the ASCOM Platform.
        /// This is harmless if the driver is already registered/unregistered.
        /// </summary>
        /// <param name="bRegister">If <c>true</c>, registers the driver, otherwise unregisters it.</param>
        private static void RegUnregASCOM(bool bRegister) {
            using (var P = new ASCOM.Utilities.Profile()) {
                P.DeviceType = "Dome";
                if (bRegister) {
                    P.Register(driverID, driverDescription);
                }
                else {
                    P.Unregister(driverID);
                }
            }
        }

        /// <summary>
        /// This function registers the driver with the ASCOM Chooser and
        /// is called automatically whenever this class is registered for COM Interop.
        /// </summary>
        /// <param name="t">Type of the class being registered, not used.</param>
        /// <remarks>
        /// This method typically runs in two distinct situations:
        /// <list type="numbered">
        /// <item>
        /// In Visual Studio, when the project is successfully built.
        /// For this to work correctly, the option <c>Register for COM Interop</c>
        /// must be enabled in the project settings.
        /// </item>
        /// <item>During setup, when the installer registers the assembly for COM Interop.</item>
        /// </list>
        /// This technique should mean that it is never necessary to manually register a driver with ASCOM.
        /// </remarks>
        [ComRegisterFunction]
        public static void RegisterASCOM(Type t) {
            RegUnregASCOM(true);
        }

        /// <summary>
        /// This function unregisters the driver from the ASCOM Chooser and
        /// is called automatically whenever this class is unregistered from COM Interop.
        /// </summary>
        /// <param name="t">Type of the class being registered, not used.</param>
        /// <remarks>
        /// This method typically runs in two distinct situations:
        /// <list type="numbered">
        /// <item>
        /// In Visual Studio, when the project is cleaned or prior to rebuilding.
        /// For this to work correctly, the option <c>Register for COM Interop</c>
        /// must be enabled in the project settings.
        /// </item>
        /// <item>During uninstall, when the installer unregisters the assembly from COM Interop.</item>
        /// </list>
        /// This technique should mean that it is never necessary to manually unregister a driver from ASCOM.
        /// </remarks>
        [ComUnregisterFunction]
        public static void UnregisterASCOM(Type t) {
            RegUnregASCOM(false);
        }

        #endregion

        /// <summary>
        /// Returns true if there is a valid connection to the driver hardware
        /// </summary>
        private bool IsConnected {
            get {
                // TODO check that the driver hardware connection exists and is connected to the hardware
                return connectedState;
            }
        }

        /// <summary>
        /// Use this function to throw an exception if we aren't connected to the hardware
        /// </summary>
        /// <param name="message"></param>
        private void CheckConnected(string message) {
            if (!IsConnected) {
                throw new ASCOM.NotConnectedException(message);
            }
        }

        /// <summary>
        /// Read the device configuration from the ASCOM Profile store
        /// </summary>
        internal void ReadProfile() {
            using (Profile driverProfile = new Profile()) {
                driverProfile.DeviceType = "Dome";
                tl.Enabled = Convert.ToBoolean(driverProfile.GetValue(driverID, traceStateProfileName, string.Empty, traceStateDefault));
                comPort = driverProfile.GetValue(driverID, comPortProfileName, string.Empty, comPortDefault);
            }
        }

        /// <summary>
        /// Write the device configuration to the  ASCOM  Profile store
        /// </summary>
        internal void WriteProfile() {
            using (Profile driverProfile = new Profile()) {
                driverProfile.DeviceType = "Dome";
                driverProfile.WriteValue(driverID, traceStateProfileName, tl.Enabled.ToString());
                driverProfile.WriteValue(driverID, comPortProfileName, comPort.ToString());
            }
        }

        /// <summary>
        /// Log helper function that takes formatted strings and arguments
        /// </summary>
        /// <param name="identifier"></param>
        /// <param name="message"></param>
        /// <param name="args"></param>
        internal static void LogMessage(string identifier, string message, params object[] args) {
            var msg = string.Format(message, args);
            tl.LogMessage(identifier, msg);
        }
        #endregion
    }
}
