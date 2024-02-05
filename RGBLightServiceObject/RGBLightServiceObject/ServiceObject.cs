using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using LedSerialControl;
using Microsoft.PointOfService;
using Microsoft.PointOfService.BasicServiceObjects;
using System.Threading;

namespace PyramidServiceObject
{
    [HardwareId("USB\\VID_03EB&PID_2404&REV_0100")]
    [ServiceObject(
        DeviceType.Lights,
        "RGBLightServiceObject",
        "RGBLightServiceObject vom Lenzenweger Julian 8)",
        1,
        14)]
    public class ServiceObject : LightsBasic
    {
        private RGBController _rgbController; // The RGB Controller object for communicating with the device

        //"All properties are initialized by the open() method"

        public override int MaxLights { get; } = 1;

        public override LightColors CapColor { get; } = LightColors.Primary | LightColors.Custom1 |
                                                        LightColors.Custom2 | LightColors.Custom3 |
                                                        LightColors.Custom4 | LightColors.Custom5;

        public override LightAlarms CapAlarm { get; } = LightAlarms.None;
        public override bool CapBlink { get; } = true; // Blinking is supported

        public override void SwitchOff(int lightnumber)
        {
            if (lightnumber != 1)
                throw new PosControlException("Invalid light Selected, only 1 Light supported",
                    Microsoft.PointOfService.ErrorCode.Illegal);
            else
                _rgbController.SetColor(RGBController.ColorNames.Off);
        }


        public override void SwitchOn(int lightnumber, int blinkOnCycle, int blinkOffCycle, LightColors colors,
            LightAlarms alarms = LightAlarms.None)
        {
            if (lightnumber != 1)
                throw new PosControlException("Invalid light Selected, only 1 Light supported",
                    Microsoft.PointOfService.ErrorCode.Illegal);
            if (blinkOffCycle < 0 || blinkOnCycle < 0)
                throw new ArgumentException("BlinkOnCycle and BlinkOffCycle must be positive");
            if (!Enum.IsDefined(typeof(LightAlarms), alarms))
                throw new InvalidEnumArgumentException(nameof(alarms), (int)alarms, typeof(LightAlarms));
            if (!Enum.IsDefined(typeof(LightColors), colors))
                throw new InvalidEnumArgumentException(nameof(colors), (int)colors, typeof(LightColors));

            cts.Cancel(); // Cancel the Blinking Thread if it is running
            _rgbController.SetColor(
                colors switch
                {
                    LightColors.Primary => RGBController.ColorNames.Green,
                    LightColors.Custom1 => RGBController.ColorNames.Red,
                    LightColors.Custom2 => RGBController.ColorNames.Yellow,
                    LightColors.Custom3 => RGBController.ColorNames.Orange,
                    LightColors.Custom4 => RGBController.ColorNames.Blue,
                    LightColors.Custom5 => RGBController.ColorNames.Magenta,
                    _ => RGBController.ColorNames.Off
                }
            );

            if (blinkOnCycle != 0 && blinkOffCycle != 0)
            {
                cts = new CancellationTokenSource();
                Thread blinkThread = new Thread(() => BlinkThread(blinkOnCycle, blinkOffCycle, cts.Token));
                blinkThread.Start();
            }
        }

        public ServiceObject()
        {
            _rgbController = new RGBController("COM4");
        }

        public override string CheckHealth(HealthCheckLevel level)
        {
            switch (level)
            {
                case HealthCheckLevel.Interactive:
                    throw new PosControlException("Interactive CheckHealth not supported",
                        Microsoft.PointOfService.ErrorCode.Illegal);

                case HealthCheckLevel.Internal:
                    int id = _rgbController.ReadID();
                    if (id == -1)
                    {
                        _healthText = "Internal HCheck: Failed";
                        return "Internal HCheck: Failed";
                    }

                    _healthText = "Internal HCheck: Successful";
                    return "Internal HCheck: Successful";

                case HealthCheckLevel.External:
                    _rgbController.SaveColor();
                    // Cycle through all colors
                    foreach (var color in Enum.GetValues(typeof(RGBController.ColorNames)))
                    {
                        _rgbController.SetColor((RGBController.ColorNames)color);
                        Thread.Sleep(1000);
                    }

                    _rgbController.ResumeColor();
                    _healthText = "External HCheck: Complete";
                    break;
            }

            return "OK";
        }

        public override DirectIOData DirectIO(int command, int data, object obj)
        {
            throw new NotImplementedException();
        }

        private string _healthText = "";
        public override string CheckHealthText => _healthText;

        static CancellationTokenSource cts = new CancellationTokenSource();

        /// <summary>
        /// This method is responsible for creating a blinking effect on the RGB light.
        /// </summary>
        /// <param name="onTime">The amount of time (in milliseconds) the light stays on during each blink cycle.</param>
        /// <param name="offTime">The amount of time (in milliseconds) the light stays off during each blink cycle.</param>
        /// <param name="token">A cancellation token that can be used to cancel the blinking effect.</param>
        void BlinkThread(int onTime, int offTime, CancellationToken token)
        {
            _rgbController.SaveColor();
            while (true)
            {
                if (token.IsCancellationRequested)
                    break;
                _rgbController.ResumeColor();
                Thread.Sleep(onTime);
                _rgbController.SetColor(RGBController.ColorNames.Off);
                Thread.Sleep(offTime);
            }
        }
    }
}