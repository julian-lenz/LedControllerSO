using System;
using System.ComponentModel;
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
            if(lightnumber != 1)
                throw new PosControlException("Invalid light Selected, only 1 Light supported", Microsoft.PointOfService.ErrorCode.Illegal);
            else
                _rgbController.SetColor(RGBController.ColorNames.Off);
        }
        
        
        public override void SwitchOn(int lightnumber, int blinkOnCycle, int blinkOffCycle, LightColors colors, LightAlarms alarms = LightAlarms.None)
        {
            if (lightnumber != 1)
                throw new PosControlException("Invalid light Selected, only 1 Light supported", Microsoft.PointOfService.ErrorCode.Illegal);
            if(blinkOffCycle < 0 || blinkOnCycle < 0)
                throw new ArgumentException("BlinkOnCycle and BlinkOffCycle must be positive");
            if (!Enum.IsDefined(typeof(LightAlarms), alarms))
                throw new InvalidEnumArgumentException(nameof(alarms), (int)alarms, typeof(LightAlarms));
            if (!Enum.IsDefined(typeof(LightColors), colors))
                throw new InvalidEnumArgumentException(nameof(colors), (int)colors, typeof(LightColors));
            /*
             * https://learn.microsoft.com/en-us/dotnet/api/microsoft.pointofservice.lightcolors?view=point-of-service-1.14
             * LightColors consists of
             * Custom1	65536	Supports first custom color (usually red).
             * Custom2	131072	Supports second custom color (usually yellow).
             * Custom3	262144	Supports third custom color.
             * Custom4	524288	Supports fourth custom color.
             * Custom5  1048576	Supports fifth custom color.
             * Primary	1       Supports primary color (usually green).
             */
            /*
             * https://learn.microsoft.com/en-us/dotnet/api/microsoft.pointofservice.lightalarms?view=point-of-service-1.14
             * LightAlarms Enum consist of
             * Custom 1 = 65536
             * Custom 2 = 131072
             * Fast = 64
             * Medium = 32
             * None = 1
             * Slow = 16
             * 
             */
            switch (colors)
            {
                case LightColors.Primary:
                    _rgbController.SetColor(RGBController.ColorNames.Green);
                    break;
                case LightColors.Custom1:
                    _rgbController.SetColor(RGBController.ColorNames.Red);
                    break;
                case LightColors.Custom2:
                    _rgbController.SetColor(RGBController.ColorNames.Yellow);
                    break;
                case LightColors.Custom3:
                    _rgbController.SetColor(RGBController.ColorNames.Orange);
                    break;
                case LightColors.Custom4:
                    _rgbController.SetColor(RGBController.ColorNames.Blue);
                    break;
                case LightColors.Custom5:
                    _rgbController.SetColor(RGBController.ColorNames.Magenta);
                    break;
            }

            if (blinkOnCycle != 0)
            {
                // Convert from ms to appropriate value between 0 and 255
                var steps = (blinkOnCycle / 27);
                if (steps > 255)
                    steps = 255;
                byte periodByte = (byte)steps;
                _rgbController.SetFlashingPeriod(periodByte);
                _rgbController.SetFlashing(true);
            }
            else
                _rgbController.SetFlashing(false);
            
            
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
                    throw new PosControlException("Interactive CheckHealth not supported", Microsoft.PointOfService.ErrorCode.Illegal);
                
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
    }
}
