using System;
using System.Threading;
using PyramidServiceObject;
using Microsoft.PointOfService;




namespace TestRGBUPOS
{
    internal class Program
    {
        public static void Main(string[] args)
        {
            // Create Pos Explorer
            PosExplorer posExplorer = new PosExplorer();
            DeviceInfo lightInfo;
            try
            {
                lightInfo = posExplorer.GetDevice("Lights");
                Console.WriteLine(lightInfo.ServiceObjectName);
            }
            catch (Exception e)
            {
                Console.WriteLine("No Light found");
                return;
            }
            
            Console.WriteLine("Light Found");
            Lights light = (Lights) posExplorer.CreateInstance(lightInfo);
            light.Open();
            Console.WriteLine("Light Opened");
            light.Claim(1000);
            Console.WriteLine("Light Claimed");
            light.DeviceEnabled = true;
            light.SwitchOn(1,0,0,LightColors.Primary, LightAlarms.None);
            //wait 1 second
            Thread.Sleep(1000);
            light.SwitchOff(1);
        }   
    }
}