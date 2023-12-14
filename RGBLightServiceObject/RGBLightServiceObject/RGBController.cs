/// This Class is used to control the RGB Light Controller from Pyramid Computers.
/// Author: Julian Lenzenweger
/// Date: 07.12.2023
///

using System;
using System.Collections.Generic;
using System.IO.Ports;
using System.Linq;


namespace LedSerialControl;

public class RGBController
{
    /// <summary>
    /// Enum representing the different colors that the LED strip can be set to.
    /// </summary>
    public enum ColorNames
    {
        Red,
        Green,
        Blue,
        White,
        Yellow,
        Orange,
        Cyan,
        Magenta,
        Off
    }

    /// <summary>
    /// Dictionary mapping ColorNames to their corresponding byte arrays.
    /// </summary>
    private static readonly Dictionary<ColorNames, byte[]> ColorDictionaryRgbw = new Dictionary<ColorNames, byte[]>
    {
        { ColorNames.Red, new byte[] { 255, 0, 0, 0 } } ,
        { ColorNames.Green, new byte[] { 0, 255, 0, 0 } },
        { ColorNames.Blue, new byte[] { 0, 0, 255, 0 } },
        { ColorNames.White, new byte[] { 0, 0, 0, 255 } },
        { ColorNames.Yellow, new byte[] { 255, 255, 0, 0 } },
        { ColorNames.Orange, new byte[] { 255, 165, 0, 0 } },
        { ColorNames.Cyan, new byte[] { 0, 255, 255, 0 } },
        { ColorNames.Magenta, new byte[] { 255, 0, 255, 0 } },
        { ColorNames.Off, new byte[] { 0, 0, 0, 0 } }
    };
    
    /// <summary>
    /// Struct representing a RGBW Value
    /// </summary>
    private class RgbwValue
    {
        public byte Red { get; set; }
        public byte Green { get; set; }
        public byte Blue { get; set; }
        public byte White { get; set; }

        public RgbwValue(byte r, byte g, byte b, byte w)
        {
            Red = r;
            Green = g;
            Blue = b;
            White = w;
        }

        ///Overload assignment operator from byte[] to RgbwValue
        /// <summary>
        /// Overloaded assignment operator from byte[] to RgbwValue
        /// </summary>
        /// <param name="values">Byte Array of length 4 with values for {red, green, blue, white}</param>
        public static implicit operator RgbwValue(byte[] values)
        {
            if (values.Length < 4)
                throw new ArgumentException("Byte Array must be of length 4");
            return new RgbwValue(values[0], values[1], values[2], values[3]);
        }
    }
    
    private enum Command
    {
        SetColor = 0xCA,
        ChangeFlashingColors = 0xD3,
        Mode = 0xD6,
        FlashingPeriod = 0xE5,
        SetID = 0xAE,
        ReadID = 0xBE
    }

    
    private static int _baudrate { get; set; }
    private const int _databits = 8;
    private SerialPort _serialPort;
    
    private byte[] _startbits = { 0x5A, 0xFF };
    private byte[] _endbits = { 0xA5 };
    

    private const byte _off = 0;
    private const byte _on = 1;
    
    // Properties for the current color and flashing state
    private RgbwValue _lastColor = new RgbwValue(0, 0, 0, 0);
    private RgbwValue _savedColor = new RgbwValue(0, 0, 0, 0);
    public bool Flashing { get; private set; } = false;
    

    
    
    /// <summary>
    /// Sets the color of the RGB LED strip.
    /// </summary>
    /// <param name="color">The color to set the LED controller to. This is an enum value.</param>
    public void SetColor(ColorNames color)
    {
        byte[] values = ColorDictionaryRgbw[color];
        var argument = ConcatArrays(values, new byte[] { 0, 0, 0, 0 });
        SendCommand(Command.SetColor, argument);
        _lastColor = values;
    }
    
    /// <summary>
    /// Sets the color of the RGB controller in RGBW Values
    /// </summary>
    public void SetColorRgbw( byte r = 0, byte g = 0, byte b = 0, byte w = 0)
    {
        byte[] values = { r, g, b, w, 0, 0, 0, 0 };
        SendCommand(Command.SetColor, values);
        _lastColor = values;
    }
    
    private void SetColorRgbw(RgbwValue color)
    {
        SetColorRgbw(color.Red, color.Green, color.Blue, color.White);
    }
    
    /// <summary>
    /// Sets the color of the RGBW Module in Percent
    /// </summary>
    /// <param name="r">Value of Red</param>
    /// <param name="g">Value of Green </param>
    /// <param name="b">Value of Blue</param>
    /// <param name="w">Value of White</param>
    public void SetColorRgbwPercent( short r = 0, short g = 0, short b = 0, short w = 0)
    {
        if (r > 100 || g > 100 || b > 100 || w > 100 || r < 0 || g < 0 || b < 0 || w < 0)
            throw new ArgumentException("Values must be between 0 and 100");
        byte red = (byte)(r*255/100);
        byte green =  (byte)(g*255/100);
        byte blue =  (byte)(b*255/100);
        byte white =  (byte)(w*255/100);
        SetColorRgbw(red, green, blue, white);
    }
    
    /// <summary>
    ///  Saves the current color of the controller internally, to resume to it later with ResumeColor()
    /// </summary>
    public void SaveColor()
    {
        _savedColor = _lastColor;
    }
    /// <summary>
    /// Sets the static color of the Module to the last saved color. Saved via SaveColor()
    /// </summary>
    public void ResumeColor()
    {
        SetColorRgbw(_savedColor);
    }
    
    /// <summary>
    /// Sets the flashing mode of the RGB LED strip.
    /// </summary>
    /// <param name="flashing">True for flashing. False for static color</param>
    public void SetFlashing(bool flashing)
    {
        if (flashing)
            SendCommand(Command.Mode, _on);
        else
            SendCommand(Command.Mode, _off);
        Flashing = flashing;
    }

    /// <summary>
    /// Sets the flashing period of the RGB LED strip.
    /// Each Step is about 27ms
    /// </summary>
    /// <param name="period">The period to set the LED strip to. This is a byte value.</param>
    public void SetFlashingPeriod(byte period)
    {
        SendCommand(Command.FlashingPeriod, period);
    }
    
    public void SetFlashingColors(ColorNames color1, ColorNames color2)
    {
        byte[] values = ConcatArrays(ColorDictionaryRgbw[color1], ColorDictionaryRgbw[color2]);
        SendCommand(Command.ChangeFlashingColors, values);
    }


    public void SetID(byte id)
    {
        SendCommand(Command.ReadID, id);
    }
    
    public int ReadID()
    {
        byte[] message = ConcatArrays(_startbits, new byte[] { (byte)Command.ReadID }, _endbits);
        _serialPort.Write(message, 0, message.Length);
        try
        {
            int id = _serialPort.ReadByte();
            return id;
        }
        catch (Exception e)
        {
            if (e is TimeoutException)
                return -1;
            throw;
        }
    }
    
    
    /// <summary>
    /// Constructor for RGBController
    /// </summary>
    /// <param name="portname">Serial Port name. Commonly COMx, eg. COM4</param>
    public RGBController(string portname)
    {
        _baudrate = 9600;
        _serialPort = new SerialPort(portname, _baudrate);
        _serialPort.DataBits = _databits; // 8 data bits
        _serialPort.StopBits = StopBits.One;
        _serialPort.Parity = Parity.None;
        _serialPort.Open();
        SetFlashing(false);
        SetColor(ColorNames.Off);
    }

    ~RGBController()
    {
        _serialPort.Close();
    }


    /// <summary>
    /// Sends a command to the RGB Controller. Adds start and end bits.
    /// </summary>
    /// <param name="command"></param>
    /// <param name="arguments"></param>
    private void SendCommand(Command command, byte[] arguments)
    {
        byte[] message = ConcatArrays(_startbits, new byte[] { (byte)command }, arguments, _endbits);
        try
        {
            _serialPort.Write(message, 0, message.Length);
            //Console.Write($"Sent Message: {ByteArrayToHexString(message)}\n");
        }
        catch (Exception e)
        {
            switch (e)
            {
                case TimeoutException _:
                    throw new TimeoutException("Timeout while sending command");
                case InvalidOperationException _:
                    throw new InvalidOperationException("Serial Port is not open");
            }
            Console.WriteLine(e);
            throw;
        }
        
    }
    private void SendCommand(Command command, byte argument)
    {
        SendCommand(command, new byte[] { argument });
    }
    
    
    // Helper Functions for byte array handling
    
    /// <summary>
    /// Help function to concat multiple byte arrays
    /// </summary>
    /// <param name="arrays"></param>
    /// <returns></returns>
    private static byte[] ConcatArrays(params byte[][] arrays)
    {
        return arrays.SelectMany(arr => arr).ToArray();
    }
    
    /// <summary>
    /// Help function to convert a byte array to a hex string
    /// </summary>
    /// <param name="bytes"></param>
    /// <returns></returns>
    private static string ByteArrayToHexString(byte[] bytes)
    {
        return BitConverter.ToString(bytes).Replace("-", "");
    }
    
    /// <summary>
    /// Converts a Hex String to Array of Bytes
    /// </summary>
    /// <param name="hex"></param>
    /// <returns>Eschowissn</returns>
    private static byte[] StringToByteArray(string hex)
    {
        int numberChars = hex.Length;
        byte[] bytes = new byte[numberChars / 2];
        for (int i = 0; i < numberChars; i += 2)
        {
            bytes[i / 2] = Convert.ToByte(hex.Substring(i, 2), 16);
        }
        return bytes;
    }
}

