# PowerBlueServerWindows
This is a BlueToothServer which runs on windows and controls the power point slide show. The client can be any mobile or desktop application.

This application Developing Environment details:
------------------------------------------------

Developed In Language: Visual Basics

Developed using IDE: Microsoft Visual Studio Express 2013 for Windows Desktop version 12.0.21005.1 REL

Microsoft .NET Framework Version: 4.5.50938



This application Testing Environment details:
------------------------------------------------
The application is tested and working fine on Environment

OS: Microsoft Windows 7 version: 6.1.7601 SP1 Build 7601

processor Bit: i7-3635QM 64 Bit Machine

RAM: 8GB


Using the Framework 32Feet.Net from "In The Hand Ltd":
------------------------------------------------------
This application is developed on top of the 32Feet.net bluetooth framework version 3.5.0.0.

Framework: 32feet.NET - Personal Area Networking for .NET

WebSite: http://32feet.codeplex.com/


For developing the application:
-------------------------------
Make sure the PC is connected to Internet.

1. Go To Visual Studio Express 2013 -> PROJECT menu -> Manage NuGet Packages -> Install 32feet.NET NuGet Packages

2. Go To Visual Studio Express 2013 -> PROJECT menu -> Add Reference -> COM -> Type Libraries -> Search for "Microsoft Office 14.0 Object Library" version 2.5 and -> check it with check box -> add this as reference.


For installing the Power Blue application: Prerequistes
-------------------------------------------------------
Make sure the PC is connected to Internet for downloading 32feet.net.


First Bluetooth:

1.Bluetooth Adapter should be enabled and Bluetooth should be switched on in Device. 
  Check this from Devicemanager in windows.

2. For clients to detect your windows machine, make sure that in 

   Bluetooth settings -> Options > Discovery -> check the box "Allow bluetooth devices to find this computer"
   
   Bluetooth settings -> Options > Connections -> check the box "Allow bluetooth devices to connect to this computer"
   
   Bluetooth settings -> Options > Connections -> check the box "Alert me when a new bluetooth device wants to connect"


Second Install: 32Feet.Net

1. Go To http://32feet.codeplex.com/

2. Download the 32feet.Net set up zip for 3.5.0.0 or higher. 
   http://32feet.codeplex.com/releases/view/88941

3. Unzip the Download & install 32feet.net software by running the setup.exe file.


Third Install: PowerBlue Server Application


Power Blue Server Installers or Setups
-------------------------------------------------------
Different versions of Power Blue Server Installers will be present under the link below.
Please find the installation instructions also in the same link.

https://github.com/srib4ufrnd/PowerBlueServerWindows/tree/master/SetUps
