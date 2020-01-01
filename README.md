# The DeviceDetector ActiveX OOP EXE server project

Windows provides information about different sorts of attached devices, like hard disks, CDROMs, printers and others. These devices can be attached to the computer via different interfaces like USB, SCSI, serial or parallel ports, etc.

See a sample video of what you can do with it in an Access Database application (DeviceDetectorAuthDemo.accdb) here:

https://www.youtube.com/watch?v=qh5hrhpRNIg

AxDeviceDetector.exe is a (32 bits) ActiveX OOP (out of process) EXE Server that exposes 2 classes:

* DeviceDetector
  * This is the class that will raise the events when they're notified by Windows, in real time, as a device is plugged (or a network drive is connected, or media is inserted in a device) or respectively unplugged in the system.
* DeviceInfo
  * This is a wrapper class, around the functionality provided by the [DeviceInfo C DLL](https://github.com/francescofoti/deviceinfo_dll) project (which itself is 32/64 bits).

There are three projects in this repository:

* AxDeviceDetector.vbp

  * This is the project that implements the ActiveX EXE server.
    You have to run this one with administrator privileges, so it can register the ActiveX classes (nothing is displayed if the registration succeeds).

* SaDeviceDetector.vbp

  * This is not an ActiveX server, just a standalone executable that serves as sample and demo.

    You can just run this executable, *no need to register the ActiveX server for it to function*, as it does not use the classes via ActiveX, they're privately embedded in the executable.
    This project uses the same classes, but adds the frmDetector form that displays the events in a listbox.

    **WARNING**:
    When you start this project in the Visual Basic IDE, it will tell you that the two classes (DeviceDetector and DeviceInfo) have a public interface, which is not possible for a standalone executable. This happens because the two projects share the same source files. The Visual Basic IDE will change these properties as private. Don't save the project with these changes, or you'll have to restore them back to "Multiuse" for the ActiveX server project.
  
* AxDeviceDetectorTest.vbp

  * This is a standard EXE written in VB, acting as a test client using the ActiveX server (AxDeviceDetector.exe).

There is a [blog post that explains this project on my personal blog](https://francescofoti.com/2020/01/explaining-how-to-detect-device-arrival-removal-in-an-activex-server-in-real-time/).

## Runtime requirements

The deviceinfo.dll DLL needed in this project, was produced with Visual Studio 2017 (please see the deviceinfo_dll repository readme),  so it needs the presence of the corresponding [Visual C 2017 runtime](https://support.microsoft.com/fr-ch/help/2977003/the-latest-supported-visual-c-downloads) (x86) installed on the target computer to function properly.

You'll also need the msvbvm50.dll Visual Basic 5 runtime DLL (SP3).

The downloadable zip file contains both the required DLLs, that you should place either where the exe files are, or in one of your PATH directories.

## Downloadables

* binaries
  * [devicedetector_activex_and_standalone.zip](http://devinfo.net/downloads/axdevicedetector/devicedetector_activex_and_standalone.zip) (32bits version, MD5 sum: 23701e969c110eff5a25efdf0d8de194)
    Contains: 
    * AxDeviceDetector.exe : the ActiveX server
    * SaDeviceDetector.exe: the standalone demo
    * AxDeviceDetectorTest.exe: a standard exe used to test the ActiveX server
    * deviceinfo.dll (please see my deviceinfo_dll repository)
    * msvbvm50.dll: the Visual Basic 5 runtime
