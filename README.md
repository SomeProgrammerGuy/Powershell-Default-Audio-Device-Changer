# Powershell Script to Change the Default Audio (In/Out) Device in Windows 7

## Description

For reasons only known to themselves Microsoft have supplied no official API or programmatic way of setting the default audio device in Windows.
 
NOTE: this script has been written for Windows 7 (and possibly 8.x). It can be easily altered for windows 10 by changing the various COM ID's. Also be in mind that the "non-official" way this script accesses the audio devices is in no way supported by Microsoft who could change the underlying software at any time.
 
The scope of this script is to be able to change the default audio device without the use of any additional third-party files. This is so changes can be made quicker and easier for Administrators (rather than programmers) without the need for a compiler. It also makes it a little clearer as to how it's been done.

I do recommend if you want to control the audio devices more comprehensively then one of the third-party software solutions will serve you better.

This script has been developed with PowerShell v6.2.0 and hasn’t been tested on any other version.

https://github.com/PowerShell/PowerShell/releases/

## Usage

A list of all the currently installed audio device can be obtained using:

```
Get-CaptureDevices
```

or

```
Get-RenderDevices
```

Set-DefaultAudioDevice requires a DeviceId which can be obtained using: 

```
Get-CaptureDeviceId "[DEVICENAME]" "[DEVICEINTERFACE]"
```

or

```	
Get-RenderDeviceId "[DEVICENAME]" "[DEVICEINTERFACE]"
```

Set the default audio device using:

```
Set-DefaultAudioDevice (Get-CaptureDeviceId "[DEVICENAME]" "[DEVICEINTERFACE]")
```

or

```	
Set-DefaultAudioDevice (Get-RenderDeviceId "[DEVICENAME]" "[DEVICEINTERFACE]")
```

## License
Copyright (c) William Humphreys. all rights reserved.

Licensed under the [MIT License](./LICENSE).