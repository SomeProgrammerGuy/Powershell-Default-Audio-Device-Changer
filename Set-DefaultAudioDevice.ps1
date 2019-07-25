##################################################################################
#                                                                                #
# Powershell Default Audio Device Changer                                        #
# ---------------------------------------                                        #
#                                                                                #
# Change the default audio device in windows 7 with a single PowerShell script.  # 
# This does not depend on any additional components.                             #
#                                                                                #
# Date Created : 08 March 2019                                                   #
# Date Modified: 08 March 2019                                                   #
#                                                                                #
# https://github.com/WHumphreys/Powershell-Default-Audio-Device-Changer          #
#                                                                                #
# Copyright (c) William Humphreys. all rights reserved.                          #
#                                                                                #
# MIT License                                                                    #
#                                                                                #
# Permission is hereby granted, free of charge, to any person obtaining a copy   #
# of this software and associated documentation files (the "Software"), to deal  #
# in the Software without restriction, including without limitation the rights   #
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell      #
# copies of the Software, and to permit persons to whom the Software is          #
# furnished to do so, subject to the following conditions:                       #
#                                                                                #
# The above copyright notice and this permission notice shall be included in all #
# copies or substantial portions of the Software.                                #
#                                                                                #
# THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR     #
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,       #
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE    #
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER         #
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,  #
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE  #
# SOFTWARE.                                                                      #
#                                                                                #
##################################################################################

$cSharpSourceCode = @"
using System;
using System.Runtime.InteropServices;

///////////////////////////////////////////////////////////////////////////////
//                                                                           //
//  Interface identifiers.                                                   //
//  NOTE: This script is designed for Windows 7, these Interface identifiers //
//        may be different for other Windows versions.                       //
//        (Google is your friend...)                                         //
//                                                                           //
///////////////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////////////
//                                                                          //
// The ERole enumeration defines constants that indicate the role           //
// that the system has assigned to an audio endpoint device.                //
//                                                                          //
// eConsole:         Games, system notification sounds, and voice commands. //
// eMultimedia:	     Music, movies, narration, and live music recording.    //
// eCommunications:  Voice communications (talking to another person).      //
// ERole_enum_count: The number of members in the ERole enumeration         //
//                   (not counting the ERole_enum_count member).            //
//                                                                          //
//////////////////////////////////////////////////////////////////////////////

public enum ERole : uint
{
    eConsole         = 0,
    eMultimedia      = 1,
    eCommunications  = 2,
    ERole_enum_count = 3
}

/////////////////////////////////////////////////////////
//                                                     //
// Undocumented COM-interface IPolicyConfig.           //
// Use to set default audio Capture / Render endpoint. //
//                                                     //
/////////////////////////////////////////////////////////

[Guid("F8679F50-850A-41CF-9C72-430F290290C8"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IPolicyConfig
{
    // HRESULT GetMixFormat(PCWSTR, WAVEFORMATEX **);
    [PreserveSig]
    int GetMixFormat();
    
    // HRESULT STDMETHODCALLTYPE GetDeviceFormat(PCWSTR, INT, WAVEFORMATEX **);
    [PreserveSig]
	int GetDeviceFormat();
	
    // HRESULT STDMETHODCALLTYPE ResetDeviceFormat(PCWSTR);
    [PreserveSig]
    int ResetDeviceFormat();
    
    // HRESULT STDMETHODCALLTYPE SetDeviceFormat(PCWSTR, WAVEFORMATEX *, WAVEFORMATEX *);
    [PreserveSig]
    int SetDeviceFormat();
	
    // HRESULT STDMETHODCALLTYPE GetProcessingPeriod(PCWSTR, INT, PINT64, PINT64);
    [PreserveSig]
    int GetProcessingPeriod();
	
    // HRESULT STDMETHODCALLTYPE SetProcessingPeriod(PCWSTR, PINT64);
    [PreserveSig]
    int SetProcessingPeriod();
	
    // HRESULT STDMETHODCALLTYPE GetShareMode(PCWSTR, struct DeviceShareMode *);
    [PreserveSig]
    int GetShareMode();
	
    // HRESULT STDMETHODCALLTYPE SetShareMode(PCWSTR, struct DeviceShareMode *);
    [PreserveSig]
    int SetShareMode();
	 
    // HRESULT STDMETHODCALLTYPE GetPropertyValue(PCWSTR, const PROPERTYKEY &, PROPVARIANT *);
    [PreserveSig]
    int GetPropertyValue();
	
    // HRESULT STDMETHODCALLTYPE SetPropertyValue(PCWSTR, const PROPERTYKEY &, PROPVARIANT *);
    [PreserveSig]
    int SetPropertyValue();
	
    // HRESULT STDMETHODCALLTYPE SetDefaultEndpoint(__in PCWSTR wszDeviceId, __in ERole role);
    [PreserveSig]
    int SetDefaultEndpoint(
        [In] [MarshalAs(UnmanagedType.LPWStr)] string wszDeviceId, 
        [In] [MarshalAs(UnmanagedType.U4)] ERole role);
	
    // HRESULT STDMETHODCALLTYPE SetEndpointVisibility(PCWSTR, INT);
    [PreserveSig]
	int SetEndpointVisibility();
}

[ComImport, Guid("870AF99C-171D-4F9E-AF0D-E63DF40C2BC9")]
internal class _CPolicyConfigClient
{
}

public class PolicyConfigClient
{
    public static int SetDefaultDevice(string deviceID)
    {
        IPolicyConfig _policyConfigClient = (new _CPolicyConfigClient() as IPolicyConfig);

	try
        {
            Marshal.ThrowExceptionForHR(_policyConfigClient.SetDefaultEndpoint(deviceID, ERole.eConsole));
		    Marshal.ThrowExceptionForHR(_policyConfigClient.SetDefaultEndpoint(deviceID, ERole.eMultimedia));
		    Marshal.ThrowExceptionForHR(_policyConfigClient.SetDefaultEndpoint(deviceID, ERole.eCommunications));
		    return 0;
        }
        catch
        {
            return 1;
        }
    }
}
"@

add-type -TypeDefinition $cSharpSourceCode

function Get-CaptureDevices 
{
    Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Capture\*\Properties\" | 
    Select @{Name="CaptureDeviceId";Expression={$_.PSParentPath.substring($_.PSParentPath.length - 38, 38)}}, 
           @{Name="CaptureDeviceName";Expression={$_."{a45c254e-df1c-4efd-8020-67d146a850e0},2"}}, 
           @{Name="CaptureDeviceInterface";Expression={$_."{b3f8fa53-0004-438e-9003-51a46e139bfc},6"}} | 
    Format-List
}

function Get-CaptureDeviceId 
{
    Param
    (
        [parameter(Mandatory=$true)]
        [string[]]
        $captureDeviceName,
        [parameter(Mandatory=$true)]
        [string[]]
        $captureDeviceInterface
    )

    Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\capture\*\Properties\" |
    Where {($_."{a45c254e-df1c-4efd-8020-67d146a850e0},2" -eq $captureDeviceName) -and ($_."{b3f8fa53-0004-438e-9003-51a46e139bfc},6" -eq $captureDeviceInterface)} |
    ForEach {$_.PSParentPath.substring($_.PSParentPath.length - 38, 38)}
}

function Get-RenderDevices 
{
    Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render\*\Properties\" | 
    Select @{Name="CaptureDeviceId";Expression={$_.PSParentPath.substring($_.PSParentPath.length - 38, 38)}}, 
           @{Name="CaptureDeviceName";Expression={$_."{a45c254e-df1c-4efd-8020-67d146a850e0},2"}}, 
           @{Name="CaptureDeviceInterface";Expression={$_."{b3f8fa53-0004-438e-9003-51a46e139bfc},6"}} | 
    Format-List
}

function Get-RenderDeviceId 
{
    Param
    (
        [parameter(Mandatory=$true)]
        [string[]]
        $renderDeviceName,
        [parameter(Mandatory=$true)]
        [string[]]
        $renderDeviceInterface
    )

    Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render\*\Properties\" |
    Where {($_."{a45c254e-df1c-4efd-8020-67d146a850e0},2" -eq $renderDeviceName) -and ($_."{b3f8fa53-0004-438e-9003-51a46e139bfc},6" -eq $renderDeviceInterface)} |
    ForEach {$_.PSParentPath.substring($_.PSParentPath.length - 38, 38)}
}

function Set-DefaultAudioDevice
{
    Param
    (
        [parameter(Mandatory=$true)]
        [string[]]
        $deviceId
    )

    If ([PolicyConfigClient]::SetDefaultDevice("{0.0.0.00000000}.$deviceId") -eq 0)
    {
        Write-Host "SUCCESS: The default audio device has been set."
    }
    Else
    {
        Write-Host "ERROR: There has been a problem setting the default audio device."
    }
}

# Usage

# A list of all the currently installed audio device can be obtained using:

#   Get-CaptureDevices

#   or

#   Get-RenderDevices

# Set-DefaultAudioDevice requires a DeviceId which can be obtained using: 

#   Get-CaptureDeviceId "[DEVICENAME]" "[DEVICEINTERFACE]"
	
#   or
	
#   Get-RenderDeviceId "[DEVICENAME]" "[DEVICEINTERFACE]"

# Set the default audio device using:

#   Set-DefaultAudioDevice (Get-CaptureDeviceId "[DEVICENAME]" "[DEVICEINTERFACE]") 

#   or

#   Set-DefaultAudioDevice (Get-RenderDeviceId "[DEVICENAME]" "[DEVICEINTERFACE]")


