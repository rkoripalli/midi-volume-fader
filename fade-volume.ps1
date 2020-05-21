Add-Type -TypeDefinition @'
using System.Runtime.InteropServices;

[Guid("5CDF2C82-841E-4546-9722-0CF74078229A"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IAudioEndpointVolume {
  // f(), g(), ... are unused COM method slots. Define these if you care
  int f(); int g(); int h(); int i();
  int SetMasterVolumeLevelScalar(float fLevel, System.Guid pguidEventContext);
  int j();
  int GetMasterVolumeLevelScalar(out float pfLevel);
  int k(); int l(); int m(); int n();
  int SetMute([MarshalAs(UnmanagedType.Bool)] bool bMute, System.Guid pguidEventContext);
  int GetMute(out bool pbMute);
}
[Guid("D666063F-1587-4E43-81F1-B948E807363F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IMMDevice {
  int Activate(ref System.Guid id, int clsCtx, int activationParams, out IAudioEndpointVolume aev);
}
[Guid("A95664D2-9614-4F35-A746-DE8DB63617E6"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IMMDeviceEnumerator {
  int f(); // Unused
  int GetDefaultAudioEndpoint(int dataFlow, int role, out IMMDevice endpoint);
}
[ComImport, Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")] class MMDeviceEnumeratorComObject { }

public class Audio {
  static IAudioEndpointVolume Vol() {
    var enumerator = new MMDeviceEnumeratorComObject() as IMMDeviceEnumerator;
    IMMDevice dev = null;
    Marshal.ThrowExceptionForHR(enumerator.GetDefaultAudioEndpoint(/*eRender*/ 0, /*eMultimedia*/ 1, out dev));
    IAudioEndpointVolume epv = null;
    var epvid = typeof(IAudioEndpointVolume).GUID;
    Marshal.ThrowExceptionForHR(dev.Activate(ref epvid, /*CLSCTX_ALL*/ 23, 0, out epv));
    return epv;
  }
  public static float Volume {
    get {float v = -1; Marshal.ThrowExceptionForHR(Vol().GetMasterVolumeLevelScalar(out v)); return v;}
    set {Marshal.ThrowExceptionForHR(Vol().SetMasterVolumeLevelScalar(value, System.Guid.Empty));}
  }
  public static bool Mute {
    get { bool mute; Marshal.ThrowExceptionForHR(Vol().GetMute(out mute)); return mute; }
    set { Marshal.ThrowExceptionForHR(Vol().SetMute(value, System.Guid.Empty)); }
  }
}
'@

Import-Module -Force "$PSScriptRoot\PeteBrown.PowerShellMidi.dll"

$inputDevices = Get-MidiInputDeviceInformation
foreach ($device in $inputDevices)
{
  if($device.Name -match 'loopMIDI Port'){
    Write-Host $device.Name  
    Write-Host $device.Id
    $deviceId = $device.Id
  }
}

function FadeUp {
  Param ($desiredVolume)
  Write-Host $desiredVolume
  [audio]::Mute = $false
  [audio]::Volume = 0
  while($vol -lt $desiredVolume*.01){
    [audio]::Volume  = $vol+.01
    $vol=$vol+.01
    Start-Sleep -Milliseconds 100
  }  
}

function FadeDown {
  $vol = [audio]::Volume
  while($vol -gt .01){
    [audio]::Volume  = $vol-.01
    $vol=$vol-.01
    Start-Sleep -Milliseconds 100
  }
  [audio]::Mute = $True  
}


[PeteBrown.PowerShellMidi.MidiInputPort]$inputPort = Get-MidiInputPort -id $deviceId
# set this to false if you don't want the input port to translate a zero velocity
# note on message into a note off message
$inputPort.TranslateZeroVelocityNoteOnMessage = $true

$inputPort | Get-Member -Type Event

# this lists the events available
#Write-Host "Events available for MidiInputPort ------------------------------------------- "
#$inputPort | Get-Member -Type Event

# this is just an identifier for our own use. It can be anything, but must
# be unique for each event subscription.
$noteOnSourceId = "NoteOnMessageReceivedID"
$noteOffSourceId = "NoteOffMessageReceivedID"
$controlChangeSourceId = "ControlChangeMessageReceivedID"

# remove the event if we are running this more than once in the same session.
# Write-Host "Unregistering existing event handlers ----------------------------------------- "
Unregister-Event -SourceIdentifier $noteOnSourceId -ErrorAction SilentlyContinue
Unregister-Event -SourceIdentifier $noteOffSourceId -ErrorAction SilentlyContinue
Unregister-Event -SourceIdentifier $controlChangeSourceId -ErrorAction SilentlyContinue

#register for the event
# Write-Host "Registering for input port .net object events --------------------------------- "

# Note-on event handler script block
$HandleNoteOn = 
{
	Write-Host " "
	Write-Host "Powershell: MIDI Note-on message received" -ForegroundColor Cyan -BackgroundColor Black
	Write-Host "  Channel: " -NoNewline -ForegroundColor DarkGray; Write-Host $event.sourceEventArgs.Channel  -ForegroundColor Red
	Write-Host "  Note: " -NoNewline -ForegroundColor DarkGray; Write-Host $event.sourceEventArgs.Note  -ForegroundColor Red
  Write-Host "  Velocity: " -NoNewline -ForegroundColor DarkGray; Write-Host $event.sourceEventArgs.Velocity  -ForegroundColor Red
  if($event.sourceEventArgs.Note -like 3){
    Write-Host "Fading Up"
    FadeUp $event.sourceEventArgs.Velocity
  }
}

# Note-off event handler script block
$HandleNoteOff = 
{
	Write-Host " "
	Write-Host "Powershell: MIDI Note-off message received"  -ForegroundColor DarkCyan  -BackgroundColor Black
	Write-Host "  Channel: " -NoNewline -ForegroundColor DarkGray; Write-Host $event.sourceEventArgs.Channel  -ForegroundColor Red
	Write-Host "  Note: " -NoNewline -ForegroundColor DarkGray; Write-Host $event.sourceEventArgs.Note  -ForegroundColor Red
  Write-Host "  Velocity: " -NoNewline -ForegroundColor DarkGray; Write-Host $event.sourceEventArgs.Velocity  -ForegroundColor Red
  if($event.sourceEventArgs.Note -like 3){
    Write-Host "Fading Down"
    FadeDown
  }
}

# Control Change event handler script block
$HandleControlChange = 
{
	Write-Host " "
	Write-Host "Powershell: Control Change message received"  -ForegroundColor Green -BackgroundColor Black
	Write-Host "  Channel: " -NoNewline -ForegroundColor DarkGray; Write-Host $event.sourceEventArgs.Channel  -ForegroundColor Red
	Write-Host "  Controller: " -NoNewline -ForegroundColor DarkGray; Write-Host $event.sourceEventArgs.Controller  -ForegroundColor Red
	Write-Host "  Value: " -NoNewline -ForegroundColor DarkGray; Write-Host $event.sourceEventArgs.Value  -ForegroundColor Red
}

$job1 = Register-ObjectEvent -InputObject $inputPort -EventName NoteOnMessageReceived -SourceIdentifier $noteOnSourceId -Action $HandleNoteOn
$job2 = Register-ObjectEvent -InputObject $inputPort -EventName NoteOffMessageReceived -SourceIdentifier $noteOffSourceId -Action $HandleNoteOff
$job3 = Register-ObjectEvent -InputObject $inputPort -EventName ControlChangeMessageReceived -SourceIdentifier $controlChangeSourceId -Action $HandleControlChange