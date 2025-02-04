# WARP Refresh

Contains scheduled task to refresh Cloudflare WARP auto-connect timeout.

Currently **Windows** only, since the `warp-refresh.vbs` is written in VBScript to be run as a
Windows Scheduled Task. (MacOS users are welcome to try converting the script and re-purposing it
into a similarly runnable package.)

*If you are a developer on Windows dev machine with Cloudflare WARP installed, and somewhat bothered
by occasional interruptions to video calls, or strange TLS cert error without realizing WARP
reconnected, or having to disconnect WARP every now and then, this might be for you.*

## How to use

Download [`warp-refresh.vbs`](warp-refresh.vbs) into `%UserProfile%\Documents\warp-refresh.vbs`, and
run the following in `cmd`:

```bat
schtasks /Create /SC minute /MO 60 /TN "WARP auto-connect refresh" /TR "WScript.exe \"%UserProfile%\Documents\warp-refresh.vbs\""
```

Note that the task is designated to run with your current user and does not require escalated
privileges. You may choose to change the file path to download into, but remember to change argument
value for `/TR`.

If you change your mind after adding as a scheduled task, you may delete the task by running:

```bat
schtasks /Delete /TN "WARP auto-connect refresh" /F
```

### `cmd`-only way to install quickly

`WIN+R` and run `cmd`. Enter the following:

```bat
:: Downloads the .vbs file into designated file path
:: NOTE: this line is to be directly run in cmd (to run powershell on behalf for convenience), and not directly in powershell
powershell -c "Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/dsaidgovsg/warp-refresh/main/warp-refresh.vbs' -OutFile '%UserProfile%\\Documents\\warp-refresh.vbs'

:: Installs and runs script as a scheduled task
schtasks /Create /SC minute /MO 60 /TN "WARP auto-connect refresh" /TR "WScript.exe \"%UserProfile%\Documents\warp-refresh.vbs\""
```

## How the script works

The script assumes WARP is installed, which exposes `warp-cli` in `cmd`. The script also assumes to
be run periodically (every 60 mins), and requires the user to have first manually
disconnected the WARP (i.e. slider in the WARP tray icon turned off).

Previously in commit `d6ae6a2`, the script assumes warp-cli can extract the auto-connect timeout
value. This CLI subcommand has been removed, and now it insteads only checks if WARP is currently
disconnected.

It runs the connect command (equivalent to the slider in WARP tray icon turning on), and immediately
runs disconnect command again (equivalent to the slider in WARP tray icon turning off). The effect
of this resets the auto-connect timeout back to 10800 seconds (empirically discovered).

Instead of printing to stdout, log echos are written to `warp-refresh.log` in the same directory to
where the script resides (the log filename also changes to whatever the script filename is). This is
because VBScript does not support printing to stdout. The logs are kept to a max of 100 lines to
prevent any possibility of causing storage space to run out.

Again, the user is expected to first manually disconnect the WARP when the OS starts up, if one
chooses to do so, since the script is intended to only have effect in this state, and not effect
any changes should the user prefers the WARP to stay connected.

While this is not 100% foolproof to prevent your video streams from cutting off in the process of
connect/disconnect, empirically this seems to work well enough.

## Why `.vbs`?

`.vbs` was chosen because it can be directly run as a scheduled task. While `.bat` file can also
work, it will cause a `cmd` window to pop up and disappear every time it triggers, which is
annoying.

Importantly, this also prevents any focus loss to the current working window, otherwise we would
just be transferring the annoyance of having to manually disconnect WARP, to that of having to deal
with the popup window and the loss of window focus.
