# USBPlugEvent

Usage: USBPlugEvent.exe [-l] [-d] [-i] [-r] "device_id" "cmd_to_launch"

Example: USBPlugEvent.exe -r "USB\VID_056D&PID_C07E\4687336F3936" "msg %username% the specified device was removed"

Options:

  -l  List all currently connected USB devices.  
  -d  Listen to Inserted/Removed USB devices.  
  -i  Launch the command on insert only.  
  -r  Launch the command on remove only.
  
