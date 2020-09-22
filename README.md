<div align="center">

## Disable Right Mouse Click in Webbrowser Control


</div>

### Description

Processes Mouse actions for Webbrowser Control enabling you to disable the action. This is an update due to requests from people that could not get it working. I have included a sample project to show it working.
 
### More Info
 
This example code places a HOOK into the message queue. When a new message is received it runs the Function to check the parameters of the the message. If the parameters are the same as what we are searching for, we can then choose to let it carry on or cancel the message. Full source code for the example is included and comments are shown where I can

Application must be exited by closing the form. the Form_unload event UNHOOKS the hook. Your computer may become unstable if you do not follow this, even in the Visual Basic IDE.


<span>             |<span>
---                |---
**Submitted On**   |1999-08-10 13:37:14
**By**             |[Chris Bint](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/chris-bint.md)
**Level**          |Unknown
**User Rating**    |4.8 (24 globes from 5 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[CODE\_UPLOAD314\.zip](https://github.com/Planet-Source-Code/chris-bint-disable-right-mouse-click-in-webbrowser-control__1-2897/archive/master.zip)








