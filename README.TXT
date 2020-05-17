/*****************************************************\
*             ModPlugin ActiveX Control               *
*                Samuel Gomes (Blade)                 *
*            mailto: v_2samg@hotmail.com              *
*                                                     *
*               - Distribute freely -                 *
*                                                     *
*                     ModPlugin                       *
*                 Olivier Lapicque                    *
*           mailto: olivierl@modplug.com              *
*                                                     *
*         ModPlugin UI graphics by MisterX            *
*            All other graphics by Blade              *
\*****************************************************/


1. Requirements
2. Installation
3. Uninstallation
4. Features
5. Troubleshooting
6. FAQ
7. Version History

-------------------------------------------------------
1. Requirements
-------------------------------------------------------

DX2-66 or higher
Windows 95 or NT 4.0.
A soundcard with 32-bit drivers
ModPlugin (download from http://www.modplug.com)
Visual Basic 5.0 or higher (only tested in VB, try VC,
Delphi etc. at you own risk)

NOTE:
44.1KHz sound quality is not recommended for computers
slower than a 486DX4-100.

ModPlugin ActiveX Control is a public domain utility
and may be freely distributed. I am not responsible for
any problems that could be caused by the usage of this
plug-in.
THIS PROGRAM IS ONLY INTENDED FOR NON-COMMERCIAL USE.

Right-clicking works on the control as it is supposed to.

ModPlugin ActiveX Control is just a wrapper around
ModPlugin, so it supports all the features ModPlugin
supports.

Ok, so I messed up this text file the first time. So
here it is the complete and informative Read Me. :)

-------------------------------------------------------
2. Installation
-------------------------------------------------------

First of all, it is *really* important to install
ModPlugin. Download the latest EXE version of
ModPlugin from http://www.modplug.com and install it.
Then find the file npmod32.dll in you system and copy
it to your System (Win9x/Me) or System32 (NT/2K/XP)
folder. ModPlugin v1.91.05 npmod32.dll is supplied
with this package. You may use that too.

Put the ModPlugin32.ocx file in your System/System32
folder or wherever. Go to VB's components dialog.
Find ModPlugin32.ocx and load it. VB will register it
for you. Or, use regsvr32.exe to resister
ModPlugin32.ocx.

-------------------------------------------------------
3. Uninstallation
-------------------------------------------------------

Use regsvr32.exe with the /u switch to unregister
ModPlugin32.ocx.

Then delete it from your system.

Delete npmod32.dll from your System/System32 folder
if you want.

-------------------------------------------------------
4. Features
-------------------------------------------------------

ModPlugin ActiveX Control supports all the features
supported by ModPlugin. The right-clicking works. And
you users can setup ModPlugin from the context menus.

Full ModPlugin GUI is implemented with all features.
Take a look at the example. :)

Four events, namely Loaded, Played, VolumeChanged and
Status are implemented. All ModPlugin functions works.
Even IsPlaying (Playing), IsReady (Ready) etc.

Changing of some properties are limited to
design-time. Because of the way ModPlugin works,
changing these properties at run-time requires the
playback to be stopped.

You will also find a control called PlaySound. You
may use this control to playback simple sounds. See
the example to learn how to use it. It's very simple.

-------------------------------------------------------
5. Troubleshooting
-------------------------------------------------------

Visual Basic says it cannot find npmod32.dll. What gives?
Obviously, you have not read the full Read Me! Please
read the 'Installation' part of this document. Copy
npmod32.dll to your System/System32 folder.

The ModPlugin ActiveX Control says that it cannot
setup itself and that my administrator needs to copy
the latest version of ModPlugin (npmod32.dll) to the
"System/System32" folder. What's going on?
This message has been put in the ActiveX on purpose.
It appears the either your account does not have
sufficient access right for ModPlugin ActiveX Control
to copy the lastest version of npmod32.dll to the
Windows "System/System32" folder, or, the version of
npmod32.dll present in your system is lower than
version 1.91. It is important to know that the
ModPlugin ActiveX Control was developed for version
1.91+ of ModPlugin. In either case see the topic
"Installation" and do the needfull. :)

-------------------------------------------------------
6. FAQ
-------------------------------------------------------

I cannot use the control from Visual C++, Delphi...!
As I already said, this control was tested only with
Visual Basic and may or may not work with other
development tools. Try at your own risk. Let me know
if you come up with something.

There is already a ModPlug VB ActiveX downloadable
from ModPlug Central. Why yet another ActiveX?
The ActiveX which is currently available from
ModPlug Central is an ActiveX DLL. Therefore to use
that ActiveX you have to write code and use it's
'Plugin' class. Also, you will not get the ModPlugin
User Interface (spectrum analyzer, VU-meter, seek bar
etc.) with that ActiveX. ModPlug ActiveX Control on
the other hand in an ActiveX *Control* which supports
the full ModPlugin UI and much more. If you don't
want the UI simply set the control's 'Visible'
property to 'False' and use it.

Why is the Status event fired a little later if I
use the Start/Stop button on the ModPlugin UI?
ModPlugin was designed to be used as a plug-in for
Web Browsers, hence the API it provides is very
limited. Therefore to support all the four events in
ModPlugin ActiveX Control I had to use a timer. To
make the events more zappy change the UpdateInterval
property of the control to a value less than the
default 500. 500 is 1/2 a second, 250 is 1/4 a
second and so on. Hence, for 500 the events are
fired 2 times a second, for 250 - 4 times a second
and so on.

How do I know what does each method/property/event do?
Simply open Object Browser in VB and select the
ModPlugin32 library. Now select any method/property/
event and look at the description below. You should
find a small help text.

Any tips?
Sure! Use the Position property with care. Don't
just map Position to the value property of a scroll
bar in the scroll event. It can cause some nasty
things... even program crash.

I want feature X or feature Y to be changed.
Please, e-mail me if you have any questions,
feedback or suggesions on the ModPlugin ActiveX
Control. mailto: v_2samg@hotmail.com

-------------------------------------------------------
7. Version History
-------------------------------------------------------

1.00.93 - Fixed most bugs; release version.
1.00.75 - Added new properties to control mixer.
1.00.55 - Updated project to be binary compatible.
1.00.44 - Compiled with Visual Basic 6.0 SP6
1.00.37 - Updated setup procedure in Main().
1.00.36 - Fixed Version property. Updated About dialog.
1.00.35 - More small changes and fixes (very stable).
1.00.31 - Rewrote this text file. :)
1.00.30 - Updated and corrected control help texts.
1.00.27 - More minor cosmetic changes.
1.00.23 - Made control more design mode friendly.
1.00.20 - Added Status and VolumeChanged events.
1.00.15 - Added Loaded and Played events.
1.00.00 - Initial make.

-------------------------------------------------------
EOF
-------------------------------------------------------
