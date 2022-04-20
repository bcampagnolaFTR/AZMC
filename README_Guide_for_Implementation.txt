2/10/2022
Ben Campagnola
bcampagnola@fortherecord.com
For The Record




Hello! 

Welcome to the AZMC Crestron configuration guide for the Implementation team.
This guide is written in the first person. Deal with it.

In this doc, *** indicates advanced informaion that is NOT critical for everyone to learn. 



The TL:DR;

- We now use a single set of Crestron code files for all similar courtrooms.
	For AZMC specifically, there are 2 sets of code: one for rooms with Cynap devices, and one for those without Cynap.
- The implementation team now needs to do the following additional tasks:
	Build the IP table for each processor manually.
	(whereas previously, the IP table was built by Toolbox by checking the "Send Default IP Table" option when loading an LPZ file)

	To speed up (and hopefully simplify) the process, we have created a PowerShell script to manage the configuration of
		the Crestron hardware, including building the IP tables.

	It's probably a good idea to know what the script is doing in the background, and to know how to execute the configurations
		manually, in case the script is having a bad day. These instructions will be included at the end of this document. 



GENERAL:

Q: 	"Is this new implementation process just for the AZMC project?"
A: 	No. The new process will be used for all future projects with more than a few courtrooms (lets say for now, all projects with
		more than ~4 courtrooms will use this new process)




THE FILES:

The files you need:
	(note: the version numbers listed here are just examples. You will probably have more recent version numbers of each file)

	.lpz
	AZMC_v3.02.01.lpz			-	the latest Crestron compiled code at the time of writing this document.
									This file will be installed into all of the processors in the Maricopa County courtrooms.

	.vtz
	AZMC_TSW1060_v3.02.01.vtz 	- 	like the .lpz file, there is a single .vtz file that gets installed into all touch
									panels in the project.   	

	.csv
	AZMC_CourtroomData.csv 		- 	the spreadsheet containing the unique IP address data for each room, as well as the 
									quantity of certain device types. (cameras, DSP units, etc)

	.ps1
	CrestronConfig.ps1 			-	the PowerShell script. 

	.bat
	CrestronConfig.bat 			- 	a file that needs to accompany the .ps1 file.



POWERSHELL:

How to set up Powershell on your PC:
	- 	Use a machine running Windows. You already have PowerShell. :)
	- 	Run the CrestronConfig.bat
		The CrestronConfig.ps1 file should be in the same directory as the .bat file
		*** The .bat runs the .ps1 file, but does so with some extra parameters that allow the script to run without being 
			blocked by Windows security.

	- 	At this point, if you do not already have the PSCrestron software installed, the script will automatically attempt 
		to download and run the installer. A pop-up window should appear with a software installation wizard. Manually run
		the wizard.
	- 	If the PSCrestron software fails to download or if the wizard does not run, you can manually download 
		and run the required file:
		https://sdkcon78221.crestron.com/downloads/EDK/EDK_Setup_1.0.5.3.exe
	-	Once this is complete, run the .bat file again.

	- Once the script has started, a new window opens up displaying a command prompt interface.
	- Options are:
		1. load a spreadsheet
		2. run processor configuration
		3. run panel configuration
	
		







CONFIGURING THE SPREADSHEET:

	Each courtroom gets a single row on the spreadsheet.
	The table columns are labeled as such:
		- Room_Name
		- Facility_Name
		- Subnet_Address
		- Processor_IP
		- FileName_LPZ
		* Panel_IP
		* FileName_VTZ
		- IP_ReporterWebSvc
		- IP_WyrestormCtrl
		* IP_FixedCams
		* IP_DSPs
		* IP_Recorders
		- IP_DVDPlayer
		* IP_AudicueGW
		* IP_PTZCams

		Note that some of the items listed are prefixed with a '-', and some with a '*'. 
		'-' : Indicates columns that expect a single value.
		'*' : Indicates columns that can take multiple values - one value per device.

	 	
		Room_Name:		this column refers to the alphanumeric code assigned to the AZMC courtrooms.
						e.g. "SCT23"
						It is important that this has a value, and that the value uses both letters and numbers.


		Facility_Name: 	the name of the facility that the courtroom is in. How about that.

		Subnet_Address:	this is the portion of the IP addresses that ALL devices in a single room have in common.
						e.g. if ALL devices in a room start with  '192.168.11.'  then this is the subnet address.
						(this is technically not the definition of a subnet address, but it is good enough for right now.
						Ask me if you require a more accurate explanation)
						
						Note: some rooms will have devices that span across different "3rd octets" (this is the number
							between the 2nd and 3rd [dot]).
						e.g. Lets say room XYZ_35 has 5 fixed cameras. Their full IP addresses are:
							10.218.114.252
							10.218.114.253
							10.218.114.254
							10.218.115.1
							10.218.115.2
							Here, we are spanning different 3rd octets- 114 & 115. This room has IP addresses using both of these.
						** FOR ANY ROOMS WHERE THIS IS THE CASE:
							Roll the Subnet_Address back 1 octet, and add the 3rd octet to ALL of the device "node addresses".
						e.g.
							Subnet_Address = 10.218.

							Processor_IP = 114.250
							Panel_IP = 114.251
							FixedCam01 = 114.252
							FixedCam02 = 114.253
							FixedCam03 = 114.254
							FixedCam04 = 115.1
							FixedCam05 = 115.2

						Clear as mud? Good.


						Alternatively, you can leave the subnet_address field blank, and just populate the full IP address
						into each device's appropriate column.

	 	Processor_IP: 	this is a single-value column, as each courtroom will have exactly 1 Crestron processor.
		
	 	FileName_LPZ:	this column accepts a single file name, with extension.
	 					e.g.  AZMC_v3.02.01.lpz

		Panel_IP: 		this is a multi-value column, as some courtrooms will have 1 touch panel, and others will have 2.
						Separate multiples with a tilda (~)
						e.g.  125~126
						This value specifies that there are 2 panels in this room, and that their IP addresses end in
						125 and 126, respectively.

		FileName_VTZ:	this column CAN accept multiple touch panel files. However, the expectation right now is that
						only one value will be entered here. I am expecting to load the same .vtz file to all panels. 
						In the case	that something unexpected happens, such as where the graphics fail to scale properly 
						on a smaller touch panel, we have the option of adding a 2nd .vtz file to this column in the 
						spreadsheet. Otherwise, the single .vtz file will be sent to multiple touch panels in the courtrooms
						with 2 panels.

		IP_DeviceType:	the remaining columns all have the convention "IP_DeviceType", indicating that these are the IP 
						addresses for the devices.
						Again, this is just the node IP address, meaning the part of the whole IP address that is not
							the Subnet_Address.
						For any multiples, you know what to do! (but I'll say it anyway)
						Just add a tilda (~) between the values.
						e.g. For the "IP_Recorders" column, enter the Primary_Recorder (first) & Secondary_Recorder (2nd) 
							to this column, separated by a tilda (~).
						

	That's about it!
	Do not leave any stray text characters or comments anywhere in the spreadsheet. This will throw off the script.









SYSTEM CONFIGURATION:

	In the previous versions of the AZMC code, and in many other FTR-installed systems, there is a folder in the SIMPL Windows 
	program titled "System Configuration". This folder contained the logic programming that specified many of the room-specific 
	details, such as the total number of microphones, how the displays were controlled, video streaming & routing settings, 
	local device names, etc.

	Fast forward to now - all of the programming formerly found in the "System Configuration" folder is now bulit into
	the touch panel interface. We can call it the "config page". 
	Ok fine, I'll call it the "config page", and you can call it whatever you want. Be that way.

	By pressing and for 5 seconds holding the ForTheRecord logo at the top-center of the interface, we are magically 
	whisked away to the Config Page. There are 9 "Subpages" built into the Config Page, each of which can be selected
	from the list on the left. The subpages are:
		- Comms
		- Cams
		- Devices
		- HDMI Inputs
		- Video
		- DSP
		- Mutes
		- Mute Groups
		- Rooms

	I'll cover the 1st subpage here for now. 

	Subpage01, Comms: 	- 	These controls are for the small group of IP address & MAC address information that can not be
							built into the spreadsheet & script.

							The data that is entered into the Config Page and subpages is stored into NonVolatile RAM on 
							the processor, which means that when there's a power outage, or we need to reset the program, 
							it will still remember the settings we put in when it comes back online.


							On the top of the Config Page are 2x important buttons:
							- Default Settings
							- Restart Program

							Procedure:
							After entering the appropriate data into the Comms subpage (4 different parameters), you should
							should press each of the 3 "commit-->" buttons on the page, from top to bottom. Then, press the
							Restart Program button once. 

							Note: This is the only time we will need to restart the program. The other 8 pages of settings
							are all applied immediately to the program any time the values change, including checkboxes,
							numeric + \ - controls, and the text string input boxes.


							When the program restarts, you can try out the "Default Settings" button.
							Press the button while you are on subpages 2 thru 9, and the settings for that page will be
							populated with what Jaime believes to be the "most common" settings across the different rooms.
							If you hold the button down for > 5seconds, and then release the button, all controls on 
							subpages 2-9 will be set to "Default" configuration.

							^^ This is an excellent starting point for almost all rooms. 


							Once you have gone through the subpages 2-9 and confirmed the settings to match actual room
							configuration, you can exit the Config Page... 


							...and then we're all done! 
							Does it work?? 


							Q: Why is the DSP IP_Address entered in the Config Page, but it also shows up in the spreadsheet?
							A: The SIMPL Windows program does actually use the IP Table sockets to connect to the DSP,
								but only for an online-status indicator. The DSP device control is sent via C# module.



























