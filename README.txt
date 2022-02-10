2/10/2022
Ben Campagnola
bcampagnola@fortherecord.com
For The Record




Hello! 
Welcome to the AZMC Crestron Powershell configuration guide.




OVERVIEW:

What information does this document cover?

1. Understanding how and why the AZMC Crestron programming and setup process is different
2. The files, structures, and file management
3. How to set up Powershell on your PC to run the script
4. How to operate the script and configure the spreadsheet
5. How to configure the Crestron processors
6. How to configure the Crestron panels
7. Troubleshooting





1. Understanding how and why the AZMC Crestron programming and setup process is different:

	To my knowledge, in all previous systems installed by FTR, each courtroom would get a unique copy of the Crestron code.
	This seems reasonable, as Crestron code generally uses details that are specific to each room.
	e.g. IP addresses of AV hardware, the number of cameras, the number of mics, etc.
		The Crestron processors need these varying details in order to function properly in each room, thus the unique Crestron programs also need to vary in this manner.

	We are contracted to install AV/recording systems in 230+ rooms for Maricopa County (AZMC). 
	Using the previous code deployment strategy, we would end this project with 230+ unique copies of the Crestron code.

	This is a problem on a number of fronts:
	- File management is cumbersome: 
		All of the manual repetition involved in creating and maintaining hundreds of copies of code poses a high likelihood of 
		user-error, and takes a long time to manage
	- Code becomes impossible to debug or update after day-1:
		Imagine that we just completed the project (230 rooms complete, with 230 active copies of the code), and we 
		discover a critical bug in the code. How do we resolve this?
		In this case, we have no choice but to open the 230 programs individually, correct the problem in each program, save each 
		as a new version number, and load all 230 processors.

		The same applies for times when the client has a request for additional features.


				230 unique copies of code
	[	Room.001 code w/ details specific to room 001 	]
	[	Room.002 code w/ details specific to room 002 	]
	[	Room.003 code w/ details specific to room 003 	]
	.....a few days later.....
	[	Room.230 code w/ details specific to room 230 	]


	Q: How do we resolve this problem?
		In Programming-Land, the solution is called "abstraction". 
		Somewhere down the list of secondary definitions:
			abstraction~ the process of considering something independently of its associations, attributes, or concrete 
			accompaniments

		Translated to English, and applied to our situation: 
		We want to build a version of courtroom Crestron code without all of these little details, like IP addresses and 
		quanitities of things.

	Q: But how do the rooms work if we don't include the room details in the code??
		We add the room details to a spreadsheet, and apply these details to each room via a scripting language.
		Thus, we then have ONE working copy of code, and a big spreadsheet containing data on the variances between each room.

			1x code file 						  1x spreadsheet
	[	All-in-one Crestron Code 	]   +   [	Room.001 details 	]
											[	Room.002 details	]
											[	Room.003 details	]
											......
											[	Room.230 details	]

	Crestron has a scripting language - Toolbox Script Manager. It is a bit of a relic, with much of it today being 
		non-functional. We can get it to work for us here, but the time has come to move on from the 1990s.
	We will instead use Windows PowerShell.


	Q: 1 Code file + 1 Spreadsheet? This is amazing!! Why have we not done this with EVERY Crestron program, EVER!?? 
		Well, as with everything, there's a trade-off.
		Configuring Crestron hardware and loading code is pretty simple stuff, right? 
		Well, we just made these processes about 4x more complex and technically demanding by introducing PowerShell into the mix.

		Once this methodology of code "abstraction" is in place, we can't go back to using Toolbox.
		Anyone who works on the AZMC systems will need to be trained on the new processes, including how to use PowerShell and
			the courtroom data spreadsheet. 

		With proper documentation (like this README!) and perhaps a training session or two, we will tame the beast and
			save ourselves a gajillion man-hours.

	Q: Is this new process just for the Maricopa County (AZMC) project?
		No. 
		- The new process has already been retroactively deployed to the 48 King County (KCCH) courtrooms.
		- It may be retroactively deployed to a few other existing clients.
		- And it will likely be deployed to all future projects, at least for the next 1-2 years, while we build an automated
			web service that manages all this stuff automatically.





2. The files, structures, and file management:
	
	General Notes 
		There are just a few files in this deployment package. The script includes a lot of automated, default file-naming
			type behavior, all of which is based on the idea that all of the files we are dealing with will be kept
			in the same directory. This isn't mandatory, but it is easy enough as there are so few files.

		In engineering this the first time through, I had a ton of issues storing all the files in a synchronized 
			OneDrive folder. It might be possible to make it work better, but it hardly seems worth the extra effort to me.
		So, to alleviate this issue without question, please pick a destination folder on your computer that is NOT part
			of a synchronized back-up service.
		If you are using github to clone (copy) and maintain the files we are dealing with, this is even more reason
			not to put them in a synchronized directory.


		Tangent#1: I absolutely hate how Windows defaults to putting everything in the  \OneDrive\Documents\ directory.
			If not because it makes everything difficult to keep track of, then because it makes the path names of its
			contained files too damn long!
			Anyway, if you want to fix this, I have what I believe is the best of many solutions. See notes at the bottom 
			of this document, under the header *** Down with Automatic OneDrive Storage ***.
			</tangent>


		Tangent#2: Are you git saavy?
		You probably should be. Github is the most widely used version control software (VCS), and it is incredibly useful
			when a group of people are working on the same set of files.
		As the name would depict, version control software is used for managing programs with many versions, and many 
			contributors.
		You seldom want to include version numbers in the names of the files that are managed through a VCS. This is 
			because the VCS manages the version numbers for you. All you need to know is "am I up-to-date with the group"?
		If we spot any files listed below that do not already have version numbers in the name - this is why!
		Also, because it is 2022, we can marvel at the awesome power of the VCS, and how it allows multiple people to
			simultaneously modify a given file without any fear of anyone's work getting overwritten.

		But, I realize this is already a heavy endeavor for some, so I'll leave it at this: if you want to learn more
			about github and VCS, I will post some online resources to the group. We could also do a training session 
			dedicated just to git.


	The files:
	AZMC_v3.02.01.lpz			-	the latest Crestron compiled code at the time of writing this document. This file 
									will be	installed into all of the processors in the Maricopa County courtrooms.

	AZMC_TSW1060_v3.02.01.vtz 	- 	like the .lpz file, there is a single .vtz file that gets installed into every touch
									panel in the project.   	

	AZMC_CourtroomData.csv 		- 	the spreadsheet containing all of the varying room details. Each courtroom gets
									one line in the spreadsheet. Any hardware devices that are specified in multiples
									will be listed in the same column, separated by a tilda (~).
									e.g. in our test setup in Boston, there are 4 fixed cameras, with node IP addresses
										of 20, 21, 22, and 23. In the spreadsheet, there is only a single "FixedCams" column,
										and in this case it is populated with the data: "20~21~22~23"

	CrestronConfig.ps1 			-	the PowerShell script. This file will automatically take the data in the spreadsheet,
									and configure the Crestron hardware with the appropriate data. The script will also
									do firmware, logging, generate reports, and will enable you to update code all 230 
									rooms with a single button-press. $$$








3. How to set up Powershell on your PC and run the script:
	Assuming you are using a Windows machine, you should have PowerShell installed already.
	If you are on Linux, then you're good enough with computers to find it yourself. =P

	Start by running PowerShell from the Start Menu (just type in "Power", and you should see it in the list).
	We do NOT need to run as Administrator.

	Left unchecked, PowerShell is a security liability, so by default it will block our first attempts at running our script.
	We need to send a command to allow the script to run.

	Once PowerShell has started, a new window opens up displaying a command prompt interface.
	[img]
	Copy-Paste the following into the prompt, and press [enter].

		Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass

	You should get a popup asking to confirm that you want to proceed. Just click 'Yes'.
	We will need to enter this command every time we want to run our script.
	(Once we have a public security certificate, this command will not be necessary, but that may take some time.)


	Lets change the working directory of the PowerShell prompt, pointing it to the directory where we have our files.
	The prompt should usually start in your Windows user folder.
	(Hint: if you know Linux Bash syntax, you know PowerShell prompt syntax)

		- On my machine, PowerShell prompt starts here:
					PS C:\Users\BenCampagnola>

	Use the 'cd' command to change directories.

		- Adding a single space after 'cd' followed by the name of a directory will add onto the current working directory.
		e.g.	   	PS C:\Users\BenCampagnola>  cd repos\AZMC[enter]
					...becomes...
					PS C:\Users\BenCampagnola\repos\AZMC>

		- Entering only a backslash will take us back to the root directory.
		e.g.		PS C:\Users\BenCampagnola>  cd \[enter]
					...becomes...
					PS C:\>

		- Adding a single space + [dot][dot] after 'cd' will bring us up one directory.
		e.g.		PS C:\Users\BenCampagnola>  cd ..[enter]
					...becomes...
					PS C:\Users>

	Once we have the PowerShell prompt pointed to the correct directory, simply run the script file by typing [dot][backslash],
		followed by the script name:
		e.g.		PS C:\Users\BenCampagnola\repos\AZMC>  .\CrestronConfig.ps1[enter]

	The script should start up!
	(Note: if you are having issues running the script, there are several related notes in the Troubleshooting section of this
		document)


	Most likely, if this is the first time you are running the script, it will attempt to automatically download and install
		a small software package, 'PSCrestron', which is Crestron's PowerShell module.
	Pay no attention to the snide remarks by Jonks, the passive-aggressive, self-aware PowerShell script AI,
		in the yellow text.
	He only means some of what he says.

	If all goes well, you will get a Windows installer pop-up. Follow the prompts to complete the installation of the software
		module.
	If you don't see the prompt, you will likely get some suggestions from Jonks as to what could be wrong.
	As he states, if you are connected to the internet but the script fails to install the Crestron PowerShell module, you can
		download and install it manually by browsing here:

		https://sdkcon78221.crestron.com/downloads/EDK/EDK_Setup_1.0.5.3.exe

	Ok, got the module installed?




4a. How to configure the spreadsheet:

	The spreadsheet is easy stuff.
	
	Each courtroom gets a single row on the spreadsheet.
	The table columns are labeled as such:
		- Ignore_Line
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
	'*' : Indicates columns that can take multiple values - one per device.

		Ignore_Line:	this is a data field used to determine which rooms have been loaded, and which have yet to be.
						In general, you shouldn't need to mess with this column. The script will read from and write to
						this column. It is important to understand how it functions though - 
						if the field is BLANK, as in when the script reads the data from this spreadsheet and it sees
						an "empty string", this flags the line as "HAS NOT BEEN LOADED".
						Once a room has been configured and loaded, the script will write a "1" into this column.

						This behavior only applies to broad, automated commands (to be covered later). If you specify
						that you want to load a particular room, and it has a '1' in the Ignore_Line field, it WILL still
						run.
	 	
		Room_Name:		this column refers to the alphanumeric code assigned to the AZMC courtrooms.
						e.g. "SCT23"
						It is important that this field NOT be loaded with an only-numeric value. It needs both letters
						and numbers.
						Q: Why?? 
							Because the Room_Name is one way we will be identifying specific rooms from the script.
							The other way is by the row number in the spreadsheet. If the Room_Name is only numeric,
							it might conflict with the list of row numbers.

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

	 	Processor_IP: 	this is a single-value column, as each courtroom will have exactly 1 Crestron processor.
		
	 	FileName_LPZ:	this column accepts a single file name, with extension.
	 					e.g.  AZMC_v3.02.01.lpz

		Panel_IP: 		this is a multi-value column, as some courtrooms will have 1 touch panel, and others will have 2.
						Separate multiples with a tilda (~)
						e.g.  125~126
						This value specifies that there are 2 panels in this room, and that their IP node addresses are
						125 and 126, respectively.
						(when added to the Subnet_Address of their room, their addresses would be something like:
						10.218.116.125
						10.218.116.126	)

		FileName_VTZ:	this column CAN accept multiple touch panel files. However, the expectation right now is that
						only one value will be entered here. I am expecting to load the same .vtz file to all panels. 
						In the case	that something unexpected happens, such as where the graphics fail to scale properly 
						on a smaller touch panel, we have the option of adding a 2nd .vtz file to this column in the 
						spreadsheet. Otherwise, the single .vtz file will be sent to multiple touch panels in some courtrooms.

		IP_DeviceType:	the remaining columns all have the convention "IP_DeviceType", indicating that these are the IP 
						addresses for the devices.
						Again, this is just the node IP address, meaning the part of the whole IP address other than
							the Subnet_Address.
						For any multiples, you know what to do! (but I'll say it anyway)
						Just add a tilda (~) between the values.
						e.g. For the "IP_Recorders" column, enter the Primary_Recorder (first) & Secondary_Recorder (2nd) 
							to this column, separated by a tilda (~).
						

	That's about it!
	Do not leave any stray text characters or comments anywhere on this document. The program is expecting it to be of a
		very specific size and shape.


4b. How to operate the script:

	The primary functions of the script are as follows:
		- read the data in the spreadsheet
		- automatically identify certain types of errors in the spreadsheet, such as bad or duplicate IP addresses
		- configure the Crestron equipment, based on the info in the spreadsheet
			- set Crestron devices to "authentication on", with standard FTR credentials
			- set Crestron devices to "SSL mode", forcing encryption on all socket connections
			- generate and send the IP tables for each of the Crestron processors and touch panels
			- send the processor and panel compiled code to the respective devices
			- send firmware to the Crestron hardware
			- generate reports of firmware versions
			- generate reports of any AV devices that are offline
			- generate reports of file-version comparisons
			- generate log files detailing which tasks were run, and whether they were successful or not

	Note that all of the above functions are optional.
	You can opt to do absolutely everything on every piece of hardware, or you can opt to run a single command for a 
		single device.

	Once the script is running, with the Crestron PowerShell software module installed on your machine, Jonks will present 
		you with a nested menu of options.

		1) Log Files & Reports
			a) Enable Logging
				1) Enable
				2) Disable
			b) Log Depth 
				1) Broad vs. Detailed 	- Broadly defined actions (e.g. Range based commands) vs. all granular actions 
				2) Debug On/Off 		- Try/catch error messages
				3) User Monitor On/Off 	- Tracks all user actions 
			c) Log File Name & Path
				1) Set File Name 		
				2) Set Path 	
				3) Set Full File Location 
			d) DateTime Format
				1) ms Epoch, year 2000
				2) YYYYMMDD HH:MM:SS.mmm
				3) HH:MM:SS.mmm YYYYMMDD
				4) Month DD, YYYY HH:MM:SS 
			e) Generate Reports 
				1) Firmware Versions
				2) File Versions
				3) AV Device Ping
				4) IP Table Offline Devices
				5) Processor Error Logs
			f) Filter & Manage Report Data
				1) Show All Entries
				2) Show Notices		- 	Show imperfect reports & errors
				3) Show Erroneous	- 	Show only errors, such as offline devices, processor plog errors, etc 
				4) Send To Log File 	
				5) Send To Report Files 

		2) Spreadsheet
			a) File Name, Path, & Type
				1) Set File Name 
				2) Set Path 
				3) Set Type 
				4) Set Full Location 
			b) Import File
				1) Import File Rite Meow
				2) Get Err From Most Recent Import
			c) Error Check Data
				1) Check For Duplicate IPs
				2) Check For Bad IPs
				3) Check For Duplicate RoomNames
				4) Check For Bad RoomNames - e.g. numeric-only RoomNames are not allowed
			d) Comment Mode
				1) Auto-Write Line Comments To File
				2) Any Value Blocks - If there is any value, group actions ignore this line. "" Emptystring is the only pass.
				3) 6-Bits per Device - (Set|Get) / (PreConfig, File & IPT, Firmware) / (Processor, Panel_01, [Panel_02]) 
			e) Get Data
				1) All
				2) Row
				3) Col
				4) RowCol 
			f) Set Data
				1) RowCol 
				2) All Column value 
			g) Expected Table Structure
				1) Show

		3) Device Config  
			a) Authentication
				1) Set On 
				2) Set Off
			b) SSL
				1) Set On 
				2) Set Off 
			c) Code Files
				1) Send If Behind 
			d) IP Table
				1) Print 
				2) Send 
			e) Firmware
				1) Send If Behind 

		4) Script Settings
			a) Clear Window On Menu Change
			b) List Split Character
			


	The bare minimum quick-start guide:
		- Import spreadsheet						(2,b,1)
		- Send code 								(3,c,1)
		- Send IP Tables							(3,d,2)

	If all the default settings are accurate (and they will be most of the time), then the above quick-start
		will be all that the installation team needs to do to properly configure one or more rooms.





5. How to configure the Crestron processors:
	
	Once the network settings have been set on the Crestron processors, I believe the rest of the configuration can be 
		completed through the script (if desired).
	These config bullet items include:
		1. Setting authentication to ON, using standard FTR credentials
		2. Setting SSL to ON
		3. Sending the .lpz program with the default IP Table (approx. half of the IP Table info is still built into 
			the program...)
		4. Setting the variable IP Table 	(...and the other half is populated via the spreadsheet & script)
		5. Loading firmware

	Note: The order of bulletpoints 3 & 4 is critical:
		The .lpz file needs to be loaded with the default IP table BEFORE the remaining IP Table configuration can take place.

	The complete list of bullet items 1-5 is actually in sequential order 
		per how I recommend the systems are commissioned.
	Enabling auth and SSL is not absolutely necessary right now, rather it is just the direction everything is moving, 
		so we might as well? I don't know. Maybe it will just slow things down. But it would be nice if all of the 
		Crestron devices were consistent in their configuration. 
	Similarly, firmware is probably not necessary, but just as long as we're getting in there, it would be great 
		to get done. This way we never need to get stuck loading firmware remotely.






6. How to configure the Crestron panels:





7. Troubleshooting:

	- General:
		I had a number of issues dealing with synchronized OneDrive folders causing the script to behave in odd ways.
		Ergo, please pick a directory on your hard drive that is only local to your machine.
		If you are using github to clone the AZMC repository, then you definitely do not want to be in a OneDrive folder.

	- Installing PSCrestron
		The PowerShell script should automatically download the installer file, but in case there are any issues with this,
			you can put this link into a browser, and manually download / install PSCrestron.

			https://sdkcon78221.crestron.com/downloads/EDK/EDK_Setup_1.0.5.3.exe

		If this link doesn't work for whatever reason, just try Googling "Crestron PowerShell module", and you should easily 
			find links to the correct page.
		All else fails, email me. 
		If you see any error messages popping up in the PowerShell window, or anywhere else, please attempt to copy them 
			into the email.
	- 




*** Down with Automatic OneDrive Storage ***

SO... you've made it this far. I can only assume that you share in my distaste for this rather gimmicky Windows 
behavior.

Check it out: 
The default file storage locations are kept in the registry.

As always, be very careful and deliberate when editing your registry.
Creating a system restore point is always a good idea before venturing down this road:

- Type 'regedit' in the search bar, and press [enter]
- Select 'yes' if you are prompted with a popup asking to enter the registry editor
- Once open, navgate to the following folders:
	HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\
- After clicking on the last folder in that path, "User Shell Folders", you should see a bunch of list items pop up on the
	right side of the window.
- Several of the values shown in the 'Data' column contain paths which include the text
	"\OneDrive - For The Record Group".
- Double-click on the corresponding items in the 'Name' column to open up a small control dialog window.
- Within these small dialog windows, you are free to edit the disk locations of your choice. 
- I simply removed the text "\OneDrive - For The Record Group" wherever it was found.
	e.g. Instead of  
		"C:\Users\BenCampagnola\OneDrive - For The Record Group\Documents\"
	...I modified the path to be
		"C:\Users\BenCampagnola\Documents\"

- Aside from it being a personal preference, these changes considerably improved the stability of the PowerShell script,
	and how it utilizes the PSCrestron module.

KaPow!



























