2/10/2022
Ben Campagnola
bcampagnola@fortherecord.com
For The Record




Hello! 
Welcome to the AZMC Crestron Powershell configuration guide.




OVERVIEW:

What information does this document cover?
1. Understanding how and why the AZMC Crestron programming and setup process is different
2. How to set up Powershell on your PC to run the script
3. How to operate the script and spreadsheet
4. How to configure the Crestron processors
5. Troubleshooting





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







