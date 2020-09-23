	This zip file contains two programs. The server and the client piece. 
It also contains the file structure needed for both programs. The IDI folder goes
with the client piece and the DirectTransfer, RCV1, SAVE, and XMT1 folders go with
the server. To run extract zip to root directory to get proper structure. Next 
register the four dll's in the \Direct Transfer Client Pub\dependancy and put them 
in the system or system32 folder.
	These two programs serve as a file transfer system with two means of transfer
TCP\IP - Winsock and Email - Mapi. Program is currently set up to test both programs
on the same machine. to test across the network just change the IP Address.
	To use the email side you must setup a profile on each machine and provide 
the name in the code and then change the settings on the client to use email. This 
program is commercialy in use and all feed back is appreciated.

This will work from behind a proxy server if the port you are using is open.

Before transferring data the files are zipped in a winzip format but not using winzip
and then encrypted.

You can email your comments to peterb@informationdynamics.com
Pete Barth