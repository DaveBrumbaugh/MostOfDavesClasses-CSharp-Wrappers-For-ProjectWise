This sample is designed to be loaded on file open. See SetPWVarsCE.cfg which I deploy to config/appl in MicroStation. I generally deploy the built DLL to the mdlapps folder. 

When the file is full open, the add-in checks to see if you are connected to ProjectWise and, if so, adds your ProjectWise document's general properties and custom attributes as configuration variables to your MicroStation CE session. Easy way to build ProjectWise-aware VBA's, etc.

I recently added a feature whereby the add-in lists all the associated references, the references' current path, and whether or not the reference is found. There's way more data you could log, but I thought that was sufficient for the moment. I could see Element ID for the reference attachment being useful. The log file ends up in c:\users\user.name\appdata\Roaming\Bentley\Logs\ as do all my logs.
