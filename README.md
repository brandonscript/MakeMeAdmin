# MakeMeAdmin
This tiny application will install a scheduled service on Windows 7 that perpetually promotes the specified user to administrative status, even when Group Policy overwrites the local admin users list.

MMA must be installed by an administrator (run Install.vbs as admin). From then on, any account listed in ObjectList.olx will gain adminsitrative rights every time the scheduled service runs.

## Build

You will need to set values for `{YOUR-DOMAIN}` and `{AUTHOR-ACCOUNT}` in `MainForm.cs` and `Install.vbs`. Then you should be able to compile and run in VS2012 or newer. Other than this, there is no documentation (nor will there be), nor is MMA warranted or guaranteed to compile or run on modern Windows operating systems.
