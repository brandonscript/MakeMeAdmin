# MakeMeAdmin
This tiny application will install a scheduled service on Windows 7 that perpetually promotes the specified user to administrative status, even when Group Policy overwrites the local admin users list.

There is no documentation (nor will there be), nor is it warranted or guaranteed to work on modern Windows operating systems.

## Installing

Must be installed by an administrator (run Install.vbs as admin). From then on, any account listed in ObjectList.olx will gain adminsitrative rights every time the scheduled service runs.
