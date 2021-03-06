﻿LDAPSyncOutlook
===============
Sync server LDAP contacts with Windows contacts and Microsoft Outlook persistently. OpenLDAP, Active Directory, etc. 

Windows version compatible
===============
Windows XP, Windows Vista, Windows 7, Windows 8.

Mail clients
===============
The clients are Outlook Express, Windows Mail, Windows Live mail and any client that uses the Windows contacts.

Office version compatible
===============
Microsoft Outlook 2000, Microsoft Outlook XP, Microsoft Outlook 2003, Microsoft Outlook 2007, Microsoft Outlook 2010, Microsoft Outlook 2013.

Installer
===============
https://github.com/fquiroz01/LDAPSyncOutlook/blob/master/installer/LDAPSyncOutlookInstaller-3.2.exe?raw=true

Configuration
===============
To open the configuration of software is necessary open the main program, if only the Microsoft Outlook is installed, 
<br/>The configuration for Office 2003 and lower is in the main menu <i>Tools -> Options</i>, in the Tab LDAPSyncOutlook.
<br/>The configuration for Office 2007 and upper is in the main menu <i>Options -> Add-Ins</i>, select in the list LDAPSyncOutlook and click in options.<br/>
For the Outlook Express and the similars the contacts are stored in contacts folder.

| Field | Description |
| ------------- | ----------- |
| Nombre | Name of service for example my addressbook |
| Servidor |  Name or IP for server with LDAP |
| Base | base of search of LDAP tree, for example dc=mydomain,dc=com or dc=mydomain,dc=local |
| Directorio | Microsoft Outlook's Folder where the contacts are stored |
| <b>Avanzado</b> |
| Puerto | Port the LDAP server. |
| Mi servidor requiere autenticacion | Activate when the LDAP server requires validation |
| Usuario | Name of user to validate for example usuario@mydomain.local or cn=usuario,dc=mydomain,dc=com |
| Contraseña | Password for the user |
| Actualizar cada | Update time interval in minutes |
