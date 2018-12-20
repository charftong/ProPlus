.SYNOPSIS
  Uninstalls Microsoft Office ProPlus as well as stand-alone products Visio and Project (supports 2007, 2010 and 2013 as well as both 32-bit or 64-bit installs).  
  (Optional) Can also Install Office Click-to-run.  

  .DESCRIPTION
  Uninstalls the Microsoft Office Products ProPlus as well as Visio and Project from a specified local source or share and using
  configuration XML files to configure the uninstallation process.  With the InstallOffice2016 switch, you can proceed to install Office 2016 click-to-run after the uninstall process completes.   
        
  .PARAMETER Products
  The Microsoft Office Products to uninstall. This must match the installation files for the product referred to in the SourcePath parameter.
	List of known values for Products
	Office12.ENTERPRISE - Office 2007 Enterprise
	Office12.PROPLUS - Office 2007 ProPlus
	Office14.PROPLUS - Office 2010 ProPlus
	Office14.PROPLUSR - Office 2010 ProPlus Trial
	Office14.STANDARD - Office Standard 2010
	Office15.PROPLUS - Office 2013 ProPlus
	Office15.STANDARD - Office Standard 2013
	Office16.PROPLUS - Office 2016 ProPlus	
	Office12.Visio - Visio Premium 2007
	Office14.Visio - Visio Premium 2010
	Office12.VisPro - Visio Professional 2007
	Office14.VisPro - Visio Professional 2010
	Office15.VisPro - Visio Professional 2013
	Office15.VisProR - Visio Professional 2013 Trial
	Office16.VisPro - Visio Professional 2016
	Office12.PrjPro - Project Professional 2007
	Office14.PrjPro - Project Professional 2010
	Office15.PrjPro - Project Professional 2013
	Office15.PrjProR - Project Professional 2013 Trial
	Office16.PrjPro - Project Professional 2016
  Default value is set to try and remove all known legacy products. 

  .PARAMETER UninstallSourcePath (optional)
  The location of the installation source files for the products to be removed. Can be a local or network path.   If not specified, the script will try to use the default installation location.  

  .PARAMETER WorkingDir (optional)
  The location of the config XML files that the script will create to control the product uninstallation as well as the script log files if a LoggingDir is not specified
  Defaults to %temp%

  .PARAMETER LoggingDir (optional)

  
  .PARAMETER InstallOffice (optional)
  If you pass this switch, the script will proceed to install Office Click-to-Run after the removal of the older versions of Office.  
  The default value is no

  .PARAMETER InstallSourcePath (optional)
  The location of the installation source files for Click-To-Run. Can be a local or network path.
  The parameter needs to be populated when using the "InstallOffice2016" switch. 

  .PARAMETER InstallConfigFile (optional)
  The file name of the configuration file needed for click-to-run install.   If not specified, it will default to 'configuration.xml'

  .PARAMETER SkipRemoval (optional)
  If you pass this switch, the script will skip over the Office Removal part of the script.   
  The default value is no

  .Requirements
   The account running this script will need the following permissions.
   local admin to the system that needs office uninstalled/installed
   At least NTFS 'read' access to the "UninstallSourcePath" location (if used)
   At least NTFS 'read' access to the "InstallSourcePath" location (if used)
   At least NTFS 'read/write' access to the "LoggingDir" location (if used)
   At least NTFS 'read/write' access to the "WorkingDir" location (if used)

  .TODO
	-Create a 'report' option to simply detect what the script will do prior to running it.
	-Check if uninstall XML Doc already exists in WorkingDir (if WorkingDir is not the default location)
	-add verbose log to uninstall xml file <Logging Type="Verbose" Path="%Temp% Template="Office_2016_InstallLog.txt" />
	-implement reboot after uninstall but before install

  .EXAMPLEs
   office-removal-deploy -Products 'Office15.PROPLUS' -WorkingDir 'c:\temp' -InstallOffice -ClickToRunSourcePath '\\server\share\office'
   office-removal-deploy -LoggingDir '\\server\share\office\logs' -InstallOffice -ClickToRunSourcePath '\\server\share\office' -SkipRemoval

  .NOTES
	Version:       6.0
	Author:        Joe Engelman
	Creation Date: Aug 3rd, 2016
	Purpose:       Remove Legacy Office installs(msi) and install Office Click-to-Run
	
