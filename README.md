<h1>Deploy-Python-IISSite.ps1</h1>

As a matter of testing python-based websites in IIS, I have dealt with frustations due to the methods needed to manually deploy python and fastCGI in IIS. This script seeks to reliably remove those frustrations and barriers. It can be used via powershell to deploy or undeploy python-based IIS sites in just a few moments, and will do so reliably. Why python in IIS? Because sometimes you have to play the cards you are dealt.

<b>Disclaimer:</b> If you break your stuff it is on you. Microsoft stores IIS related config files inside the Windows\System32\inetsrv\config directory. For a webserver that has existed for so long, they haven't done an adequate job at supporting this service with sanity-checking and correction utilities.

****Testing Environment:</b> This powershell script has been tested in Windows Server 2022 using Python 3.11 (64-bit)

<b>Prerequisites:</b>

<ul>
<li>You must be running Windows (I haven't tested this on other Windows OS versions, but it should work in recent versions. The older your platform is the more likely this will require additional tweaking).</li>
<li>You must have sufficient windows privileges to manage IIS and IIS websites, to install windows features, and to download and install software.</li>
<li>You must already have IIS and some specific IIS Features installed. If you are missing specific features, it will prompt you to install them.</li>
<li>You must have already installed Python on your system.</li>
<li>You SHOULD have an idea of which python packages needed by the site you are deploying if you are porting code from another site or system. Those python package names need to be placed in the python_packages.json file. If you omit this file from the runtime directory, it will fallback to a default list of packages.</li>
</ul>

<b>Optional input parameters:</b>
<ul>
  <li>Mode
    <ul>
      <li>Deploy</li>
      <li>Undeploy</li>
      <li>Reset</li>
    </ul>
  </li>
  <li>SiteName</li>
  <li>SiteLocation</li>
  <li>AppPoolName</li>
  <li>Port</li>
  <li>KeepAppPool</li>
</ul>

<b>Input parameter use case/definitions:</b>
<ul>
  <li>Mode: [string](options: Deploy/Undeploy/Reset) Used to select add, remove, or reset(*) IIS behavior of script.</li>
  <li>SiteName: [string] Indicates site to configure/deconfigure. 'SiteName' denotes the value shown in IIS/Sites list, field: Name</li>
  <li>SiteLocation: [string] (Deploy mode only) Drive/Folder location to create the IIS site folder. Drive must alreeady exist. Runtime user must have sufficient permissions to create files and folders at that location.</li>
  <li>AppPoolName: [string] (Deploy mode only) The name of the Application Pool created/used in IIS to operate the site.</li>
  <li>Port: [int] (Deploy mode only) The port used to bind your site. Must not be already used in IIS.</li>
  <li>KeepAppPool: [bool] (Undeploy mode only) Defaults to "True". Can be used to auomate removal of the selected sites Application Pool from IIS, if the specified app pool is not used by other IIS sites.</li>
</ul>

During the python virtual environment cloning process, this script is designed to first attempt to pull python packages from the internet in a connected-device scenario. If the machine is not connected to the internet, it will attempt to clone the list of python packages from the local machine. This allows for a testing environment to remain disconnected, once the python package prerequisites have been loaded into your system python. If a package is listed in the python_packages.json that cannot be found, installation will fail.

(*) On runtime, the script attempts to backup the IIS applicationHost.config file from Windows. This is to ensure that, if any type of corruption occurs, the user has a backup of this file. Microsoft, in its infinite genious, fails to provide a great way to get back to "clean, working" if this file gets messed up. If something is messed up, as a result of this script or something else, it should provide info needed for you to manually fix that file. "Reset" mode rewrites the system applicationHost.config file with this backup. Realise, if anything has occurred to IIS or IIS sites that doesn't "match" the definitions... adverse impacts could occur.
