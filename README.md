## Component Management Services focusing on Excel VB-Projects

- **Exports** any _Component_ the code has changed along with each Workbook save. 
- **Updates** any outdated _[Common Component][5]_. 
> The services only require one component installed/imported, a single code line for each service, preventing that productive Workbooks are bothered by a configured service (see [serviced or not serviced][10]).

## The services
### At a glance
CompMan's initial intention was to keep _Common&nbspComponents_ up-to-date in all VB-Projects using them. To achieve this the _Export Service_ saves the Export-File of a modified used or [hosted][7] _[Common Component][5]_ to a _Common Components Folder_ thereby maintaining properties about the origin. Subsequently the _Update outdated Common-Components_ service (with the `Workbook_Open` event) checks for any outdated used or [hosted][7] _Common&nbsp;Components_ and offers an update in a dedicated dialog which allows to check the code difference by means of WinMerge ([WinMerge English][2], [WinMerge German][3].

### _Export Changed Components_ service
Used with the _Workbook\_AfterSave_ event. Exports all _VB-Components_ of which the code differs from the recent export's Export-File to the configured _[Export Folder][8]_.
>The _Export&nbsp;Files_ not only function as a code backup in case Excel ends leaving a destroyed/unreadable VB-Project behind. In combination with a synchronization service (e.g. [sync.com][6]) it also substantially supports versioning.

### Update outdated _Common Components_ service
Used with the _Workbook\_Open_ event all _used or [hosted][7]_  _[Common Component][5]_ are checked for being up-to-date and updated if not. This is supported by a dialog which allows to display the code difference (by means of ([WinMerge English][2], [WinMerge German][3], etc.), perform the update, or skip it. The update uses the  _Export&nbsp;File_ of the public _[Common Component][5]_ in the _[Common-Components folder][9]_.  
>The _Update_ service is performed only when the Workbook is opened from within the configured [_CompManServiced_ folder][4] and all the [preconditions][10]) are meet.

## Installation
### Provision of CompMan as a servicing Workbook instance
> When [CompMan.xlsb][1] is downloaded to whichever location and opened it will setup its [default files and folder structure](#compmans-default-files-and-folders-environment) at the download location (don't worry, it may be moved afterwards). The setup completes with saving the Workbook to its dedicated parent folder and the downloaded Workbook is removed. The setup environment, i.e. the _CompManServiced_ folder may subsequently be moved to any location and even renamed.

1. [Download](#download-from-public-github-repo) and open the [CompMan.xlsb][1] Workbook  
2. When opened an explicit activation of the macros will be required, except when macros are unconditionally enabled - though not recommended by Microsoft  
3. When WinMerge is not available/installed a corresponding message is displayed. The provided link may be used to download and install it. When continued without having it installed the message will be re-displayed whenever the [CompMan.xlsb][1] Workbook is opened.  
4. Confirm CompMan's self _default environment_ setup at the location the Workbook is opened (see below).

> The CompMan services are now ready for being used by Workbooks which have the service(s) enabled (see below.

## Usage
A Workbook will only be serviced by CompMan provided:
1. A ***servicing CompMan instance*** (see [how to provide](#installation)) is open
2. The ***to-be-serviced Workbook*** Workbook has one or more of the below services enabled (see below)
3. The ***to-be-serviced Workbook*** is opened from within a sub-folder of the configured [_CompManServiced_ folder](#compmans-default-files-and-folders-environment), Note: In case of the _Synchronization service_ from within a sub-folder of the configured _Sync-Target-Folder_.
4. The ***to-be-serviced Workbook*** is the only Workbook in its parent folder (the parent folder may have sub-folders with Workbooks however)
5. WinMerge ([WinMerge English][2], [WinMerge German][3] or any other language version is installed to display the difference for any components when about to be updated by the [_Update_ service](#enabling-the-update-service)

>Note: As a consequence from the above, a productive Workbook must not be used from within the configured [_CompManServiced_ folder][4]. When a Workbook with any enabled/prepared service is opened when located elsewhere the user will not be bothered by any means, i.e. will not even recognize CompMan at all - even when open/available.

>Note: Even when a Workbook has the export an/or the update service [enabled](#usage), the service will be denied without notice when the above (pre)conditions are not met.

### Enabling the _Export_ service
The _Export_ service is performed whenever the Workbook is saved from within the configured _[CompMan serviced root folder][4] and all the [preconditions][10] are met.
1. Install: From the Common Components folder import the _mCompManClient.bas_ (available after CompMan has been setup) which serves as the link to the CompMan services.
2. Prepare: Into the Workbook module copy the following:  
 ```vba
 Private Const COMMON_COMPONENTS_HOSTED = vbNullstring<br>Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)<br>
    mCompManClient.CompManService mCompManClient.SRVC_EXPORT_CHANGED, COMMON_COMPONENTS_HOSTED
End Sub
```

### Enabling the _Update_ service
1. Install: From the Common Components folder import the _mCompManClient.bas_ (available after CompMan has been setup) which serves as the link to the CompMan services
2. Prepare: Into the Workbook module copy the following:
```vb
Private Const HOSTED_RAWS = vbNullstring
Private Sub Workbook_Open()
    mCompManClient.CompManService mCompManClient.SRVC_UPDATE_OUTDATED, HOSTED_RAWS
End Sub
```
Despite the import of the _mCompManClient_ this is the only required modification in a VB-Project for this service.

3. Hosted _Common Components_: In case the Workbook hosts one or more _[Common Components](#common-components]_ copy into the Workbook module:  
```vb
Private Const COMMON_COMPONENTS_HOSTED = <component-name>[,<component-name]...
```

## Other
### Download from public GitHub repo
It may appear pretty strange when downloading first from a public GitHub repo but its quite straight forward as the below image shows.  
![](assets/DownloadFromGitHubRepo.png)

## Contribution
Contribution of any kind is welcome, raising issues specifically.


[1]:https://github.com/warbe-maker/VBA-Components-Management-Services/blob/master/CompMan.xlsb?raw=true
[2]:https://winmerge.org/downloads/?lang=en
[3]:https://winmerge.org/downloads/?lang=de
[4]:https://github.com/warbe-maker/VBA-Component-Management/blob/master/SpecsAndUse.md#compman-serviced-root-folder
[5]:https://github.com/warbe-maker/VBA-Component-Management/blob/master/SpecsAndUse.md#common-components
[6]:https://sync.com
[7]:https://github.com/warbe-maker/VBA-Component-Management/blob/master/SpecsAndUse.md#hosted-versus-used-common-components
[8]:https://github.com/warbe-maker/VBA-Component-Management/blob/master/SpecsAndUse.md#export-folder
[9]:https://github.com/warbe-maker/VBA-Component-Management/blob/master/SpecsAndUse.md#common-components-folder
[10]:https://github.com/warbe-maker/VBA-Component-Management/blob/master/SpecsAndUse.md#serviced-or-not-serviced