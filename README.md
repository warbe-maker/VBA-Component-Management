## <u>Comp</u>onent <u>Man</u>agement Services for Excel VB-Projects

- **Exports** any _Component_ the code has changed along with each Workbook save. 
- **Updates** any outdated _[Common Component][5]_. 

<a id="raw-url" href="https://raw.githubusercontent.com/warbe-maker/VBA-Component-Management/master/CompMan/source/mCompManClient.bas">Download mCompManClient.bas</a>


https://raw.github.com/warbe-maker/VBA-Component-Management/master/CompMan/source/mCompManClient.bas


## Provision
Services are provided with an absolute minimum intervention in the serviced Workbook:
1. A <a href="https://github.com/warbe-maker/VBA-Component-Management/raw/refs/heads/master/CompMan/source/mCompManClient.bas?raw=true" download>mCompManClient</a> Standard Module imported as an interface to CompMan's services
2. One code line in the _Workbook\_Open_ event procedure for the update service
3. One code line in the _WorkBook\_After\_Save_ event procedure initiates the service provided the required conditions are meet (see [Serviced or not serviced Workbooks](#serviced-or-not-serviced)).> The services only require one component installed/imported, a single code line for each service, preventing that productive Workbooks are bothered by a configured service (see [serviced or not serviced][10]).

## The services
### At a glance
CompMan's initial intention was to keep _Common&nbspComponents_ up-to-date in all VB-Projects using them. To achieve this the _Export Service_ saves the Export-File of a modified used or [hosted][7] _[Common Component][5]_ to a _Common Components Folder_ thereby maintaining properties about the origin. Subsequently the _Update outdated Common-Components_ service (with the `Workbook_Open` event) checks for any outdated used or [hosted][7] _Common&nbsp;Components_ and offers an update in a dedicated dialog which allows to check the code difference by means of WinMerge ([WinMerge English][2], [WinMerge German][3].

### _Export Changed Components_ service
Used with the _Workbook\_AfterSave_ event. Exports all _VB-Components_ of which the code differs from the recent export's Export-File to the configured _[Export Folder][8]_.
>The _Export&nbsp;Files_ not only function as a code backup in case Excel ends leaving a destroyed/unreadable VB-Project behind. In combination with a synchronization service (e.g. [sync.com][6]) it also substantially supports versioning.

### Update outdated _Common Components_ service
Used with the _Workbook\_Open_ event all _used or [hosted][7]_  _[Common Component][5]_ are checked for being up-to-date and updated if not. This is supported by a dialog which allows to display the code difference (by means of ([WinMerge English][2], [WinMerge German][3], etc.), perform the update, or skip it. The update uses the  _Export&nbsp;File_ of the public _[Common Component][5]_ in the _[Common-Components folder][9]_.  
>The _Update_ service is performed only when the Workbook is opened from within the configured [_CompManServiced_ folder][4] and all the [preconditions][10]) are meet.

## Installation of CompMan as a Workbook/VBProject servicing instance
1. Download the <a href="https://github.com/warbe-maker/VBA-Component-Management/raw/refs/heads/master/CompMan.xlsb?raw=true" download>`CompMan.xlsb` Workbook</a>
2. Move the downloaded Workbook into a folder you will regard as the serviced root folder and open the Workbook.[^1] 
3. When WinMerge is not available/installed a corresponding message is displayed. The provided link may be used to download and install it. When continued without having installed it the message will be re-displayed whenever the <a href="https://github.com/warbe-maker/VBA-Component-Management/raw/refs/heads/master/CompMan.xlsb?raw=true" download>`CompMan.xlsb` Workbook</a> is opened.  
4. Confirm CompMan's self-setup. It will setup the [default files and folder environment][12].

> CompMan's services are now ready for being used by Workbooks which have the service(s) enabled (see below).

## Usage
A Workbook will only be serviced by CompMan provided:
1. A ***servicing CompMan instance*** (see [how to provide](#installation)) is open
2. The ***to-be-serviced Workbook*** Workbook has one or more of the below services enabled (see below)
3. The ***to-be-serviced Workbook*** is opened from within a sub-folder of the configured [_CompManServiced_ folder][4].
4. The ***to-be-serviced Workbook*** is the only Workbook in its parent folder (the parent folder may have sub-folders with Workbooks however)
5. WinMerge ([WinMerge English][2], [WinMerge German][3] or any other language version is installed to display the difference for any components when about to be updated by the [_Update_ service](#enabling-the-update-service)

> As a consequence from the above it may not be appropriate to use a Workbook productively from within the [_CompManServiced_ folder][4] but have it copied elsewhere outside for using it. The Workbook may be copied back at any time for developing/maintaining its VBProject. By the way: When a Workbook is opened outside the [_CompManServiced_ folder][4] the execution of enabled/prepared services will be denied without notice.

### Enabling the _Export_ service
The _Export_ service is performed whenever the Workbook is saved from within the configured _[CompMan serviced root folder][4] and all the [preconditions][10] are met.
1. Install: From the Common Components folder import the _mCompManClient.bas_ (available after CompMan has been setup) which serves as the link to the CompMan services.
2. Prepare: Into the Workbook module copy the following:  
 ```vb
Option Explicit
Private Const COMMON_COMPONENTS_HOSTED = vbNullstring

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    mCompManClient.CompManService mCompManClient.SRVC_EXPORT_CHANGED, COMMON_COMPONENTS_HOSTED
End Sub
```

### Enabling the _Update_ service
1. Install: From the Common Components folder import the _mCompManClient.bas_ (available after CompMan has been setup) which serves as the link to the CompMan services
2. Prepare: Into the Workbook module copy the following:
```vb
Option Explicit
Private Const COMMON_COMPONENTS_HOSTED = vbNullstring

Private Sub Workbook_Open()
    mCompManClient.CompManService mCompManClient.SRVC_UPDATE_OUTDATED, COMMON_COMPONENTS_HOSTED
End Sub
```
Despite the import of the _mCompManClient_ this is the only required modification in a VB-Project for this service.

3. [Hosted _Common Components_][11]: In case the Workbook hosts one or more _Common Components_, copy into the Workbook module:  
```vb
Private Const COMMON_COMPONENTS_HOSTED = <component-name>[,<component-name]...
```

## Other
### Download from public GitHub repo
It may appear pretty strange when downloading first from a public GitHub repo but its quite straight forward as the below image shows.  
![](assets/DownloadFromGitHubRepo.png)

## Contribution
Contribution of any kind is welcome, raising issues specifically.

[^1]:When opened, an explicit activation of the macros may be required. My proposal is to sign the VBProject and adjust the Macro security correspondingly (unconditional enabling is not recommended by Microsoft)

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
[11]:https://github.com/warbe-maker/VBA-Component-Management/blob/master/SpecsAndUse.md#hosted-versus-used-common-components
[12]:https://github.com/warbe-maker/VBA-Component-Management/blob/master/SpecsAndUse.md#compmans-environment

