## Component Management Services focusing on Excel VB-Projects

- **Exports** any _Component_ the code has changed along with each Workbook save. 
- **Updates** any outdated _Used/Hosted&nbsp;[Common Component][9]. 
- **Synchronizes** the VB-Project of two Workbooks. 
> All services only require one component installed/imported plus a single code line for each service, still guaranteeing that a productive Workbook is not bothered by any of the configured services at all.

## Disambiguation

| Term             | Meaning                  |
|------------------|------------------------- |
| _Component_       | Generic term for a _VB-Project's_ _VBComponent_ (_Class Module_,  _Data Module_, _Standard Module_, or _UserForm_) |
| _[Common&nbsp;Component](#common-components)_ | A _Component_ providing services for a certain subject, dedicated for being used in any VB-Project<br>|
| _Procedure_           | Generic term for any _VB-Component's_ (public or private) `Property`, `Sub`, or `Function`|
|_Service_             | Generic term for any _Public Property_, _Public Sub_, or _Public Function_ of a _Component_ |
| _Servicing&nbsp;Workbook_ | The service providing Workbook, either the _[CompMan.xlsb][1]_ Workbook (when it is open) or the _CompMan Add-in_ when it is set up and open. |
| _Serviced&nbsp;Workbook_ | The Workbook prepared for being [serviced](#enabling-the-services-serviced-or-not-serviced).
|_VB&#8209;Project_    | Used synonymous with Workbook |
| _Sync&#8209;Source&#8209;Workbook_   | A _Workbook/VP-Project_ temporarily copied to the [CompMan's serviced root folder_](#configuration-changes) for being modified - and finally synchronized with its origin Workbook.|
| _Sync&#8209;Target&#8209;Workbook_ | A productive _Workbook/VP-Project_ temporarily moved to the configured [_Serviced Synchronization Target Folder_](#configuration-changes) for being synchronized with its corresponding _Sync-Source-Workbook_ when opened. |
| _Workbook&nbsp;parent&nbsp;folder_ | A folder dedicated to a _Workbook/VB-Project_. Note that an enabled Workbook is only [serviced](#enabling-the-services-serviced-or-not-serviced) when it is **exclusive** in its parent folder. Other Workbooks may be located in sub-folders however.|

## Services
### _Export Changed Components_
Used with the _Workbook\_BeforeSave_ event. Exports all _VB-Components_ of which the code has changed (i.e. differs from the recent export's Export-File), to the configured _Export Folder_ of which the name defaults to _source_. These _Export&nbsp;Files_ not only function as a code  backup in case Excel ends up with a destroyed VB-Project, which may happen every now and then - but only functions as a versioning means (e.g. when [GitHub][5] is used for instance). When a used or hosted _[Common Component](#common-components)_ has been modified and exported, a _[Pending Release](#pending-release-management)_ is registered  See also which are handled specifically.

### _Update Outdated Common Components_
Used with the _Workbook\_Open_ event all  _Used&nbsp;Common&nbsp;Components_ are checked whether they are outdated. In case, a dialog is displayed which allows to display the code difference (by means of ([WinMerge English][3], [WinMerge German][4], etc.) perform the update or skip it. The update uses the  _Export&nbsp;File_ of the _Raw&nbsp;Common&nbsp;Component_ in the _Common&nbsp;Components_ folder. 

### _Synchronize VB-Project_
The service allows a productive Workbook to remain in use while the VB-Project is developed, modified, maintained, etc., in a copy of it. When all changes had been done the VB-Project of the productive Workbook is synchronized. The benefit: A significant shorter downtime for the productive Workbook. 
> As with the _Export_ and _Update_ service, this service has no user interface! The service is invoked when the _Sync-Target-Workbook_ (i.e. the temporarily moved productive Workbook) is opened from within the configured [_ Serviced Synchronization Target Folder_](#configuration-changes). Provided the _Sync-Source-Workbook_ (i.e the copy of the productive Workbook) resides in the [_CompManServiced_ folder](#compmans-default-files-and-folders-environment).

## Installation
### Provision of CompMan as a servicing Workbook instance
> When [CompMan.xlsb][1] is downloaded to whichever location and opened it will setup its [default files and folder structure](#compmans-default-files-and-folders-environment) at the download location (don't worry, it may be moved afterwards). The setup completes with saving the Workbook to its dedicated parent folder and the downloaded Workbook is removed. The setup environment, i.e. the _CompManServiced_ folder may subsequently be moved to any location and even renamed.

1. [Download](#download-from-public-github-repo) and open the [CompMan.xlsb][1] Workbook  

2. When opened an explicit activation of the macros will be required, except when macros are unconditionally enabled - though not recommended by Microsoft

3. When WinMerge is not available/installed a corresponding message is displayed. The provided link may be used to download and install it. When continued without having it installed the message will be re-displayed whenever the [CompMan.xlsb][1] Workbook is opened.

4. Confirm CompMan's self _default environment_ setup at the location the Workbook is opened (see below).

> The CompMan services are now ready for being used by Workbooks which have the service(s) enabled (see below.

## Usage
### Enabling the services (serviced or not serviced)
A Workbook will only be [serviced](#enabling-the-services-serviced-or-not-serviced) by CompMan provided
1. A ***servicing CompMan instance*** (see [how to provide](#installation)) is open
2. The ***to-be-serviced Workbook*** Workbook has one or more of the below services enabled (see below)
3. The ***to-be-serviced Workbook*** is opened from within a sub-folder of the configured [_CompManServiced_ folder](#compmans-default-files-and-folders-environment), Note: In case of the _Synchronization service_ from within a sub-folder of the configured _Sync-Target-Folder_.
4. The ***to-be-serviced Workbook*** is the only Workbook in its parent folder (the parent folder may have sub-folders with Workbooks however)
5. WinMerge ([WinMerge English][3], [WinMerge German][4] or any other language version is installed to display the difference for any components when about to be updated by the [_Update_ service](#enabling-the-update-service)

>Note: As a consequence from the above, a productive Workbook must not be used from within the configured [_CompManServiced_ folder][8]. When a Workbook with any enabled/prepared service is opened when located elsewhere the user will not be bothered by any means, i.e. will not even recognize CompMan at all - even when open/available.

>Note: Even when a Workbook has one or more services [enabled](#enabling-the-services-serviced-or-not-serviced), the service is denied without notice when the above (pre)conditions are not met.

### Enabling the _Export_ service
The _Export_ service is performed whenever the Workbook is saved from within the configured [_CompManServiced_ folder](#compmans-default-files-and-folders-environment) and all the [preconditions](#enabling-the-services-serviced-or-not-serviced) are met.
1. From the Common Components folder import the _mCompManClient.bas_ (available after CompMan has been setup) which serves as the link to the CompMan services
2. Into the Workbook module copy the following:
```vb
Private Const HOSTED_RAWS = vbNullstring
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    mCompManClient.CompManService mCompManClient.SRVC_EXPORT_CHANGED, HOSTED_RAWS
End Sub
```
3. See _[Common Components](#common-components)_ in case the Workbook hosts one/some

### Enabling the _Update_ service
The _Update_ service is performed whenever the Workbook is opened from within the configured [_CompManServiced_ folder](#compmans-default-files-and-folders-environment) and all the [preconditions](#enabling-the-services-serviced-or-not-serviced) are meet.
1. Make sure CompMan is [provided](#installation) and open
1. From the _Common-Components_ folder import the _mCompManClient.bas_ (available after CompMan has been setup) which serves as the interface to the CompMan services.
2. Into the Workbook module copy the following:
```vb
Private Const HOSTED_RAWS = vbNullstring
Private Sub Workbook_Open()
    mCompManClient.CompManService mCompManClient.SRVC_UPDATE_OUTDATED, HOSTED_RAWS
End Sub
```
Despite the import of the _mCompManClient_ this is the only required modification in a VB-Project for this service.

3. In case the Workbook hosts one or more _[Common Components](#common-components]_ copy into the Workbook module:<br>
`Private Const HOSTED_RAWS = <component-name>[,<component-name]...`  
> This will only be the case when a Common Component ([in the sense of CompMan](#common-components)) is hosted, i.e. developed and maintained in a - preferably dedicated - Workbook. However, any Workbook may declare one of its components as a hosted _Common Component_ (watch-out conflicts!). When declared and modified the Export-File will be copied to the _Common Components Folder_ and the _Revision Number_ will be increased.

## Other
### Status of the CompMan Add-in
| Status             | Meaning |
|--------------------|---------|
| configured         | A valid, existing [_Add-in folder_](#compmans-default-files-and-folders-environment) is specified. CompMan's Add-in instance  may be setup/renewed  |
| not configured     | No [_Add-in folder_](#compmans-default-files-and-folders-environment) is specified. CompMan's Add-in instance cannot be setup/renewed (the button is not visible) |
| setup              | CompMan's Add-in instance is setup, i.e. available in the configured [_Add-in folder_](#compmans-default-files-and-folders-environment) |
| not setup          | CompMan's Add-in instance is not setup, i.e. not available in the configured [_Add-in folder_](#compmans-default-files-and-folders-environment). |
| paused             | CompMan's Add-in instance is indicated 'currently paused'. I.e. even when open it will  programmatic-ally be ignored |
| open               | CompMan's Add-in instance is open (it may be paused however) |
| not open           | CompMan's Add-in instance is not open. It will be opened when setup/renew and when Excel is started because the setup/renew establishes auto-open. |
| auto-open setup    | With setup/renew the auto-open has implicitly setup/established. |
| auto&#8209;open&nbsp;not&nbsp;setup| When the [_Add-in folder_](#compmans-default-files-and-folders-environment) is de-configured (no folder is selected when with 'Configure') a setup auto-open is removed |

#### Pending Release Management
When a used or hosted Common Component's code is modified and exported the component is registered as _Pending Release_, which is the release of the modification pending to become publicly available in the Common-Components folder. Any _Pending Release_ component is available in the Add-Ins menu from where the modification can be release to public one by one once the modification had become final.

#### The concept of "hosted" Common Components
Experience has shown than only a dedicated Workbook/VB-Project is appropriate for the development and especially the testing of a _Common Component_. It is required for the provision of a comprehensive test environment which also supports regression testing. _CompMan_ supports this concept by allowing to specify a _Common Component_ as being hosted in a Workbook. However, practice has shown that a modification or amendment  of a _Common Component_ is often triggered by a VB-Project just using, i.e. not hosting, it. _CompMan_ therefore supports this by keeping a record of which Workbook/VB-Project has last modified it.

#### The services
CompMan's initial intention was to keep _Common&nbspComponents_ up-to-date in all VB-Projects using them. To achieve this the _Export Service_ saves the Export-File of a modified used or hosted _Common Component to a _Common Components Folder_ thereby keeping a record of the modifying Workbook together with an incremented [_Revision Number_](#the-revision-number). Subsequently the _Update-Outdated-Common-Components_ service (by with the `Workbook_Open` event) checks for any outdated used or hosted _Common&nbsp;Components_ and offers an update in a dedicated dialog which allows to check the code difference by means of WinMerge ([WinMerge English][3], [WinMerge German][4].

### Download from public GitHub repo
It may appear pretty strange when downloading first from a public GitHub repo but its quite straight forward as the below image shows.  
![](assets/DownloadFromGitHubRepo.png)

## Contribution
Contribution of any kind is welcome, raising issues specifically.

[^1]: The terms _Raw_ and _Clone_ follow [_GitHub_][5] terminology
[^2]: That's the concept I follow for all my _Common Component_ (many of them used in [CompMan's VB-Project][1]

[1]:https://github.com/warbe-maker/VBA-Components-Management-Services/blob/master/CompMan.xlsb?raw=true
[2]:https://warbe-maker.github.io/2021/02/06/Programatically-updating-Excel-VBA-code.html
[3]:https://winmerge.org/downloads/?lang=en
[4]:https://winmerge.org/downloads/?lang=de
[5]:https://github.com
[6]:https://warbe-maker.github.io/vba/common/2021/02/19/Common-VBA-Components.html
[7]:https://warbe-maker.github.io/vba/excel/component/management/common/components/2023/06/12/VB-Project-development-towards-professionalism.html
[8]:https://github.com/warbe-maker/VBA-Components-Management/blob/master/SpecsAndUse.md#compmans-environment
[9]:https://github.com/warbe-maker/VBA-Component-Management/blob/master/SpecsAndUse.md#common-components