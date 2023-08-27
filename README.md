## Component Management Services focusing on Excel VB-Projects

> The services **Export** (any _Component_ the code has changed), **Update** (any outdated _Used&nbsp;[Common Components](#common-components)_), and **Synchronize** (the VB-Project of two Workbooks) only requires one component installed/imported with a single code line for each service, by guaranteeing that a productive Workbook is not bothered by these services at all.

## Disambiguation

| Term             | Meaning                  |
|------------------|------------------------- |
| _Component_       | Generic term for a _VB-Project's_ _VBComponent_ (_Class Module_,  _Data Module_, _Standard Module_, or _UserForm_) |
| _Common&nbsp;Component_ | 1. A _Component_ which potentially may be used in any Excel VB-Project<br> 2. A _Component_ [hosted](#the-concept-of-hosted) in an (optionally dedicated) Workbook which claims it as the _Raw_ component. This Workbook hosts the source of the component for being development, maintained, and tested, while other Workbooks/VB-Projects are using a _Clone_ of the component. [^1]   |
|_Used&nbsp;Common&nbsp;Component_ | The copy of a _Raw&#8209;Component_ in a _Workbook/VP&#8209;Project_ using it. _Clone-Components_ may be automatically kept up-to-date by the _UpdateOutdatedCommonComponents_ service.  The term _clone_ is borrowed from GitHub but has a slightly different meaning because the clone is usually not maintained but the _raw_ |
| _Procedure_           | Generic term for any _VB-Component's_ (public or private) _Property_, _Sub_, or _Function_ |
| _Raw&nbsp;Common&nbsp;Component_ | The instance of a _Common&nbsp;Component_ which is regarded the developed, maintained and tested 'original', [hosted](#the-concept-of-hosted) in a dedicated _Raw&#8209;Host_ Workbook. The term _raw_ is borrowed from GitHub and indicates the original version of something |
| _Raw&#8209;Host_      | The Workbook/_VB-Project_ which hosts the _Raw-Component_ |
|_Service_             | Generic term for any _Public Property_, _Public Sub_, or _Public Function_ of a _Component_ |
| _Servicing&nbsp;Workbook_ | The service providing Workbook, either the _[CompMan.xlsb][1]_ Workbook (when it is open) or the _CompMan Add-in_ when it is set up and open. |
| _Serviced&nbsp;Workbook_ | The Workbook prepared for being [serviced](#enabling-the-services-serviced-or-not-serviced).
|_VB&#8209;Project_    | Used synonymous with Workbook |
| _Sync&#8209;Source&#8209;Workbook_   | A _Workbook/VP-Project_ temporarily copied to the [CompMan's serviced root folder_](#configuration-changes) for being modified - and finally synchronized with its origin Workbook.|
| _Sync&#8209;Target&#8209;Workbook_ | A productive _Workbook/VP-Project_ temporarily moved to the configured [_Serviced Synchronization Target Folder_](#configuration-changes) for being synchronized with its corresponding _Sync-Source-Workbook_ when opened. |
| _Workbook&nbsp;parent&nbsp;folder_ | A folder dedicated to a _Workbook/VB-Project_. Note that an enabled Workbook is only [serviced](#enabling-the-services-serviced-or-not-serviced) when it is **exclusive** in its parent folder. Other Workbooks may be located in sub-folders however.|

## Services
### _Export Changed Components_
Used with the _Workbook\_BeforeSave_ event, all _VB-Components_ code is compared with its previous _Export&nbsp;File_. When the code had changed the component is re-exported to the configured _Export Folder_ of which the name defaults to _source_. These _Export&nbsp;Files_ not only function as a code  backup in case Excel ends up with a destroyed VB-Project, which may happen every now and then - but only functions as a versioning means (e.g. when [GitHub][5] is used for instance). See also _[Common Components](#common-components)_ which are handled specifically.

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

>Note: As a consequence from the above, a productive Workbook must not be used from within the configured [_CompManServiced_ folder](#compmans-default-files-and-folders-environment). When a Workbook with any enabled/prepared service is opened when located elsewhere the user will not be bothered by any means, i.e. will not even recognize CompMan at all - even when open/available.

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

### Enabling the _Synchronization_ service
The _Synchronize VB-Project_ service is performed when the Workbook is opened from within the configured [_Serviced Synchronization Target Folder_](#configuration-changes) and all the [preconditions](#enabling-the-services-serviced-or-not-serviced) are meet
1. Make sure CompMan is [provided](#installation) and open
1. From the _Common Components_ folder import the _mCompManClient.bas_ which serves as the interface for the CompMan services.
2. In the Workbook module copy the following:
```vb
Private Sub Workbook_Open()
    '~~ The statement is only required for a Workbook which may be synchronized.
    '~~ However, the statement has no effect unless the Workbook is opened from within the configured 'Serviced Synchronization Target Folder'.
    mCompManClient.CompManService mCompManClient.SRVC_SYNCHRONIZE
End Sub
```

### Using the _Synchronization Service_
#### Steps
1. Follow the [Installation](#installation) instructions and the [Enabling the _Synchronization_ service](#enabling-the-synchronization-service) instructions. 
2. Copy the productive Workbook into a dedicated folder under the configured [_CompManServiced_ folder](#compmans-default-files-and-folders-environment) folder
3. Modify the VB-Project as intended while the productive Workbook remains in use!.
4. When the development/modification has been finished, ****close the Workbook!**** and proceed with the next steps.
5. Move the productive Workbook to the configured [_Serviced Synchronization Target Folder_](#configuration-changes) and open it. In case this folder has yet not been configured, switch to the open [CompMan.xlsb][1] and use the displayed _Config_ Worksheet.
6. Open the moved Workbook and follow the synchronization steps. The synchronization will be done on a working copy (name with a suffix _Synced_).
7. When the synchronization has finished, save the working copy as the new productive Workbook - e.g. by dropping the _Synced_ suffix from the name and moving it back to the "production location" it originates.
8. When everything has finally turned out perfect the remaining Workbook from step 3 may be removed

#### _Worksheet_ synchronization
- ***New Worksheets***
  - The Sync-Source-Worksheet is cloned to the Sync-Target-Workbook
  - Back-links to the Sync-Source-Workbook are eliminated
  - All concerned Names scope is synchronized
- ***Obsolete Worksheets*** are removed
- ***Worksheets' Name <span style="color:red">****or!****</span> Code-Name change*** are synchronized
- Not yet implemented: ***Worksheets owned by the VB-Project***, that means those protected and without any unlocked (input) cell, are synchronized by default - disregarding any change.

>Attention! <span style="color:red">The _Name_ ***and*** the _CodeName_ of a Worksheet must never both be changed.</span> When a Worksheet's _Name_ ***and*** its _CodeName_ is changed at the same time the concerned sheet will be considered new and the (no longer identifiable as such) corresponding sheet will be considered obsolete - which in such a case is definitely not what was intended.
#### _References_ synchronization
New References are added and obsolete References are removed.
#### _VB-Components_ synchronization
Synchronized are all types of VB-Components:  (_Standard&nbsp;Module_, _Data&nbsp;Module_, _Class&nbsp;Module_, _UserForm_). New components are added, obsolete components are removed, and of changed components the code is updated.

#### _Names_ synchronization
- In order to provide a full transparent synchronization of Names they are synchronized ++before++ the Worksheets are synchronized [^2] - though this makes it a bit more complex. When a Worksheet's name is about to change, the _Names synchronization_ has to deal with an old-named Worksheet while the source Name refers already the new Worksheet name. 
- When the referred range of a Name has changed this is not synchronized. Instead this issue may be skipped or used to interrupt the sync in order to synchronize the change in the source Worksheet's layout in the corresponding target Worksheet manually. In case the 'change' is caused by the fact that a new row/column has been inserted by the Workbook user while the VB-Project has been maintained this issue can just be ignored, i.e. skipped.
- For a best possible transparent process, multiply named ranges are handled separated by removing all target- and adding all corresponding source-Names.
- ***Synchronization of obsolete Names:*** A Name is regarded obsolete when it only exists in the Sync-Target-Workbook but not (no longer) in the Sync-Source-Workbook
- ***Synchronization of new Names:*** A Name is considered new when it (the Name's 'mere name' [^3]) exists in the Sync-Source-Workbook but not in the Sync-Target-Workbook.
- ***Synchronization of changed Names and/or Scopes:*** A ranges Name and/or its Scope may bee changed in the Sync-Source-Workbook and will accordingly synchronized in the Sync-Target-Workbook. 
- ***Manual pre-synchronization preparation:*** When a synchronization is intentionally terminated (interrupted respectively) this will only be done in order to manually synchronize a design-change. To support this, the Sync-Source-Workbook and the Sync-Target-Workbook's working copy will be closed and the Sync-Source-Workbook re-opened. In the open dialog "manual pre-synchronization" will be chosen and once done the Workbook closed and re-opened with the option "continue with the ongoing synchronization".  

[^2]: When during the Worksheet synchronization a new sheet is cloned all Names are cloned too which obstructs a transparent Names synchronization.
[^3]: A Name objects 'mere name' is one without a sheet-name-prefix

#### _Sheet-Shape_ synchronization
Still under construction!
New Shapes (including ActiveX-Controls) are added, obsolete Shapes are removed. The Properties of all Shapes are synchronized. However, though largely covered the properties synchronization may still be incomplete. 

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
 
### Common Components
_Common Components_ are considered a key to the productivity and performance of VB-Projects, provided they are well designed and carefully tested. CompMan's aim is to support keeping them up-to-date in Vb-Proje TS using them. Even when not completely imported/used the serve as a rich source for procedures being copied.

#### The concept of "hosted" Common Components
Experience has shown than only a dedicated Workbook/VB-Project is appropriate for the development and especially the testing of a _Common Component_. It is required for the provision of a comprehensive test environment which also supports regression testing. _CompMan_ supports this concept by allowing to specify a _Common Component_ as being hosted in a Workbook. However, practice has shown that a modification or amendment  of a _Common Component_ is often triggered by a VB-Project just using, i.e. not hosting, it. _CompMan_ therefore supports this by keeping a record of which Workbook/VB-Project has last modified it.

#### The services
CompMan's initial intention was to keep _Common&nbspComponents_ up-to-date in all VB-Projects using them. To achieve this the _Export Service_ saves the Export-File of a modified used or hosted _Common Component to a _Common Components Folder_ thereby keeping a record of the modifying Workbook together with an incremented [_Revision Number_](#the-revision-number). Subsequently the _Update-Outdated-Common-Components_ service (by with the `Workbook_Open` event) checks for any outdated used or hosted _Common&nbsp;Components_ and offers an update in a dedicated dialog which allows to check the code difference by means of WinMerge ([WinMerge English][3], [WinMerge German][4].

#### The _Revision Number_
CompMan maintains for_Common Components a _Revision Number_, increased whenever it is modified. The _Revision Number_ is maintained in a file _CompMan.dat_ located in the serviced Workbook's parent  folder and kept in sync with the _Revision Number_ in a file _ComComps.dat_ located in [the _Common Components_ folder](#the-common-components-folder).


### Other CompMan specific files

| File                        | Location             | Description               |
|-----------------------------|----------------------|---------------------------|
| _***CompMan.dat***_         | The <u>serviced</u> Workbook's parent folder | PrivateProfile file for the registration of all  _Hosted&nbsp;Common&nbsp;Components_ and all _Used&nbsp;Common&nbsp;Components_. |
| _***CompMan.Service.trc***_ | The <u>serviced</u> Workbook's parent folder | Execution trace of the performed CompMan service, available only when the VB-Project's _Conditional Compile Argument_<br><nobr>`XcTrc_mTrc = 1` (mTrc is installed/used) or<br>`XcTrc_clsTrc = 1` (clsTrc is installed/used) is set. |
|  _***CompMan.Service.log***_| The <u>serviced</u> Workbook's parent folder | Log file for the executed CompMan services.|
| _***CommComp.dat***_        | The [Common-Components](#the-common-components-folder) folder | A _PrivateProfile_ file with sections representing Common Components with various information like the hosting Workbook and the _[Revision-Number](#the-revision-number)_ for instance.|

### Multiple computers involved in VB-Project's development/maintenance
When the [Common-Components](#compmans-default-files-and-folders-environment) folder is handled/managed as a GitHub repository it will be easy to keep an up-to-date clone on various computers. Currently the location of the [Common-Components](#compmans-default-files-and-folders-environment) folder is fixed and cannot be re-configured/located on a network. However, the whole environment, i.e. the [_CompManServiced_ folder](#compmans-default-files-and-folders-environment) folder may be moved to/kept at any location.

### CompMan.xlsb versus CompMan as Add-in
All services are provided by an open [CompMan.xlsb][1] Workbook even when it is additionally [setup as Addin](#setup-as-add-in). The Addin only provides the services when the [CompMan.xlsb][1] Workbook is not open. When the Addin is paused and the  [CompMan.xlsb][1] Workbook is not open no services are provided until the [CompMan.xlsb][1] Workbook is open again and the Addin is continued. The Workbook may be closed then. The advantage of the **Addin** is that it remains (almost) invisible. That's all.
> While any Workbook can use the services either form an open [CompMan.xlsb][1] Workbook ++or++ from the Addin, the [CompMan.xlsb][1] Workbook itself requires the **Addin** to update its own outdated _Used&nbsp;Common&nbsp;Components_.

### CompMan's default files and folders environment
```txt
CompManServiced
  +---CompMan
  |    +--Addin
  |    +--source
  |    +--CompMan.xlsb
  |    +--WinMerge.ini
  | 
  +---Common-Components
       +--CompManClient.bas
```

| File/Folder Name | Meaning and usage                             |
|------------------|-----------------------------------------------|
|_CompManServiced_ | Default root **folder** [serviced](#enabling-the-services-serviced-or-not-serviced) by CompMan.. The folder may be moved and/or renamed. When the [CompMan.xlsb][1] Workbook is opened it recognizes the parent of its parent folder as the [serviced](#enabling-the-services-serviced-or-not-serviced) root folder and keeps a record of it in the _CompMan.cfg_ file.|
|_CompMan_         | Default parent **folder** of the [CompMan.xlsb][1] Workbook. The name preferably defaults to the Workbook's base name.|
|_Addin_           | **Folder** for the [CompMan.xlsb][1] Workbook when [configured](#configuration-changes-compmans-config-worksheet) as [Addin (CompMan.xlsa)](#compmanxlsb-versus-compman-as-add-in).<br><u>The folder name however must not be altered!</u>|
|_source_          | Default **folder** name for the Export-Files of changed components exported with each Save event. This folder is maintained for each [serviced](#enabling-the-services-serviced-or-not-serviced) Workbook in the Workbook's dedicated parent folder. The folder name may be changed by means of the _Config_ Worksheet.|
|_CompMan.cfg_     | ***PrivateProfile*** file which keeps the current CompMan configuration. It is used with each open and adjusted on the fly if required. The file ensures that each subsequent download of the ComMan.xlsb Workbook works with the configuration saved with the last close of it.|
|_CompMan.xlsb_    | The Workbook **file** originally opened copied (saved as) into its dedicated parent folder along with the initial setup. Once the opened Workbook has been saved to the new setup location it is deleted.|
|_WinMerge.ini_    | **File** used by CompMan to display code changes by means of _***WinMerge***_ with the options to ignore empty code lines and ignore case differences.|
|_Common&#8209;Components_ | Default **folder** where CompMan maintains a copy of the Export-File of [hosted](#the-concept-of-hosted) _Common&nbsp;Component_. These _Export Files_ function as the source for a [serviced](#enabling-the-services-serviced-or-not-serviced) Workbook's (possibly) outdated _Used&nbsp;Common&nbsp;Components_. Though primarily maintained by CompMan, the folder may contain any Export File of a VBComponent considered ***Common***. Manually added Export-Files are regarded ***Common Component orphans*** until a Workbook claims it [Hosted](#the-concept-of-hosted) |
|_CompManClient.bas_| Export **file** of the _Common Component_ [hosted](#the-concept-of-hosted) by [CompMan.xlsb][1] for being imported into any to-be-serviced Workbook. See [Enabling the services](#enabling-the-services-serviced-or-not-serviced).| 

6. The [CompMan.xlsb][1] Workbook requires two _Macro Security Settings:_ <br> 1. _Trust Center_ > _Trust Center Settings_ > _Macro Settingsin_: "Deactivate all macros except this signed" and "Trust the access to the VBAProject Object model"
7. Providing the code in the IDE with a (SelfCert.exe) signature will be the perfect fit with the above setting 

### Configuration changes (CompMan's _Config_ Worksheet)

| Item, means                      | Meaning, usage                                                            |
|----------------------------------|---------------------------------------------------------------------------|
| _CompMan's&nbsp;serviced&nbsp;root&nbsp;folder_ | May be moved to any other location and/or renamed when the Workbook is closed. When the [CompMan.xlsb][1] Workbook is opened again it by default regards the parent folder of the parent folder as the [serviced](#enabling-the-services-serviced-or-not-serviced) root **folder**. I.e. any Workbook in a subsequent folder will be [serviced](#enabling-the-services-serviced-or-not-serviced) provided it is [enabled](#enabling-the-services-serviced-or-not-serviced). |
| _Export folder_                  | **Name** of the folder established and used in any [serviced](#enabling-the-services-serviced-or-not-serviced) Workbook's parent folder to store the Export-Files of changed (or yet not exported) VBComponents.|
| _Serviced&nbsp;Sync&#8209;Target&#8209;Folder_ | **Folder** into which a Workbook (for which the _Synchronize VB-Project_ has been [enabled](#enabling-the-services-serviced-or-not-serviced) is temporarily moved and opened in order to have its _VB-Project_ synchronized with the corresponding (same named) Workbook residing in its dedicated folder within the configured [_CompManServiced_ folder](#compmans-default-files-and-folders-environment). When the configuration of the _Serviced Synchronization Target Folder_ is terminated, i.e. no folder is selected, the _Serviced Synchronization Target Folder_ becomes 'not configured'.|
| _SyncArchive folder_ | **Folder** obligatory for the _Synchronize VB-Project_ service which archives a _Sync-Target-Workbook_ before it is synchronized with its corresponding _Sync-Source-Workbook_. When the _Synchronization Archive Folder_ selection dialog is terminated, i.e. no folder is selected, the _Synchronization Archive Folder_ becomes 'not configured'.|
| _CompMan&#8209;Workbook&nbsp;status_        | CompMan's current status which may be changed by the _Setup Auto-open_/_Remove Auto-open_ **Command Button** |
|<a id='setup-as-add-in'></a><nobr>_CompMan Addin status_ | Status provided by the _Provide Add-in/Give up Add-in_ **Command Button**.|
|***Setup&nbsp;Auto&#8209;open***<br>***Remove&nbsp;Auto&#8209;open***| Sets up or reomes the Auto-Open for the _CompMan.xlsb_.|
|***Pause&nbsp;Add&#8209;in***<br>***Continue&nbsp;Add&#8209;in***| **Command Button** to temporarily pause and subsequently continue the setup Add-in.|
|</a><nobr>***Give up Add-in***<br>***Provide Add-in***| ***Provide Add-in*** establishes the CompMan.xlsb as Add-in automatically opened when Excel starts.<br>***Give up Add-in*** removes the Addin (even when it is currently open, which requires a couple of tricks).| 


### Download from public GitHub repo
It may appear pretty strange when downloading first from a public GitHub repo but is is quite straight forward as the below image shows.  
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
