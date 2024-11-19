## Abstract
CompMan provides services for professional and semi-professional Excel VBProject developers with special services regarding _[Common Components](#common-components)_.

## Services
### _Export service_
Exports **modified** components into a serviced Workbook's dedicated folder (`...\CompMan\source`). Modifies used/[hosted](#the-concept-of-hosted-common-components) _Common Components_ are (re)registered a _pending release_ considering that the modifications are not final yet.
> When combined with a cloud synchronization service like [sync.com][2] the _Export service_ thereby provided a full versioning of modification on a component base (instead of one on the VBProject as a whole).

### _Update Service_
Updates any component in a [serviced Workbook's](#serviced-or-not-serviced) VBProject of which the code is not/no longer identical with the current public _Common Component's_ code whereby the proposed update may be postponed or denied. In the latter case the component's registration state is turned into _private_, i.e. excluded from future updates. See the [concept of Common Components](#concept)

## Serviced Workbooks/VBPprojects
The services are provided with a minimum intervention in VBProjects. An additional component is used as an interface to CompMan's services, one code line in the _Workbook\_Open_ and one in the _WorkBook\_Before\_Close_ event procedure initiates the service provided the required conditions are meet (see [Serviced or not serviced Workbooks](#serviced-or-not-serviced)).

## _Common Components_
### Concept
Components of which procedures (Properties, Methods, Procedures, Functions) are potentially usable/valuable for other VBProjects are maintained in a dedicate folder and automatically (dialog based) updated in Workbooks/VBProjects using them.

### Public Common Components
Public _Common Components_ are maintained (as _Export-Files_) in a `Common-Components` folder in the _CompMan-serviced-root-folder_, which is the parent folder of a [serviced Workbook's](#serviced-or-not-serviced) parent folder. The following properties of Common Components are maintained in the `Common-Components` folder in a Private Profile file called `Common-Components\CommComps.dat` in which each Common Component is represented as a _Section_:

| Value name | Meaning, purpose |
|------------|------------------|
| ***KindOfOriginComponent***| Registration status of the last "published" component in its origin Workbook, either `used` or `hosted`|
|***LastModAt***             | The date/time in the format `yyyy-mm-ddd hh:mm:ss (UTC)` the origin _Common Component_ had last been modified.|
|***LastModBy***             | User who last modified the Common Component. |
|***LastModExpFileOrigin***  | The origin _Export-File_ which had been copied into the `Common-Components` folder for being available public. |
|***LastModIn***             | The Workbook/VBProject in which the Common Component had last been modified.|
|***LastModOn***             | The Computer's name on which the Common Component had last been modified.|

#### Used/Hosted Common Components
Properties of used/[hosted](#the-concept-of-hosted-common-components) Common Components are maintained in a _Private Profile File_ in the [serviced Workbooks](#serviced-or-not-serviced) parent folder in a dedicated `CompMan` folder (`...\CompMan\CommComps.dat`):

| Value name               | Meaning, purpose |
|--------------------------|------------------|
|***KindOfComponent***     |`used` or `hosted`|
|***LastModAt***           | The date/time in the format `yyyy-mm-ddd hh:mm:ss (UTC)` the origin of the used/[hosted](#the-concept-of-hosted-common-components) Common Component had last been modified.|
|***LastModBy***           | User who last modified the Common Component |
|***LastModExpFileOrigin***| The origin _Export-File_ which had been imported via the `Common-Components` folder by CompMan's _[Update service](#update-service)_ - or the first time imported manually.|
|***LastModIn***           | The Workbook/VBProject in which the used Common Component had last been modified.|
|***LastModOn***           |The Computer's name on which the used/hosted Common Component had last been modified.|

## Other
### The concept of _hosted_ versus _used_ Common Components
Careful testing, including regression testing is a key issue for components potentially used in more than one Workbook/VBProject. The only way to achieve this and provide all the means for it is a dedicated Workbook/VBProject. This dedication is expressed and documented by the fact that the Common Component is _hosted_ in its dedicated Workbook - whereas in contrast, in all other Workbooks it is just _used_. Although the hosting Workbook is dedicated for testing a _Common Component_ may be (ad-hoc) modified in any Workbook using it (sometimes appropriate in reality).

### Manual interventions
Developers using Common Components likely have established a routine for their management and also a routine for frequently exporting all components. Any of them likely will become obsolete when CompMan services are established and used. Considering the first said, CompMan integrates manual intervention by a [housekeeping](#housekeeping) routine performed before the _[Export-](#export-service)_ and the _[Update service](#update-service)_. Typical manual interventions may bee manually copying the _Export-File_ of a modified _Common Component_ into the `Common-Components` folder or manually importing/re-importing a public _Common Component_ from the folder.

### Housekeeping
Housekeeping is an effort prior the _Export-Service_ and the _Update_Service which provides up-to-date and consistent data.
#### Housekeeping public _Common Components_
Maintains the properties for each known _Common Component_ in the Private Profile file `Common-Components\CommComps.dat`.  
- **Removes** component representing sections of which the corresponding _Export-File_ not/no longer exists in the `Common-Components` folder  
- **Adds** missing component representing sections for new _Export-Files_ in the `Common-Components` folder which likely will have been copied manually. Therefore, when the _Serviced Workbook's_ corresponding _Export-File_ is identical, this one is regarded the origin (***LastModExpFileOrigin***). If not, the origin is regarded unknown and set identical with the _Export-File_ in the `Common-Components` folder.

#### Housekeeping serviced public _Common Components_
To be aware of: In the _Serviced Workbook_ a component is regarded public when a corresponding section exists in the `Common-Components\CommComps.dat` file or when it is indicated _hosted_ in the _Serviced Workbook's_  `CompMan\CommComps.dat` Private Profile file_ (may yet not be public however).
- **Hosted**: Common Components are registered in the _Serviced Workbook's_ `CompMan\CommComps.dat` file
- **Not hosted**: Sections representing no longer _hosted_ Common Components are removed
- **Kind of component**: When a component in the _Serviced Workbook_ is yet not registered but its name is identical with a known public _Common Component_ it is requested being registered as a _used_ or a _private_ component (in case it by chance just has the same name)
- **Obsolete**: Sections representing components not/no longer known a public _Common Component_ are removed
- **Pending Release**: Maintains a consistent pending releases data base by:
  - Removing _pending release Common Components_ of which the registered _Export-File_ is already identical with the current public _Common Component's Export-File_ in the `Common-Components` folder.
  - Removing entries in the Pending.dat file without a corresponding _Export-File_ in the `Common-Components\PendingReleases` folder
  - Removing _Export-Files without a corresponding entry in the `Pending.dat` file in the `Common-Components\Pending` folder
- **Pending Release outstanding**: Registers a yet not registered pending component when its `ServicedLastModAt` property is greater the the `PublicLastModAt` property and the current code differs from the public code
- **Properties**: Maintains for all serviced _Common Components_ the properties in the serviced
' workbook's `CompMan\CommComps.dat` file - specifically when the ***LastModAt*** property differs from the public ***LastModAt*** property although the code is identical with the public version.

### Serviced or not serviced
The following requirements have to met in order to have/get a [Workbook serviced by CompMan](#serviced-or-not-serviced):
1. A ***servicing CompMan instance*** is open (either the [CompMan.xlsb][1]) Workbook or its Addin instance
2. The ***to-be-serviced Workbook*** Workbook has been prepared to call a service
3. The ***to-be-serviced Workbook*** is opened from within a sub-folder of the configured [_CompManServiced_ root folder](#compmans-environment)
4. The ***to-be-serviced Workbook*** is the only Workbook in its parent folder (the parent folder may have sub-folders with Workbooks however)
5. WinMerge ([WinMerge English][3], [WinMerge German][4] or any other language version is installed to display the difference for any components when about to be updated.


### CompMan's Environment
When the CompMan.xlsb Workbook has been downloaded, moved into its dedicated folder and opened it runs a SelfSetup check - and when not existing to provide - its environment of folders and files. See below.
```txt
CompManServiced                            CompMan serviced root folder
  +---Common-Components                    Public Common Component folder 
  |   +--CommComps.dat                     Private Profile file for public Common Component properties
  +---CompMan
  |   |   +--CompMan                           *) CompMan's own CompMan environment folder (see below)
  |   |      +--source
  |   |      |  +--<export files>
  |   |      |  +-- .....
  |   |      +--CommComps.dat
  |   |      +--ExecTrace.log
  |   |      +--Services.log
  |   |      +--ServicesSummary.log
  |   +--CompMan.xlsb                      CompMan Workbook providing services when open 
  |   +--WinMerge.ini                      Configuration for the dsiplay of code differences
  |      
  +--CompManAddin                             Folder for CompMan's Addin instance when setup
  +--<serviced-workbook>
  |   +--CompMan                           *) CompMan's environment folder (see below)
  |   |  +--source
  |   |  |  +--<export files>
  |   |  |  +-- .....
  |   |  +--CommComps.dat
  |   | +--ExecTrace.log
  |   |  +--Services.log
  |   +--yyyyyy.xls                       Serviced Workbook
  +--<serviced-workbook>
  ...
 ```

| File/Folder         | Location | Description           |
|---------------------|----------|------------|
|`CompManServiced`    |any      |Folder recognized by CompMan as the serviced root folder, the parent of CompMan.xlsb parant folder, may have any name, may be renamed at any time. However, when an Addin instance had been setup in needs to be re-setup.|
|`Common-Components`  |serviced root folder| Folder dedicated for all public _Common Components_|
|`CompComps.dat`      |`Common-Component` folder|Private Profile file for [used/hosted Common Components](#usedhosted-common-components) properties.|
|`CompMan`            |any [serviced Workbook's](#serviced-or-not-serviced) root folder | CompMan's envirionment folder |
|`source`             |CompMan's environment folder| Folder, stores all Export-Files.|
|`CompComps.dat`      |CompMan's environment folder| Private Profile file for [used/hosted Common Components](#usedhosted-common-components) properties.|
|`ExecTrace.log`      |CompMan's environment folder| Execution trace log of performed services, available only when the VB-Project's _Conditional Compile Argument_<br><nobr>`mTrc = 1` (i.e.mTrc is installed/used) or<br>`clsTrc = 1` (clsTrc is installed/used) is set. |
|`Services.log`       |CompMan's environment folder| Log file with the last 5 services provided for the [serviced Workbook](#serviced-or-not-serviced).|
|`ServicesSummary.log`|CompMan's environment folder| Log file with a service summary for the last 10 services provided for any [serviced Workbook](#serviced-or-not-serviced), Maintained in CompMan's own `CompMan` folder only.|

### CompMan.xlsb versus CompMan as Add-in
All services are provided by an open [CompMan.xlsb][1] Workbook even when it is additionally [setup as Addin](#setup-as-add-in). When the Addin is paused and the  [CompMan.xlsb][1] Workbook is not open no services are provided until the [CompMan.xlsb][1] Workbook is open again and the Addin is continued. The advantage of the **Addin** is that it remains (almost) invisible and the [CompMan.xlsb][1] Workbook needs not to be opened - it may be configured to auto-open however. That's all.
> While any Workbook can use the services either form an open [CompMan.xlsb][1] Workbook ++or++ from the Addin, the [CompMan.xlsb][1] Workbook itself requires the **Addin** to keep its used/hosted _Common Components_ up-to-date.

### Configuration changes (CompMan's _Config_ Worksheet)

| Item, means                      | Meaning, usage                       |
|----------------------------------|--------------------------------------|
| _CompMan's&nbsp;serviced&nbsp;root&nbsp;folder_ | May be moved to any other location and/or renamed when the Workbook is closed. When the [CompMan.xlsb][1] Workbook is opened again it by default regards the parent folder of the parent folder as the serviced root **folder**. I.e. any Workbook in a subsequent folder will be [serviced](#serviced-or-not-serviced) provided it is enabled for it.|
|_Export folder_                  | **Name** of the folder established and used for any [serviced Workbook](#serviced-or-not-serviced) Workbook. When its name is changed EnvironmentHousekeeping will forward any outdated folder to its new name. However, when the changed name is changed again it needs to be done in the form old-nam>new-name.|
| _CompMan&#8209;Workbook&nbsp;status_        | CompMan's current status which may be changed by the _Setup Auto-open_/_Remove Auto-open_ **Command Button** |
|<a id='setup-as-add-in'></a><nobr>_CompMan Addin status_ | Status provided by the _Provide Add-in/Give up Add-in_ **Command Button**.|
|***Setup&nbsp;Auto&#8209;open***<br>***Remove&nbsp;Auto&#8209;open***| Sets up or reomes the Auto-Open for the _CompMan.xlsb_.|
|***Pause&nbsp;Add&#8209;in***<br>***Continue&nbsp;Add&#8209;in***| **Command Button** to temporarily pause and subsequently continue the setup Add-in.|
|</a><nobr>***Give up Add-in***<br>***Provide Add-in***| ***Provide Add-in*** establishes the [CompMan.xlsb][1] as Add-in automatically opened when Excel starts.<br>***Give up Add-in*** removes the Addin (even when it is currently open, which requires a couple of tricks).| 

## Appendix
### Formatting
`physically reprented item` is formatted like code is
_technical item_ is formatted italic

[1]:https://github.com/warbe-maker/VBA-Components-Management-Services/blob/master/CompMan.xlsb?raw=true
[2]:https://www.sync.com/
[3]:https://winmerge.org/downloads/?lang=en
[4]:https://winmerge.org/downloads/?lang=de