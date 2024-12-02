## Abstract
CompMan provides services for professional and semi-professional Excel VBProject developers with special services regarding _[Common Components](#common-components)_.

## Disambiguation
| Term             | Meaning                  |
|------------------|------------------------- |
| _Component_       | Generic term for a _VB-Project's_ _VBComponent_ (_Class Module_,  _Data Module_, _Standard Module_, or _UserForm_) |
| _[Common&nbsp;Component](#common-components)_ | A component providing services for a certain subject, potentially being used in more than one VB-Project. |
| _Procedure_           | Generic term for any _VB-Component's_ (public or private) `Property`, `Sub`, or `Function`|
|_Service_             | 1. What CompMan provides<br>2. Generic term for any public _Property_, _Sub_, or _Function_. |
| _Servicing&nbsp;Workbook_ | The service providing Workbook (ThisWorkbook), either the _[CompMan.xlsb][1]_ Workbook (when it is open) or the _CompMan Add-in_ when it is set up and open. |
| _Serviced&nbsp;Workbook_ | The `ActiveWorkbook` prepared for being [serviced](#enabling-the-services-serviced-or-not-serviced) (ActiveWorkbook).
|_VB&#8209;Project_    | Used synonymous with Workbook |
| _Workbook&nbsp;parent&nbsp;folder_ | A folder dedicated to a _Workbook/VB-Project_. Note that an enabled Workbook is only [serviced](#enabling-the-services-serviced-or-not-serviced) when it is **exclusive** in its parent folder. Other Workbooks may be located in sub-folders however.|
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

### _Hosted_ versus _Used Common Components_
Careful testing of a _Common Component_ is crucial for their performance boosting impact on VBProject development. Such a careful testing (in a best practice including regression testing) can only be achieved in a dedicated Workbook/VBProject. This dedication is expressed and documented by the fact that a Workbook/VBProject claims the _Common Component_ being _hosted_ in it. In contrast, it will be regarded _used_ in all other Workbooks. However, although the hosting Workbook is dedicated for testing a _Common Component_ may (ad-hoc) be modified in any Workbook using it.

### Pending releases
When a _Common Component_ is modified in a Workbook/VBProject using or hosting it the modification may remain incomplete/not finalized for some time. When the Workbook is saved (early and often hopefully) the _Common Component_ will become "pending release". I.e. a copy of the Export-File is copied to the `PendingReleases` folder in the `Common-Components` folder, waiting for being released to public (moved into the `Common-Components` folder).

### Release to public
When the modification of a Common-Component had been finalized it will be released to public.Manually by copying the Export-File to the `Common-Components` folder or via the `CompMan` menu int the VBE (only available when there's at least one _Common Component_ pending release. The menu provides the chance to finally check which modifications had been done by having the pending release code compared with the public code.

### Properties of _Common Components_
The below properties are maintained for _Common Components_ in Private Profile files in order to document  who, when, where, on which computer last modified the component :
1. For any serviced Workbook (`...\CompMan\CommComps.dat`)
2. For any _Common Component in the Common-Components folder (`...\Common-Components\CommComps.dat`)
3. For any pending release (`...\Common-Components\PendingReleases.dat`)

| Value name               | Meaning, purpose |
|--------------------------|------------------|
|***KindOfComponent***     |`used` or `hosted`|
|***LastModAt***           |Last modified date/time in the format `yyyy-mm-ddd hh:mm:ss (UTC)`.|
|***LastModBy***           |User who last modified the _Common Component_. |
|***LastModExpFileOrigin***|The full name of the origin _Export-File_. I.e. the Export-File which had finally been copied into the `Common-Components` folder.|
|***LastModIn***           |The Workbook/VBProject in which the _Common Component_ had last been modified.|
|***LastModOn***           |The Computer's name on which the used/hosted Common Component had last been modified.|

## Other
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
The following is requirements have Workbook being serviced by CompMan:
1. A ***servicing CompMan instance*** is open. That is either the [CompMan.xlsb][1]) Workbook or its [Addin instance](#compmanxlsb-versus-compman-as-add-in)
2. The ***to-be-serviced Workbook*** Workbook has been prepared to call a service
3. The ***to-be-serviced Workbook*** is opened from within a sub-folder of the configured [_CompManServiced_ root folder](#compmans-environment)
4. The ***to-be-serviced Workbook*** is the only Workbook in its parent folder (the parent folder may have sub-folders with Workbooks however)
5. WinMerge ([WinMerge English][3], [WinMerge German][4] or any other language version is installed to display the difference for any components when about to be updated.

### CompMan's Environment
### The scheme
When the _[CompMan.xlsb][1]_ Workbook has been downloaded, possibly moved to its dedicated folder (a folder in which it is the only Workbook) and opened the environment is checked and if not yet done established by a self-setup. Below is a scheme of .
```txt
<CompMan serviced root folder> CompMan serviced root folder
 +--Common-Components          Public Common Component folder 
 |  +--CommComps.dat           Private Profile file for public Common Component properties
 |
 +--CompMan                    Folder dedicated to the CompMan.xlsb Workbook 
 |  +--CompMan                 *) CompMan's own CompMan environment folder (see below)
 |  |  +--<export files>       Folder for all exported (new or modified) components 
 |  |     +--.....
 |  |     +--.....
 |  |
 |  |  +--CommComps.dat
 |  |  +--ExecTrace.log
 |  |  +--Services.log
 |  |  +--ServicesSummary.log  
 |  +--CompMan.xlsb            CompMan Workbook providing services when open 
 |  +--WinMerge.ini            Configuration for the dsiplay of code differences
 |  
 +--CompManAddin               Folder for CompMan's Addin instance when setup
    +--CompMan.xlam            CompMan's Addin instance Workbook - when configured
  ```
#### Folders and files
All folders and files are automatically "self-setup" when the _[CompMan.xlsb][1]_ Workbook is downloaded, moved to its serviced root folder and opened.
##### CompMan serviced root folder
This folder is the key for Workbooks/VBProjects for being serviced because only Worbook located therein will be serviced, provided they are configured accordingly. This folder is recognized whenever the _[CompMan.xlsb][1]_ Workbook or the [Addin instance](#compmanxlsb-versus-compman-as-add-in) is opened (as the Parent.Parent folder). Consequently the folder may have any name, may be renamed at any time or its whole content may be moved to any other location.
##### _Common-Components_ folder
Sub-folder of the [CompMan serviced root folder](#compman-serviced-root-folder) with a fixed name (not configurable). The folder stores all public _Common Components_, their Export-File respectively.
##### _CommComps.dat_ in Common-Components folder
1. _Private Profile_ file in the [Common-Components folder](#common-components-folder) for the maintenance of properties concerning the origin of public _Common Components_.
2. _Private Profile_ file in the CompMan service folder (in any serviced Workbooks dedicated folder) for the maintenance of properties concerning the origin of public _Common Components_ in case of an update. When a _Common Component_ is modified in a Workbook using or hosting it these properties all point to the exporting Workbook/VBProject.

##### _CompMan_ folder
1. Dedicated folder for the _[CompMan.xlsb][1]_ Workbook (analog to any other serviced Workbook's dedicated folder). 
2. Folder maintained in any serviced Workbooks dedicated folder containing CompMan service relevant files and the Export-Folder.

##### _Export_ folder
Folder in the [_CompMan_ folder](#compman-folder) storing all exported (because modified) components of a serviced Workbook/VBProject. The folder's default name is `source` but the name may be configured locally.
##### ExecTrace.log
File in the [_CompMan_ folder](#compman-folder), maintained due to the fact that CompMan uses the Common Component mTrc/clsTrc to trace and log its execution. This exceutin trace is triggered by the  _Conditional Compile Argument_ <nobr>`mTrc = 1`</nobr> .
##### Services.log
File for logging the result of the Export and the Update service.
##### `CompManAddin`
Folder in the [CompMan serviced root folder](#compman-serviced-root-folder) for the CompMan's [Addin instance](#compmanxlsb-versus-compman-as-add-in) (when configured).

### CompMan.xlsb versus CompMan as Add-in
First and foremost services may be provided by an open _[CompMan.xlsb][1]_ Workbook (even when it is additionally [setup as Addin](#setup-as-add-in). Because the Addin instance once configured is automaticall opend when Excel starts, the CompMan services are available automatically. The advantage: CompMan remains (almost) invisible and the _[CompMan.xlsb][1]_ Workbook needs not to be opened - alternatively it may be configured to auto-open however. That's all.
> While any Workbook can use the services either form an open _[CompMan.xlsb][1]_ Workbook ++or++ from the Addin, the _[CompMan.xlsb][1]_ Workbook itself requires the **Addin** to keep its used/hosted _Common Components_ up-to-date.
> Attention with VBProjects developed/maintained for clients: Public procedures of an open _[CompMan.xlsb][1]_ Workbook or a configured Addin may be used in any VBProject. but when the Workbook runs in another environment (at the client's site for instance) the VBProject will not run/compile). Developers need to make sure a Workbook runs without error even when CompMan is not available at all.

### Configuration changes
CompMan, when opened, displays a _Config_ Worksheet which allows several configuration changes.

| Item                      | Meaning, usage                       |
|---------------------------|--------------------------------------|
| _CompMan's&nbsp;serviced&nbsp;root&nbsp;folder_ | **Not configurable!**<br>When _[CompMan.xlsb][1]_ is opened it regards the parent folder of its own dedicated folder as the serviced root folder. I.e. any Workbook in a subsequent dedicated folder will be [serviced](#serviced-or-not-serviced) provided the Workbook is prepared for it.|
|_\<export folder>_          | **Name configurable**<br>When the name is changed any old named export folder will be renamed as soon as a Workbook is serviced in the future. This is possible because the `CompMan.cfg` keeps a record with the history name(s).|
| _CompMan&#8209;Workbook&nbsp;status_ | Shows the current status which may be changed by the _Setup Auto-open_/_Remove Auto-open_ **Command Button** |
|<a id='setup-as-add-in'></a><nobr>_CompMan Addin status_ | Status provided by the _Provide Add-in/Give up Add-in_ **Command Button**.|

|***Setup&nbsp;Auto&#8209;open***<br>***Remove&nbsp;Auto&#8209;open***| Sets up or reomes the Auto-Open for the _[CompMan.xlsb][1]_.|
|***Pause&nbsp;Add&#8209;in***<br>***Continue&nbsp;Add&#8209;in***| **Command Button** to temporarily pause and subsequently continue the setup Add-in.|
|</a><nobr>***Give up Add-in***<br>***Provide Add-in***| ***Provide Add-in*** establishes the _[CompMan.xlsb][1]_ as Add-in automatically opened when Excel starts.<br>***Give up Add-in*** removes the Addin (even when it is currently open, which requires a couple of tricks).| 

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

## Appendix
### Formatting
`physically reprented item` is formatted like code is
_technical item_ is formatted italic

[1]:https://github.com/warbe-maker/VBA-Components-Management-Services/blob/master/CompMan.xlsb?raw=true
[2]:https://www.sync.com/
[3]:https://winmerge.org/downloads/?lang=en
[4]:https://winmerge.org/downloads/?lang=de
