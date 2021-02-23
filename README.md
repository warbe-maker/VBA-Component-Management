# Management of Excel VB-Project Components
Services for the installation, export, update, and synchronization of _Clone-Components_ with their _Raw-Component_,
by means of a _CompMan_ Addin Workbook. The Addin's development instance is this repo and provides the service to 
establish itself as Addin.
See also [Programatically updating Excel VBA code][2]


# Services
## _ExportChangedComponents_
Used with the _Workbook_Before_Save_ event it compares the code of any component in a _VB-Project_ with its last _Export File_ and re-exports it when different. The service is essential for _VB-Projects_ which host _Raw-Components_ in order to get them registered as available for other _VB-Projects_. Usage by any _VB-Project_ in a development status is appropriate as it is not only a code backup but also perfectly serves versioning - even when using [GitHub][]. Any _Component_ indicated a _hosted Raw-Component is registered as such with its _Export File_ as the main property.<br>
The service also checks a _Clone-Component_ modified within the VB-Project using it a offers updating the _Raw-Component_ in order to make the modification permanent. Testing the modification will be a task performed with the raw hosting project.

## _UpdateRawClones_
The service is used with the _Workbook\_Open_ event. It checks each _Component_ for being known/registered as _Raw-Component_,  _hosted_ by another _VB-Project_ - which means it is a _Clobmne-Component_. If yes their code is compared and suggested for being updated if different.

## _SyncVbProject_
**pending implementation**<br>
Synchronizing VB-Projects is a means for temporarily uncoupling the productive use of a Workbook from the development and maintenance of the VBA code. By this the down time of the Workbook can be minimized. Without such a sync service, the productive use of a Workbook has to stop while the VB-Project is modified which may take several days!.

In contrast to the other two services, code synchronization is not supported via Workbook events automating them, but has to be initiated manually. The advantage: Even the Workbook component can fully be synchronized.

### Syncronized elements
| Element | Extent of synchronization |
| ------- | ------------------------- |
|_Standard Modules_<br>_Class Modules_<br>_UserForms_| Code and module name synchronized |
|_Data Module_|**Workbook**: Code and module name synchronized<br>**Worksheet**: only partially (see [Worksheet synchronization](#worksheet-synchronization) and [Planning the release of a VB-Project modification](#planning-the-release-of-a-vb-project-modification))|
|_References_ | missing added and obsolete removed|
|_Controls_ | still in question |

### Worksheet synchronization
While the code of the development instance of a VB-Project is modified the productive instance will (can) continuously be used for data changes.  Because the user is able not only to change data but also the name and the position if a sheet, synchronization depends on an unchanged _CodeName_ as the only stable way to address a sheet in VBA code.

Design changes on a worksheet may only be made in concert with the **release of a VB-Project** (_VB-Raw-Project_) to production:
- Insert of new named ranges
- Change of a sheet's _CodeName_


# Installation
The _Component Management Services_ may be used simply by having the _development instance Workbook_ opened (see [Usage without Addin instance](#usage-without-addin-instance)) or by having them setup as _Addin-Workbook_.

1. Download and open [CompMan.xlsb][1]


2. Use the built-in Command button to run the _Renew_ service.
2.1. Follow the instructions to identify a location for the Add-in - preferably a dedicated folder like ../CompMan/Add-in. The folder will hold the following files:
   - CompMan.cfg    ' the basic configuration
   - CompMan.xlam   ' the Add-in
   - HostedRaws.dat ' the specified raws hosted in any Workbook
   - RawHost.dat    ' the Workbooks which claim raws hosted
   
2.2. Follow the instructions to identify a 'serviced root'
  It will:
   - ask to confirm or change the basic configuration
   - initially setup or subsequently renew the CompMan Add-in by saving a copy  of the development instance as Add-in (mind the fact that this is a multi-step process which may take some seconds)

Once the Addin is established it will automatically be loaded with the first Workbook opened having it referenced. See the Usage below for further required preconditions.

### Workbooks/VB-Projects hosting raws or using raw clones
1. Copy the following into the Workbook component
```vb
Option Explicit

Private Const HOSTED_RAWS = ""

Private Sub Workbook_Open()
#If CompMan Then
    mCompMan.UpdateRawClones uc_wb:=ThisWorkbook _
                           , uc_hosted:=HOSTED_RAWS
#End If
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
#If CompMan Then
    mCompMan.ExportChangedComponents ec_wb:=ThisWorkbook _
                                   , ec_hosted:=HOSTED_RAWS
#End If
End Sub
```
2. For a Workbook which hosts _Raw-Components_ specify them in the HOSTED_RAWS constant. If its more then one, have the component's names delimited with commas.

## Usage without Addin instance


## Planning the release of a VB-Project modification

...



[1]:https://gitcdn.link/repo/warbe-maker/VBA-Components-Management-Services/master/CompMan.xlsb
[2]:https://warbe-maker.github.io/warbe-maker.github.io/vba/excel/code/component/management/2021/02/05/Programatically-updating-Excel-VBA-code.html