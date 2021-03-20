# Excel Common Components Management
Services for the management of
- individual Common Raw VBA Components hosted in one and used in many other Workbooks/VB-Projects
- VB-Raw-Projects and their corresponding VB-Clone-Project, i.e. synchronizing the code of the clone project with the code in its corresponding raw project

Whereby the services are available via an Addin-Workbook which is setup/established/renewed by means of a _Development-Instance_ (this repo!) which provides the _Renew_ service.<br>
See also [Programmatically updating Excel VBA code][2]

## Disambiguation
Terms used in this VB-Project and all posts related to the matter.

| Term             | Meaning                  |
|------------------|------------------------- |
|_Component_       | Generic _VB&#8209;Project_ term for a _VB-Project-Component_ which may be a _Class Module_ , a  _Data Module_, a _Standard Module_, or a _UserForm_  |
|_Common Component_ | A _VB-Component_ which is hosted in one and commonly used by two or more _VB-Projects_ |
|_Clone&#8209;Component_ <br> | The copy of a _Raw&#8209;Component_ in a _VP&#8209;Project_ using it. _Clone-Components_ are kept up-to-date by an _UpdateRawClones_ service.<br>The term _clone_ is borrowed from GitHub but has a slightly different meaning because the clone is usually not maintained but the _raw_. |
|_Procedure_     | Any - Public or Private _Property_, _Sub_, or _Funtion_ of a _Component_. See also _Service_.
|_Raw&#8209;Component_ | The instance of a _Common Component_ which is regarded the developed, maintained and tested 'original', hosted in a dedicated _Raw&#8209;Host_ Workbook. The term _raw_ is borrowed from GitHub and indicates the original version of something |
|_Raw&#8209;Host_      | The Workbook/_VB-Project_ which hosts the _Raw-Component_ |
|_Service_       | Generic term for any _Public Property_, _Public Sub_, or _Public Funtion_ of a _Component_ |
|_VB&#8209;Clone&#8209;Project_ | A _VP-Project_ which is a copy (i.e regarding the VB-Project code a clone) of a corresponding  _VB&#8209;Raw&#8209;Project_. The code of the clone project is kept up-to-date by means of a code synchronization service. |
|_VB-Project_     | In the present case this term is used synonymously with Workbook |
|_VB&#8209;Raw&#8209;Project_   | A code-only _VP-Project_ of which all components are regarded _Raw-Components_. A _Raw-Project_ is kind of a template for the productive version of it. In contrast to a classic template it is the life-time raw code base for the productive _Clone-Project_.  The service and process of 'synchronizing' the productive (clone) code with the raw is part of the _Component Management_.|
| _Workbook-_, or<br>_VB&#8209;Project&#8209;Folder_ | A folder dedicated to a Workbook/VB-Project with all its Export-Files and other project specific means. Such a folder is the equivalent of a Git-Repo-Clone (provided Git is used for the project's versioning which is recommendable |


# Management of Excel VB-Project Components
Services for the installation, export, update, and synchronization of _Clone-Components_ with their _Raw-Component_,
by means of a _CompMan_ Addin Workbook. The Addin's development instance is this repo and provides the service to 
establish itself as Addin.
See also [Programatically updating Excel VBA code][2]


# Services
## _ExportChangedComponents_
Used with the _Workbook_Before_Save_ event it compares the code of any component in a _VB-Project_ with its last _Export-File_ and re-exports it when different. The service is essential for _VB-Projects_ which host _Raw-Components_ in order to get them registered as available for other _VB-Projects_. Usage by any _VB-Project_ in a development status is appropriate as it is not only a code backup but also perfectly serves versioning - even when using [GitHub][]. Any _Component_ indicated a _hosted Raw-Component is registered as such with its _Export-File_ as the main property.<br>
The service also checks a _Clone-Component_ modified within the VB-Project using it a offers updating the _Raw-Component_ in order to make the modification permanent. Testing the modification will be a task performed with the raw hosting project.

## _UpdateRawClones_
Used with the _Workbook\_Open_ event, checks each _Component_ in the VB-Project for being known/registered as _Raw-Component_ hosted by another _VB-Project_ by comparing the Export-Files. When they differ, the raw's _Export-File_ is used to 'renew' the _Clone-Component_.

## _SyncVBProject_
Synchronizes the code of a _Target-Workbook/VBroject_ with a corresponding _Source-Workbook/VBProject_ with the following covered:
- _Standard Modules_ new, obsolete, code change
- _Class Modules_ new, obsolete, code change
- _UserForms_ new, obsolete, code/design change
- _Worksheets_ new, obsolete, code change
- _Workbook_ (Document-Module): code change
- _Shapes_ new, obsolete, properties
- _References_ new, obsolete

## _UpdateRawClones_
The service is used with the _Workbook\_Open_ event. It checks each _Component_ for being known/registered as _Raw-Component_,  _hosted_ by another _VB-Project_ - which means it is a _Clobmne-Component_. If yes their code is compared and suggested for being updated if different.

## _SyncTargetWithSourceWb_

**pending implementation**<br>

### Aim, Purpose
Service for temporarily copied productive Workbooks for modifying the VB-Project while the productive Workbook remains in use. By this minimizing the down time of the productive Workbook to the time required for the "back-syncronization" of the modified VB-Project.


### Coverage, syncronization extent

| Element.   | Extent of synchronization |
| ---------- | ------------------------- |
|_References_| New are inserted, obsolete are removed |
|_Standard Modules_<br>_Class Modules_<br>_UserForms_| New are inserted, obsolete are removed and  changed code is updated |
|_Data Module_|**Workbook**: Code and module name synchronized<br>**Worksheet**: only partially (see [Worksheet synchronization](#worksheet-synchronization) and [Planning the release of a VB-Project modification](#planning-the-release-of-a-vb-project-modification))|
|_References_ | missing added and obsolete removed|
|_Form-Shapes_ | still in question |
|_ActiveX-Controls_| None. May be added in future |

### Worksheet synchronization
While the code of the development instance of a VB-Project is modified the productive instance will (can) continuously be used for data changes.  Because the user is able not only to change data but also the name and the position if a sheet, synchronization depends on an unchanged _CodeName_ as the only stable way to address a sheet in VBA code.

Design changes on a worksheet such as inserting new columns/rows may only be made in concert with the **release of a VB-Project** (_VB-Raw-Project_) to production:


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
pending description

## Planning the release of a VB-Project modification

pending description



[1]:https://gitcdn.link/repo/warbe-maker/VBA-Components-Management-Services/master/CompMan.xlsb
[2]:https://warbe-maker.github.io/warbe-maker.github.io/vba/excel/code/component/management/2021/02/05/Programatically-updating-Excel-VBA-code.html
