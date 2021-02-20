# VBA-Common-Components-Management
Methods for the management of Common VBA Components, i.e. update/sync, export changed, hosted, used Modules, Class Modules, and UserForms.
Available through a plug-in Workbook which has a development instance (this repo) which provides the method to replace/renew the active plug-in Workbook.<br>
See also [Programatically updating Excel VBA code][2]

@[:markdown](Disambiguation.md)

# Disambiguation
The terms below are not only those used in this post but also used with the implementation of the _Component Management_.

| Term             | Meaning                  |
|------------------|------------------------- |
|_Component_       | Generic _VB-Project_ term for a _Class Module_, a  _Data Module_, a _Standard Module_, or a _UserForm_  |
|_Common Component_| A _Component_ which is used by two or more VB-Projects |
| _Raw_,<br>_Raw-Component_ | The instance of a _Common Component_ which is regarded the developed, maintained and tested 'original', hosted in a dedicated _Raw-Host_ Workbook. |
| _Clone_,<br>_Clone-Component_,<br>_Raw-Clone_ | The copy of a _Raw- Component_ in a _VP-Project_ using it |
|_Clone-Project_ | A _VP-Project_ derived from a _Raw-Project_ |
|_Procedure_     | Any - Public or Private _Property_, _Sub_, or _Funtion_ of a _Component_. See also _Service_.
|_Raw-Host_.     | The Workbook/_VP-Project_ which hosts the _Raw-Component_ |
|_Raw-Project_   | A code-only _VP-Project_ of which all components are regarded _Raw-Components_. A _Raw-Project_ is kind of a template for the productive version of it. In contrast to a classic template it is the life-time raw code base for the productive _Clone-Project_.  The service and process of 'synchronizing' the productive (clone) code with the raw is part of the _Component Management_.|
|_Service_       | Generic term for any _Public Property_, _Public Sub_, or _Public Funtion_ of a _Component_ |
|_VB-Project_     | In the present case this term is used synonymously with Workbook |
| _Workbook-_, or<br>_VB-Project-Folder_ | A folder dedicated to a Workbook/VB-Project with all its Export Files and other project specific means. Such a folder is the equivalent of a Git-Repo-Clone (provided Git is used for the project's versioning which is recommendable |

# Services
## _ExportChangedComponents_
Used with the _Workbook_Before_Save_ event it compares the code of any component in a _VB-Project_ with its last _Export File_ and re-exports it when different. The service is essential for _VB-Projects_ which host _Raw-Components_ in order to get them registered as available for other _VB-Projects_. Usage by any _VB-Project_ in a development status is appropriate as it is not only a code backup but also perfectly serves versioning - even when using [GitHub][]. Any _Component_ indicated a _hosted Raw-Component is registered as such with its _Export File_ as the main property.<br>
The service also checks a _Clone-Component_ modified within the VB-Project using it a offers updating the _Raw-Component_ in order to make the modification permanent. Testing the modification will be a task performed with the raw hosting project.

## _UpdateRawClones_
The service is used with the _Workbook\_Open_ event. It checks each _Component_ for being known/registered as _Raw_  _hosted_ by another _VB-Project_. If yes, its code is compared with the _Raw's Export File and suggested for being updated if different.

## Installation
### Installation of the _CompMan_ Add-in
1. Download and open [CompManDev.xlsb][1]
2. Follow the instructions to identify a location for the Add-in - preferably a dedicated folder like ../CompMan/Add-in. The folder will hold the following files:
   - CompMan.cfg    ' the basic configuration
   - CompMan.xlam   ' the Add-in
   - HostedRaws.dat ' the specified raws hosted in any Workbook
   - RawHost.dat    ' the Workbooks which claim raws hosted
   
3. Follow the instructions to identify a 'serviced root'
4. Use the built-in Command button to run the _Renew_ service. It will:
   - ask to confirm or change the basic configuration
   - initially setup or subsequently renew the CompMan Add-in by saving a copy  of the development instance as Add-in (mind the fact that this is a multi-step process which may take some seconds)

Once the Add-in is established it will automatically be loaded with the first Workbook opened having it referenced. See the Usage below for further required preconditions.

### Installation for Workbooks/VB-Projects hosting raws or using raw clones
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
2. For a Workbook which hosts _Raw_Components_ specify them in the HOSTED_RAWS constant. If its more then one, have the component's names delimited with commas.

> ++**Be aware:**++ The Workbook component will be one of which the code cannot be updated by any means because it contains the code executed to perform the update. Thought this will only be relevant for Raw/Clone-VB-Projects which are yet not supported. However, as a consequence only calls to procedures provided with all arguments will remain in the Workbook component code and all the rest will be in a dedicated mWorkbook component.


[1]:https://gitcdn.link/repo/warbe-maker/VBA-Components-Management-Services/master/CompManDev.xlsb
[2]:https://warbe-maker.github.io/warbe-maker.github.io/vba/excel/code/component/management/2021/02/05/Programatically-updating-Excel-VBA-code.html