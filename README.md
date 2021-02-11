# VBA-Common-Components-Management
Methods for the management of Common VBA Components, i.e. update/sync, export changed, hosted, used Modules, Class Modules, and UserForms.
Available through a plug-in Workbook which has a development instance (this repo) which provides the method to replace/renew the active plug-in Workbook.<br>
See also [Programatically updating Excel VBA code][2]

# Services
## The _ExportChangedComponents_ service
Used with the _Workbook_Before_Save_ event it compares the code of any component in a _VB-Project_ with its last _Export File_ and re-exports it when different. The service is essential for _VB-Projects_ which host _Raw-Components_ in order to get them registered as available for other _VB-Projects_. Usage by any _VB-Project_ in a development status is appropriate as it is not only a code backup but also perfectly serves versioning - even when using [GitHub][]. Any _Component_ indicated a _hosted Raw-Component is registered as such with its _Export File_ as the main property.<br>
The service also checks a _Clone-Component_ modified within the VB-Project using it a offers updating the _Raw-Component_ in order to make the modification permanent. Testing the modification will be a task performed with the raw hosting project.

## The _UpdateRawClones_ service
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
