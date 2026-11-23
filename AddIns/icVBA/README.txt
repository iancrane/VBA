icVBA Import and Export
=======================

Open icVBA.xlsm and the saved macro-enabled workbook you want to work on in
the same desktop Excel session. Excel must allow icVBA macros to run and have
"Trust access to the VBA project object model" enabled. The target VBA project
must be unlocked. Press Alt+F11, then Ctrl+G to open the Immediate Window.

In the named-workbook commands, use the exact filename of the open target
workbook, including its extension.

Import into a Workbook
----------------------

Enter the following command to update existing project-specific and shared
components from their source files.

    Call Application.Run("'icVBA.xlsm'!ImportNamedVBAProject", "MyWorkbook.xlsm", False)

Use True for shared components only. Shared-only is also the default when you
omit the last argument.

    Call Application.Run("'icVBA.xlsm'!ImportNamedVBAProject", "MyWorkbook.xlsm", True)

Import replaces existing components whose source files match by name and
type. Components without a matching file are left in place. To bring in a
new module, use File > Import File in the target project in the VBA editor.
Keep a form's .frx file beside its .frm file when it has one. Both export and
import skip document modules such as ThisWorkbook and Sheet1. Import leaves
the workbook open with unsaved changes so you can review them before saving.

Export a Workbook
-----------------

To export the active workbook, select it in Excel and click Export VBA Project
on the icVBA toolbar in the Add-ins tab.

To export a workbook by name, use the Immediate Window.

    Call Application.Run("'icVBA.xlsm'!ExportNamedVBAProject", "MyWorkbook.xlsm")

Export writes standard modules (.bas), class modules (.cls), and forms (.frm,
with any companion .frx files). Export overwrites existing source files with
the same names. Review the resulting changes in Git before committing them.

Source Folders
--------------

For an ordinary target workbook, icVBA uses the existing folder
MyWorkbook.xlsm.Source beside MyWorkbook.xlsm if present. Otherwise it uses
MyWorkbook.Source beside the workbook and creates it on export if needed.
The icVBA workbook itself always uses icVBA.xlsm.Source.

Shared modules are identified by the filename list in cVBAProjectFiles.IsShared.
They use AddIns\Shared beside the icVBA folder in this checkout. For other
workbooks, icVBA prefers an existing Shared folder under the applicable
T:\git\VBA\AddIns, T:\git\VBA\Distribution, or T:\git\VBA\Services root.
If T:\git\VBA\AddIns exists, LR, OfficeConnector, OfficeConnectorImport,
and icVisionsLogHelper use its LR, OC, OCI, and LogHelper folders respectively
for project-specific files.

icVBA reads and writes files directly on disk. To import source from a Git
checkout, make sure the resolved source folders point to that checkout.

Other Targets
-------------

Select the target workbook in Excel before using an active-workbook command.

    Call Application.Run("'icVBA.xlsm'!ImportCurrentVBAProject", False)
    Call Application.Run("'icVBA.xlsm'!ExportCurrentVBAProject")

Use these commands to target icVBA.xlsm itself. Self-import skips the modules
that implement import and export and the support classes they use.

    Call Application.Run("'icVBA.xlsm'!ImportThisVBAProject", False)
    Call Application.Run("'icVBA.xlsm'!ExportThisVBAProject")

The active-workbook and icVBA import commands also default to shared-only if
you omit the argument. The ExportLR/ImportLR, ExportOC/ImportOC, and
ExportQTFix/ImportQTFix shortcuts target LR.xls, OfficeConnector.xls, and
Excel 2016 QueryTable Fix.xlsm respectively. These import shortcuts use the
shared-only default.
