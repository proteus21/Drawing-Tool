Imports Inventor
Imports System.Runtime.InteropServices
Imports Microsoft.Win32

Namespace PrintPDFs
    <ProgIdAttribute("PrintPDFs.StandardAddInServer"), _
    GuidAttribute("1b40ff78-773a-40b0-a374-ae9e155897e3")> _
    Public Class StandardAddInServer
        Implements Inventor.ApplicationAddInServer
        Friend Shared Instance As StandardAddInServer

        ' Inventor application object.
        ' Inventor application object.

        Public ThisApplication As Inventor.Application
        Private WithEvents m_PrintPDFsButtonDef As ButtonDefinition
        Private WithEvents m_InputEvents As UserInputEvents
        Private WithEvents m_EditFieldTextButtonDef As ButtonDefinition

#Region "ApplicationAddInServer Members"

        Public Sub Activate(ByVal addInSiteObject As Inventor.ApplicationAddInSite, ByVal firstTime As Boolean) Implements Inventor.ApplicationAddInServer.Activate

            ' This method is called by Inventor when it loads the AddIn.
            ' The AddInSiteObject provides access to the Inventor Application object.
            ' The FirstTime flag indicates if the AddIn is loaded for the first time.

            ' TODO:  Add ApplicationAddInServer.Activate implementation.
            ' e.g. event initialization, command creation etc.

            ' Initialize AddIn members.
            Instance = Me
            ThisApplication = addInSiteObject.Application
            m_InputEvents = ThisApplication.CommandManager.UserInputEvents

            Dim controldefs As ControlDefinitions
            controldefs = ThisApplication.CommandManager.ControlDefinitions

            m_PrintPDFsButtonDef = controldefs.AddButtonDefinition( _
                                        "Print PDF Files", _
                                        "PrintCustomPDFs", _
                                        CommandTypesEnum.kQueryOnlyCmdType, _
                                        "{1b40ff78-773a-40b0-a374-ae9e155897e3}", _
                                        "Print PDFs to agreed standards!", _
                                        "Print PDFs to Company and client standards")
            If firstTime Then
                ' adds the button to the drawing annotation command bar.
                Dim DWGcommandbar As Inventor.CommandBar
                DWGcommandbar = ThisApplication.UserInterfaceManager.CommandBars.Item( _
                                "DLxDrawingAnnotationPanelCmdBar")
                DWGcommandbar.Controls.AddButton(m_PrintPDFsButtonDef)
                ' Also creates a new command bar (toolbar) and makes it visible.
                Dim Commandbars As CommandBars
                Commandbars = ThisApplication.UserInterfaceManager.CommandBars

                Dim CustomCommandbar As CommandBar
                CustomCommandbar = Commandbars.Add("CustomTools", _
                                                 "CustomAddInMacros", , _
                                                 "{1b40ff78-773a-40b0-a374-ae9e155897e3}")

                CustomCommandbar.Visible = True
                CustomCommandbar.Controls.AddButton(m_PrintPDFsButtonDef)
            End If
        End Sub

        Public Sub Deactivate() Implements Inventor.ApplicationAddInServer.Deactivate

            ' This method is called by Inventor when the AddIn is unloaded.
            ' The AddIn will be unloaded either manually by the user or
            ' when the Inventor session is terminated.

            ' TODO:  Add ApplicationAddInServer.Deactivate implementation

            ' Release objects.
            Marshal.ReleaseComObject(ThisApplication)
            ThisApplication = Nothing

            System.GC.WaitForPendingFinalizers()
            System.GC.Collect()

        End Sub

        Public ReadOnly Property Automation() As Object Implements Inventor.ApplicationAddInServer.Automation

            ' This property is provided to allow the AddIn to expose an API 
            ' of its own to other programs. Typically, this  would be done by
            ' implementing the AddIn's API interface in a class and returning 
            ' that class object through this property.

            Get
                Return Nothing
            End Get

        End Property

        Public Sub ExecuteCommand(ByVal commandID As Integer) Implements Inventor.ApplicationAddInServer.ExecuteCommand

            ' Note:this method is now obsolete, you should use the 
            ' ControlDefinition functionality for implementing commands.

        End Sub
        Private Sub m_PrintPDFsButtonDef_OnExecute(ByVal Context As Inventor.NameValueMap) Handles m_PrintPDFsButtonDef.OnExecute
            Dim PrintForm As New Form1
            PrintForm.Show(New WindowWrapper(ThisApplication.MainFrameHWND))
        End Sub

#End Region

#Region "COM Registration"

        ' Registers this class as an AddIn for Inventor.
        ' This function is called when the assembly is registered for COM.
        <ComRegisterFunctionAttribute()> _
        Public Shared Sub Register(ByVal t As Type)

            Dim clssRoot As RegistryKey = Registry.ClassesRoot
            Dim clsid As RegistryKey = Nothing
            Dim subKey As RegistryKey = Nothing

            Try
                clsid = clssRoot.CreateSubKey("CLSID\" + AddInGuid(t))
                clsid.SetValue(Nothing, "PrintPDFs")
                subKey = clsid.CreateSubKey("Implemented Categories\{39AD2B5C-7A29-11D6-8E0A-0010B541CAA8}")
                subKey.Close()

                subKey = clsid.CreateSubKey("Settings")
                subKey.SetValue("AddInType", "Standard")
                subKey.SetValue("LoadOnStartUp", "1")

                'subKey.SetValue("SupportedSoftwareVersionLessThan", "")
                subKey.SetValue("SupportedSoftwareVersionGreaterThan", "12..")
                'subKey.SetValue("SupportedSoftwareVersionEqualTo", "")
                'subKey.SetValue("SupportedSoftwareVersionNotEqualTo", "")
                'subKey.SetValue("Hidden", "0")
                'subKey.SetValue("UserUnloadable", "1")
                subKey.SetValue("Version", 0)
                subKey.Close()

                subKey = clsid.CreateSubKey("Description")
                subKey.SetValue(Nothing, "PrintPDFs")

            Catch ex As Exception
                System.Diagnostics.Trace.Assert(False)
            Finally
                If Not subKey Is Nothing Then subKey.Close()
                If Not clsid Is Nothing Then clsid.Close()
                If Not clssRoot Is Nothing Then clssRoot.Close()
            End Try

        End Sub

        ' Unregisters this class as an AddIn for Inventor.
        ' This function is called when the assembly is unregistered.
        <ComUnregisterFunctionAttribute()> _
        Public Shared Sub Unregister(ByVal t As Type)

            Dim clssRoot As RegistryKey = Registry.ClassesRoot
            Dim clsid As RegistryKey = Nothing

            Try
                clssRoot = Microsoft.Win32.Registry.ClassesRoot
                clsid = clssRoot.OpenSubKey("CLSID\" + AddInGuid(t), True)
                clsid.SetValue(Nothing, "")
                clsid.DeleteSubKeyTree("Implemented Categories\{39AD2B5C-7A29-11D6-8E0A-0010B541CAA8}")
                clsid.DeleteSubKeyTree("Settings")
                clsid.DeleteSubKeyTree("Description")
            Catch
            Finally
                If Not clsid Is Nothing Then clsid.Close()
                If Not clssRoot Is Nothing Then clssRoot.Close()
            End Try

        End Sub

        ' This property uses reflection to get the value for the GuidAttribute attached to the class.
        Public Shared ReadOnly Property AddInGuid(ByVal t As Type) As String
            Get
                Dim guid As String = ""
                Try
                    Dim customAttributes() As Object = t.GetCustomAttributes(GetType(GuidAttribute), False)
                    Dim guidAttribute As GuidAttribute = CType(customAttributes(0), GuidAttribute)
                    guid = "{" + guidAttribute.Value.ToString() + "}"
                Finally
                    AddInGuid = guid
                End Try
            End Get
        End Property

#End Region
#Region "hWnd Wrapper Class"
        ' This class is used to wrap a Win32 hWnd as a .Net IWind32Window class. 
        ' This is primarily used for parenting a dialog to the Inventor window.
        '
        ' For example: 
        ' myForm.Show(New WindowWrapper(m_inventorApplication.MainFrameHWND))
        '
        Public Class WindowWrapper
            Implements System.Windows.Forms.IWin32Window
            Private _hwnd As IntPtr

            Public Sub New(ByVal handle As IntPtr)
                _hwnd = handle
            End Sub

            Public ReadOnly Property Handle() As IntPtr _
              Implements System.Windows.Forms.IWin32Window.Handle
                Get
                    Return _hwnd
                End Get
            End Property
        End Class
#End Region
    End Class

End Namespace


