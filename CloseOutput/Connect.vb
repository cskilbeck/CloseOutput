Imports System
Imports Microsoft.VisualStudio.CommandBars
Imports Extensibility
Imports EnvDTE
Imports EnvDTE80

Public Class Connect
	
    Implements IDTExtensibility2
	Implements IDTCommandTarget

    Private _applicationObject As DTE2
    Private _addInInstance As AddIn

    Public Sub New()
    End Sub

    Public Sub OnConnection(ByVal application As Object, ByVal connectMode As ext_ConnectMode, ByVal addInInst As Object, ByRef custom As Array) Implements IDTExtensibility2.OnConnection
        _applicationObject = CType(application, DTE2)
        _addInInstance = CType(addInInst, AddIn)
        If connectMode = ext_ConnectMode.ext_cm_Startup Then
            Dim commands As Commands2 = CType(_applicationObject.Commands, Commands2)
            Dim command As Command = commands.AddNamedCommand2(_addInInstance, "CloseOutput", "CloseOutput", "CloseOutput", False, 0)
        End If
    End Sub

    Public Sub OnDisconnection(ByVal disconnectMode As ext_DisconnectMode, ByRef custom As Array) Implements IDTExtensibility2.OnDisconnection
    End Sub

    Public Sub OnAddInsUpdate(ByRef custom As Array) Implements IDTExtensibility2.OnAddInsUpdate
    End Sub

    Public Sub OnStartupComplete(ByRef custom As Array) Implements IDTExtensibility2.OnStartupComplete
    End Sub

    Public Sub OnBeginShutdown(ByRef custom As Array) Implements IDTExtensibility2.OnBeginShutdown
    End Sub
	
    Public Sub QueryStatus(ByVal commandName As String, ByVal neededText As vsCommandStatusTextWanted, ByRef status As vsCommandStatus, ByRef commandText As Object) Implements IDTCommandTarget.QueryStatus
        If neededText = vsCommandStatusTextWanted.vsCommandStatusTextWantedNone Then
            If commandName = "CloseOutput.Connect.CloseOutput" Then
                status = CType(vsCommandStatus.vsCommandStatusEnabled + vsCommandStatus.vsCommandStatusSupported, vsCommandStatus)
            Else
                status = vsCommandStatus.vsCommandStatusUnsupported
            End If
        End If
    End Sub

    Public Sub Exec(ByVal commandName As String, ByVal executeOption As vsCommandExecOption, ByRef varIn As Object, ByRef varOut As Object, ByRef handled As Boolean) Implements IDTCommandTarget.Exec
        handled = False
        If executeOption = vsCommandExecOption.vsCommandExecOptionDoDefault Then
            If commandName = "CloseOutput.Connect.CloseOutput" Then
                _applicationObject.Windows.Item(Constants.vsWindowKindFindSymbolResults).Visible = False
                _applicationObject.Windows.Item(Constants.vsWindowKindOutput).Visible = False
                _applicationObject.Windows.Item(WindowKinds.vsWindowKindErrorList).Visible = False
                _applicationObject.ExecuteCommand("Edit.SelectionCancel")
                handled = True
                Exit Sub
            End If
        End If
    End Sub
End Class
