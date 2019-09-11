Attribute VB_Name = "modMenu"
Option Explicit

Private MnuEvt      As clsVBECmdHandler
Private EvtHandlers As New Collection


Public Sub auto_open()

    Call CreateVBEMenu
    '// Call CreateXLMenu

End Sub


Public Sub auto_close()

    Call RemoveVBEMenu

End Sub


Private Sub CreateVBEMenu()
    
    Dim objToolbar  As CommandBar
    Dim objToolbarItem As Object
    
    Dim objMenu     As CommandBarPopup
    Dim objMenuItem As Object
    
    On Error Resume Next
    Application.VBE.CommandBars("VBA Utilities").Delete
    On Error GoTo 0

    Set objToolbar = Application.VBE.CommandBars.Add("VBA Utilities", , False, False)
    objToolbar.Visible = True
    With objToolbar
        Set objToolbarItem = .Controls.Add(Type:=msoControlButton)
        objToolbarItem.Caption = "Make Config File"
        objToolbarItem.FaceId = 6370
        objToolbarItem.OnAction = "MakeConfigFile"
        Call MenuEvents(objToolbarItem)
        
        Set objToolbarItem = .Controls.Add(Type:=msoControlButton)
        objToolbarItem.Caption = "Export"
        objToolbarItem.FaceId = 6354
        objToolbarItem.OnAction = "Export"
        Call MenuEvents(objToolbarItem)
        
        Set objToolbarItem = .Controls.Add(Type:=msoControlButton)
        objToolbarItem.Caption = "Import"
        objToolbarItem.FaceId = 3177
        objToolbarItem.OnAction = "Import"
        Call MenuEvents(objToolbarItem)
        
    End With
    
    Set objMenu = Application.VBE.CommandBars(1).Controls.Add(Type:=msoControlPopup)
    With objMenu
        objMenu.Caption = "E&xport for VCS"

        Set objMenuItem = .Controls.Add(Type:=msoControlButton)
        objMenuItem.OnAction = "MakeConfigFile"
        Call MenuEvents(objMenuItem)
        objMenuItem.Caption = "&Make Config File"

        Set objMenuItem = .Controls.Add(Type:=msoControlButton)
        objMenuItem.OnAction = "Import"
        Call MenuEvents(objMenuItem)
        objMenuItem.Caption = "&Import"

        Set objMenuItem = .Controls.Add(Type:=msoControlButton)
        objMenuItem.OnAction = "Export"
        Call MenuEvents(objMenuItem)
        objMenuItem.Caption = "&Export"

    End With

    Set objMenuItem = Nothing
    Set objMenu = Nothing

End Sub


Private Sub MenuEvents(ByVal objMenuItem As Object)

    Set MnuEvt = New clsVBECmdHandler
    Set MnuEvt.EvtHandler = Application.VBE.Events.CommandBarEvents(objMenuItem)
    EvtHandlers.Add MnuEvt

End Sub


Private Sub CreateXLMenu()

    MenuBars(xlWorksheet).Menus.Add Caption:="E&xport for VCS"
    With MenuBars(xlWorksheet).Menus("Export for VCS").MenuItems
        .Add Caption:="&Make Config File", _
             OnAction:="MakeConfigFile"
        .Add Caption:="&Import", _
             OnAction:="Import"
        .Add Caption:="&Export", _
             OnAction:="Export"
    End With

End Sub


Private Sub RemoveVBEMenu()

    On Error Resume Next

    Application.VBE.CommandBars(1).Controls("Export for VCS").Delete

    '// Clear the EvtHandlers collection if there is anything in it
    While EvtHandlers.Count > 0
        EvtHandlers.Remove 1
    Wend

    Set EvtHandlers = Nothing
    Set MnuEvt = Nothing

    Application.CommandBars("Worksheet Menu Bar").Controls("E&xport for VCS").Delete
    On Error GoTo 0

End Sub

'// RibUI callbacks
Public Sub btnMakeConfig_onAction(control As IRibbonControl)
    Call MakeConfigFile
End Sub
Public Sub btnExport_onAction(control As IRibbonControl)
    Call Export
End Sub
Public Sub btnImport_onAction(control As IRibbonControl)
    Call Import
End Sub

