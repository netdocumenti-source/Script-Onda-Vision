Option Explicit
Dim SezionaleIva

Sub OnDataLoaded(Entity)
	'MsgBox "SEZIONALE IVA:" & Cstr(Entity.SezionaleIva), btnOk
	SezionaleIva=Entity.SezionaleIva
	
End Sub

'--- INTERCETTO IL CAMBIAMENTO DI ANAGRAFICA CLIENTE
'--- e quindi reimposto il Sezionale Iva con quello letto all'apertura del documento
Sub Idanagrafica_change(Entity)
	Entity.SezionaleIva=SezionaleIva

End Sub

Sub AllowAction(action, Entity, allowObject)
   If (action = "azione") Then
        allowObject.Enable = true
   End If
End Sub

Sub BeforeAction(action, Entity, MessageList)

End Sub

Sub AfterAction(action, Entity)

End Sub
