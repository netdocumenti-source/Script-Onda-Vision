Option Explicit

Sub OnDataLoaded(Entity)

End Sub

Sub OnInitGrid(Slave, Columns, RowActions, GridActions)
Dim stato


    If Slave = "RigheDocumento" Then
        Dim col_H  
        col_H = CreateColumn("NCP_H",dTypeInt)  
        col_H.Description = "H (mm)" 
        col_H.Width = 100  
        col_H.Order = 4  
        Columns.Add col_H   
        
        Dim col_L  
        col_L = CreateColumn("NCP_L",dTypeInt)  
        col_L.Description = "L (mm)"  
        col_L.Width = 100  
        col_L.Order = 5  
        Columns.Add col_L   

        'RowActions.Add CreateAction("RimuoviRiga", "Rimuovi riga")  
    End If
    
End Sub

Sub Idanagrafica_change(Entity)

Dim dtAnagraficheCliente
Dim Annotazione
Dim ChkAnnotazione
	
		'---Verifico se devo attivare il pop up delle notifiche
		If Entity.IdAnagrafica > 0 Then
			Set dtAnagraficheCliente = datadb.GetDataTable("SELECT NCPAnnotazioni,NCPNchkAnnotazioni FROM [STDAnagrafiche] WHERE IdAnagrafica = " & CCur(Entity.IdAnagrafica))
			Annotazione=GetField(dtAnagraficheCliente.Rows(0), "NCPAnnotazioni")
			ChkAnnotazione=GetField(dtAnagraficheCliente.Rows(0), "NCPNchkAnnotazioni")
			if ChkAnnotazione then
				MsgBox "Attenzione! " & Annotazione, btnOk
			end if
		end if

end sub


Sub AllowAction(action, Entity, allowObject)
   allowObject.Enable = true
End Sub

Sub BeforeAction(action, Entity, MessageList)

End Sub

Sub AfterAction(action, Entity)

End Sub
