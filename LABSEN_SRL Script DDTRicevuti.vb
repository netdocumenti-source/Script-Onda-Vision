Option Explicit

Sub OnDataLoaded(Entity)

End Sub

Sub AllowAction(action, Entity, allowObject)
   If (action = "azione") Then
        allowObject.Enable = true
   End If
End Sub

Sub BeforeAction(action, Entity, MessageList)
Dim riga, msg, Sconto1, Articolo     
Dim dt, row 

   'MsgBox "Segue Fattura:" & CStr(Entity.SegueFattura), btnOk
   'MsgBox "", btnYesNo
   If action = "Save" Then
   	
	   
    For Each riga in Entity.RigheDocumento
    	Articolo=riga.Articolo
    	Sconto1=riga.Sconto1
    	
    	if Cdbl(Sconto1)>0 then
	    	'MsgBox "Sconto1:" & CStr(riga.Sconto1) & "- Articolo:" & Cstr(riga.Articolo), btnOk
			Set dt = dataDB.GetDataTable("SELECT PercSconto1Acq,Codart FROM MAGAnagraficaArticoli WHERE Codart = '" & Articolo & "'")  
			
			If (dt.Rows.Count > 0) Then  
			 	dt.Rows(0)("PercSconto1Acq") = Sconto1  
			 	dataDB.UpdateDataTable dt, "SELECT PercSconto1Acq,Codart FROM MAGAnagraficaArticoli"  
			End If  
		end if    	
    	
    	'if (CStr(riga.Causale1)="CPT") or (CStr(riga.Causale1)="ACQ") then
    	'	Entity.SegueFattura=True	
    	'end if
    	
		'if (riga.RigaChiusaManuale=0) And CStr(riga.Causale1)="SMLT" And CStr(riga.TipoRiga)="Articolo" then
  			'And (riga.Causale1<>"CPT") And (riga.TipoRiga=1)
  			
  			'MsgBox "riga.TipoRiga: " & CStr(riga.TipoRiga), btnOk
  			'MsgBox "riga.Causale1: " & CStr(riga.Causale1), btnOk
  			
  			'--- Questo stub di codice crea un messaggio standard di OndaVision
            'set msg = CreateMessage("Impostare la chiusura della riga " & riga.IdRiga & "-" & riga.RigaChiusaManuale, msgError)    
            'MessageList.Add msg    

			'set colChiusa = GetColumn("RigheDocumento.RigaChiusaManuale")  
		    'colChiusa = CInt(1)  


         '   riga.RigaChiusaManuale=true
            
            
		'end if
    Next    

   End If          

End Sub

Sub AfterAction(action, Entity)

End Sub
