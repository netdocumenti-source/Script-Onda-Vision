Option Explicit
Dim IdListinoAnagrafica

Sub OnDataLoaded(Entity)
Dim dtListinoCliente

	'Identifico IdListino se l'l'anagrafica è già assegnata
	If Entity.IdAnagrafica > 0 Then
		Set dtListinoCliente = datadb.GetDataTable("SELECT IdListino FROM [STDListinoTesta] WHERE IdAnagrafica = " & CCur(Entity.IdAnagrafica))
		if dtListinoCliente.Rows.Count > 0 then
			IdListinoAnagrafica=GetField(dtListinoCliente.Rows(0), "IdListino")
		
		else
			IdListinoAnagrafica=0
			
		end if
		'MsgBox "Listino identificato :" & IdListinoAnagrafica, btnOk
			
	Else
		'MsgBox "Anagrafica Cliente non impostata " & Entity.IdAnagrafica, btnOk
			
	END IF
	
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

		'GridActions.Add CreateAction("ImportaArticolo", "Importa articolo")  
        'GridActions.Add CreateAction("ImportaArticoloXML", "Importa articoli da XML")  
        'RowActions.Add CreateAction("RimuoviRiga", "Rimuovi riga")  
    End If
    
End Sub

Sub Idanagrafica_change(Entity)
Dim dtListinoCliente
Dim dtAnagraficheCliente
Dim Annotazione
Dim ChkAnnotazione

		'--- Ricerco il Prezzo del Listino Cliente
		If Entity.IdAnagrafica > 0 Then
			Set dtListinoCliente = datadb.GetDataTable("SELECT IdListino FROM [STDListinoTesta] WHERE IdAnagrafica = " & CCur(Entity.IdAnagrafica))	
			if dtListinoCliente.Rows.Count > 0 then
				IdListinoAnagrafica=GetField(dtListinoCliente.Rows(0), "IdListino")
				'MsgBox "Listino identificato :" & IdListinoAnagrafica, btnOk
				
			else
				IdListinoAnagrafica=0
				'MsgBox "Nessun Listino identificato ", btnOk
			
			end if
			
		Else
			MsgBox "Anagrafica Cliente non impostata " & Entity.IdAnagrafica, btnOk
			
		End If
	
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


'Intercetto al tab il cambio di valore del campo NCP_H
Sub RigheDocumentoNCP_H_Change(Entity,riga)
	Dim dtListinoClienteArticolo
	Dim dtPrezzoVenditaArticolo
	Dim Codart
	Dim Prezzo
	Dim Mq
	Dim UM

	Codart=riga.Articolo
	If Instr(Codart,"'") then
		Codart=Replace(codart,"'","''")
	end if
	'MsgBox "Codart :" & Codart, btnOk
	UM=riga.UnitaMisura
	
	
	select case Um
	
		case "PZ"
		
		
		case "ML"
			riga.Extensions("NCP_H")=0
		
		case "MQ"
			
			'-----------------------------------------------------------------------------
			'Se esiste il listino dell'Anagrafica allora vado a cercare il prezzo dell'articolo in STDListinoRighe.
			'Se l'articolo però non è stato inserito nel listino, allora vado a cercare il prezzo in MAGAnagraficaArticoli.
			'Se non esiste neanche in anagrafica, imposto il prezzo = 0
			'-----------------------------------------------------------------------------
			If IdListinoAnagrafica>0 Then
				Set dtListinoClienteArticolo = datadb.GetDataTable("SELECT ValoreUnitario FROM [STDListinoRighe] WHERE IdListino = " & CCur(IdListinoAnagrafica) & " and CodArt='" & CStr(Codart) & "'")
				If (dtListinoClienteArticolo.Rows.Count > 0) Then  
					Prezzo=GetField(dtListinoClienteArticolo.Rows(0), "ValoreUnitario")
					'MsgBox "Valore Unitario identificato :" & Prezzo, btnOk
			
				else
					Set dtPrezzoVenditaArticolo=datadb.GetDataTable("SELECT PREZZOVENDITA FROM [MAGAnagraficaArticoli] WHERE CODART='" & CStr(Codart) & "'")
					If (dtPrezzoVenditaArticolo.Rows.Count > 0) Then  
						Prezzo=GetField(dtPrezzoVenditaArticolo.Rows(0), "PrezzoVendita")
					    Mq=(CCur(riga.Extensions("NCP_H"))*CCur(riga.Extensions("NCP_L")))/1000000
					    Prezzo=CCur(Prezzo)*CCur(Mq)
						'MsgBox "Valore Unitario non identificato per il listino " & IdListinoAnagrafica, btnOk
					else
						Prezzo=0
					end if
					
				end if

		    	riga.ValoreUnitario=CCur(Prezzo)
				
				'--- Se Prezzo>0 allora effettuo il ricalcolo 
			    if CCur(Prezzo)>0 then
			    	Call Entity.RicalcolaRighe()
			    	
			    end if


			Else
				'---------------------------------------------------------------------------------------------------------------------------
				'Caso in cui il Cliente non ha un listino e pertanto vado direttamente nell'anagrafica dell'articolo a vedere il prezzo di vendita
				'---------------------------------------------------------------------------------------------------------------------------
				Set dtPrezzoVenditaArticolo=datadb.GetDataTable("SELECT PREZZOVENDITA FROM [MAGAnagraficaArticoli] WHERE CODART='" & CStr(Codart) & "'")
				If (dtPrezzoVenditaArticolo.Rows.Count > 0) Then  
					Prezzo=GetField(dtPrezzoVenditaArticolo.Rows(0), "PrezzoVendita")
				    Mq=(CCur(riga.Extensions("NCP_H"))*CCur(riga.Extensions("NCP_L")))/1000000
				    Prezzo=CCur(Prezzo)*CCur(Mq)
					'MsgBox "Valore Unitario non identificato per il listino " & IdListinoAnagrafica, btnOk
				else
					Prezzo=0
					
				end if
				
				riga.ValoreUnitario=CCur(Prezzo)
			    if CCur(Prezzo)>0 then			    	
			    	Call Entity.RicalcolaRighe()
			    	
			    end if
					
			End If
			
	end Select
	
	
End Sub  



'Intercetto al tab il cambio di valore del campo NCP_L
Sub RigheDocumentoNCP_L_Change(Entity,riga)
	Dim dtListinoClienteArticolo
	Dim dtPrezzoVenditaArticolo
	Dim Codart
	Dim Prezzo
	Dim Mq
	Dim Um


	Codart=riga.Articolo
	If Instr(Codart,"'") then
		Codart=Replace(codart,"'","''")
	end if
	
	UM=riga.UnitaMisura
	'MsgBox "UM identificato :" & UM, btnOk
	
	select case Um
	
		case "PZ"
		
		
		case "ML"
			'MsgBox "UM identificato :" & UM, btnOk
			'MsgBox "IdListinoAnagrafica :" & IdListinoAnagrafica, btnOk
			riga.Extensions("NCP_H")=0
			If IdListinoAnagrafica>0 Then
				Set dtListinoClienteArticolo = datadb.GetDataTable("SELECT ValoreUnitario FROM [STDListinoRighe] WHERE IdListino = " & CCur(IdListinoAnagrafica) & " and CodArt='" & CStr(Codart) & "'")
				If (dtListinoClienteArticolo.Rows.Count > 0) Then  
					Prezzo=GetField(dtListinoClienteArticolo.Rows(0), "ValoreUnitario")
					'MsgBox "Valore Unitario identificato :" & Prezzo, btnOk
					Prezzo=CCur(Prezzo)*CCur(riga.Extensions("NCP_L"))
					Prezzo=CCur(Prezzo)/1000
								
				else
					Set dtPrezzoVenditaArticolo=datadb.GetDataTable("SELECT PREZZOVENDITA FROM [MAGAnagraficaArticoli] WHERE CODART='" & CStr(Codart) & "'")
					If (dtPrezzoVenditaArticolo.Rows.Count > 0) Then  
						Prezzo=GetField(dtPrezzoVenditaArticolo.Rows(0), "PrezzoVendita")
						'MsgBox "Valore Unitario non identificato per il listino " & IdListinoAnagrafica, btnOk
						Prezzo=CCur(Prezzo)*CCur(riga.Extensions("NCP_L"))
						Prezzo=CCur(Prezzo)/1000
						
					else
						Prezzo=0
					end if
				
				end if

				'MsgBox "Valore Unitario identificato :" & Prezzo, btnOk

				riga.ValoreUnitario=CCur(Prezzo)
				if Prezzo>0 then
			    	Call Entity.RicalcolaRighe()
			    	
			    end if
				
			Else
				'MsgBox "ML. Valore Unitario non identificato per il listino " & IdListinoAnagrafica, btnOk
				Set dtPrezzoVenditaArticolo=datadb.GetDataTable("SELECT PREZZOVENDITA FROM [MAGAnagraficaArticoli] WHERE CODART='" & CStr(Codart) & "'")
				If (dtPrezzoVenditaArticolo.Rows.Count > 0) Then  
					Prezzo=GetField(dtPrezzoVenditaArticolo.Rows(0), "PrezzoVendita")
					'MsgBox "Valore Unitario non identificato per il listino " & IdListinoAnagrafica, btnOk
					Prezzo=CCur(Prezzo)*CCur(riga.Extensions("NCP_L"))
					Prezzo=CCur(Prezzo)/1000
					
				else
					Prezzo=0
				end if

				'MsgBox "Valore Unitario identificato :" & Prezzo, btnOk

		    	riga.ValoreUnitario=CCur(Prezzo)
				if Prezzo>0 then
			    	Call Entity.RicalcolaRighe()
			    	
			    end if

				
			End If
			

					
		case "MQ"
			If IdListinoAnagrafica>0 Then
				Set dtListinoClienteArticolo = datadb.GetDataTable("SELECT ValoreUnitario FROM [STDListinoRighe] WHERE IdListino = " & CCur(IdListinoAnagrafica) & " and CodArt='" & CStr(Codart) & "'")
				If (dtListinoClienteArticolo.Rows.Count > 0) Then  
					Prezzo=GetField(dtListinoClienteArticolo.Rows(0), "ValoreUnitario")
					'MsgBox "Valore Unitario identificato :" & Prezzo, btnOk
			
				else
					Set dtPrezzoVenditaArticolo=datadb.GetDataTable("SELECT PREZZOVENDITA FROM [MAGAnagraficaArticoli] WHERE CODART='" & CStr(Codart) & "'")
					If (dtPrezzoVenditaArticolo.Rows.Count > 0) Then  
						Prezzo=GetField(dtPrezzoVenditaArticolo.Rows(0), "PrezzoVendita")
						'MsgBox "Valore Unitario non identificato per il listino " & IdListinoAnagrafica, btnOk
						
					else
						Prezzo=0
					end if


				end if
				'MsgBox "Valore Unitario identificato :" & Prezzo, btnOk

				if CCur(riga.Extensions("NCP_H"))>0 then
				    'Prezzo=CCur(riga.ValoreUnitario)
				    'MsgBox "Prezzo = " & Prezzo & "*" & CCur(riga.Extensions("NCP_H")) & "*" & CCur(riga.Extensions("NCP_L")), btnOk
				    Mq=(CCur(riga.Extensions("NCP_H"))*CCur(riga.Extensions("NCP_L")))/1000000
				    'MsgBox "Mq=" & CCur(Mq), btnOk
				    Prezzo=CCur(Prezzo)*CCur(Mq)
			
				else
					MsgBox "Attenzione!! Inserire un valore di H>0.", btnOk					
				
				end if
			    
			    riga.ValoreUnitario=CCur(Prezzo)
			    if Prezzo>0 then			    	
			    	Call Entity.RicalcolaRighe()
			    	
			    end if
				
			Else
				'MsgBox "MQ. Valore Unitario non identificato per il listino " & IdListinoAnagrafica, btnOk

				if CCur(riga.Extensions("NCP_H"))>0 then
					Set dtPrezzoVenditaArticolo=datadb.GetDataTable("SELECT PREZZOVENDITA FROM [MAGAnagraficaArticoli] WHERE CODART='" & CStr(Codart) & "'")
					If (dtPrezzoVenditaArticolo.Rows.Count > 0) Then  
						Prezzo=GetField(dtPrezzoVenditaArticolo.Rows(0), "PrezzoVendita")
						'MsgBox "Valore Unitario non identificato per il listino " & IdListinoAnagrafica, btnOk
						
					else
						Prezzo=0
					end if

				    'MsgBox "Prezzo = " & Prezzo & "*" & CCur(riga.Extensions("NCP_H")) & "*" & CCur(riga.Extensions("NCP_L")), btnOk
				    Mq=(CCur(riga.Extensions("NCP_H"))*CCur(riga.Extensions("NCP_L")))/1000000
				    'MsgBox "Mq=" & CCur(Mq), btnOk
				    Prezzo=CCur(Prezzo)*CCur(Mq)
			
				else
				    'Prezzo=CCur(Prezzo)*CCur(riga.Extensions("NCP_L"))
					MsgBox "Attenzione!! Inserire un valore di H>0.", btnOk
				end if
			    
			    riga.ValoreUnitario=CCur(Prezzo)
			    if Prezzo>0 then			    	
			    	Call Entity.RicalcolaRighe()
			    	
			    end if
				
			End If
						
		
		
	end Select
	
	
End Sub  



Sub AllowAction(action, Entity, allowObject)
   allowObject.Enable = true
End Sub

Sub BeforeAction(action, Entity, MessageList)

End Sub

Sub AfterAction(action, Entity)

End Sub



