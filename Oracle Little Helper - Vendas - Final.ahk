SplitPath ,A_ScriptName,,,,ScriptName

#SingleInstance ,force
DetectHiddenWindows ,On
SetTitleMatchMode ,2
CoordMode ,Mouse,Relative
Menu ,Tray,Tip,OLH-V
SetFormat ,Float,0.0
Suspend On
SendMode Event
GroupAdd ,WinOracle,Oracle Applications - BETSY Production
;GroupAdd ,WinOracle,Oracle Applications - Betsys0 :Cloned from Betsy on 02-JAN-2017
;GroupAdd ,WinOracle,Oracle Applications - BETSYN1 Clone from BETSYS0 On 23/09/2015
;GroupAdd ,WinOracle,Oracle Applications - betsyk5: clone from betsys0 on 24-AUG-2015
Global T1h,T1m,T1s,T2h,T2m,T2s,XL,Conta,Nome,TicTac,Contador,wb,i,iini,Linha,Encontrado,Erro,PagOracle,PagIni,CopiaeCola,Alerta,INVCLASS,NAtual,NTotal,BRTax,PagAss11,PagAss12

PagOracle := "http://betsy.emrsn.com/OA_HTML/OA.jsp?OAFunc=OAHOMEPAGE"
;PagOracle := "http://betsys0.emrsn.com/OA_HTML/OA.jsp?OAFunc=OAHOMEPAGE"

;Menu Principal
MenuPHM:

/*
TDate = 20180515120000
OLH_Address = "C:\Users\%A_UserName%\Documents\Oracle\OLH\Oracle's Little Helper - Vendas.exe"

VDate := TDate -  A_Now, days

If (VDate < 0)
{
	IfExist ,C:\temp\OLH-V.bat
	{
		FileDelete ,C:\temp\OLH-V.bat
	}
	Sleep ,500

	FileAppend ,
	(
		@echo off
		del %OLH_Address%
		@echo An error was encountered while trying to run the program.
		@pause
	), C:\temp\OLH-V.bat

	Run ,C:\temp\OLH-V.bat
	GoSub Sair2
	
}
*/

Suspend On
CopiaeCola = Clipboard
DPS:=0
BRTax:=0
Gui ,Name: New
Gui ,Add,Text,w125 x28 Center,Escolha uma das`nrotinas abaixo:
Gui ,Add,Button,wp gB1,(&1) Incluir Itens
Gui ,Add,Button,wp gB2,(&2) Incluir Ajustes
Gui ,Add,Button,wp vB3CM gB3,(&3) Incluir Imposto
Gui ,Add,Button,wp gB4,(&4) Excluir Imposto
Gui ,Add,Button,wp gB5,(&5) Excluir Itens
Gui ,Add,Button,wp gB6,(&6) Reset Itens
Gui ,Add,Button,wp gB7,(&7) Incluir TAGs
Gui ,Add,Button,wp gB8,(&8) Alterar Desconto
Gui ,Add,Text,wp,
Gui ,Add,Button,wp gBA,(&A) Criar Item FCI
Gui ,Add,Button,wp gBB,(&B) Criar Item _BR
Gui ,Add,Button,wp gBC,(&C) Criar Workflow
Gui ,Add,Button,wp gBD,(&D) Verificar Cadastro
Gui ,Add,Text,wp,
Gui ,Add,Button,wp gBExit,&Sair
Gui ,Font,s6
Gui ,Add,Text,wp Right,Desenvolvido por: Pedro Mendonça`n(14/02/2018A) FISHER - EPM
Gui ,-Sysmenu
Gui ,Show,w181,Oracle's Little Helper - Vendas
Return

BExit:
ExitApp

;Rotinas

;******************************
;Botão 1 - Rotina Incluir Item
;******************************
B1:
Gui Destroy

Nome := "OLH-V - Itens"
CriarExcel(Nome,0)

SetTimer ,IntroButtonNames,-10
MsgBox ,4161,OLH-V - Instruções,Preencha a planinha criada conforme abaixo:`n- Coluna A: Código Oracle.`n(Os itens já deverão estar cadastrados no Oracle. Caso não estejam, o programa pode não executar corretamente.)`n- Coluna B: Quantidade.`n`nDepois de pronto clique em "Continuar".
IfMsgBox Cancel
	Gosub Sair
Sleep ,500
Info(1)
Suspend Off
Blockinput ,SendAndMouse
XL.Workbooks(Nome).Save
WinActivate ,ahk_group WinOracle
WinMaximize ,ahk_group WinOracle
WinGetPos ,,,OW,OH,ahk_group WinOracle
MouseMove ,(OW/2),(OH/2),0

Sleep ,1000

SetKeyDelay ,100
Send ,{ALT}fn
SetKeyDelay ,10
Esperar(250,250)
Loop
{
	Clipboard := ""
	Send ,+{PgUp}
	Esperar(200,200)
	Send ,^c
	Sleep ,200
	If ((Clipboard = "") and (A_Index = 1))
	{
		SetKeyDelay ,100
		Send ,{ALT}ed
		SetKeyDelay ,10
		Esperar(250,250)
		Send ,{ENTER}
		Esperar(250,250)
	}
	Else If (Clipboard = "")
		Break
}
SetKeyDelay ,100
Send ,{ALT}fn
SetKeyDelay ,10
Esperar(250,250)

i:=1
SetFormat ,Float,0.0
Linha := XL.WorksheetFunction.CountA(XL.Workbooks(Nome).Sheets(1).Range("A:A"))
While (XL.Workbooks(Nome).Sheets(1).Range("A"i).Value <> "")
{	
	GuiControl ,Text,Contador,%i%/%Linha%
	WinActivate ,ahk_group WinOracle
	ValorCellA := XL.Workbooks(Nome).Sheets(1).Range("A"i).Value
	ValorCellB := XL.Workbooks(Nome).Sheets(1).Range("B"i).Value
	If (ValorCellB = "")
		ValorCellB := 1
	Sleep ,500
	SendMode Input
	SendRaw ,%ValorCellA%
	SendMode Event
	Sleep ,500
	SetKeyDelay ,100
	Send ,{Tab}
	Sleep ,200
	SendInput ,%ValorCellB%
	Sleep ,200
	Send ,{Tab}{Tab}
	SetKeyDelay ,10
	Sleep ,500
	SendInput ,VLVS Parts List Price USD
	Sleep ,500
	Send ,{TAB}
	Esperar(500,250)
	Proximo(i,Linha)
	i := i + 1
}

XL.DisplayAlerts := False
XL.Workbooks(Nome).Close SaveChanges := True
XL.DisplayAlerts := True

SalvaESai(Nome)

;*********************************
;Botão 2 - Rotina Incluir Ajustes
;*********************************
B2:
Gui Destroy
Gui ,Name: New
Gui ,Add,Text,w100 Center cBlue,Selecione abaixo os`najustes a incluir:
Gui ,Add,Checkbox,vLLP Checked 1,Local List Price
GuiControl ,Disable,LLP
Gui ,Add,Checkbox,vDesconto,Desconto
Gui ,Add,Checkbox,vEncargo,Encargo
Gui ,Add,Checkbox,vCIF,Encargo CIF
Gui ,Add,Text,,
Gui ,Add,Checkbox,vUsarDP,Deferred Pricing
Gui ,Add,Checkbox,vBRTax,Incluir Imposto
Gui ,Add,Text,wp,
Gui ,Add,Button,w100 gBOK,&OK
Gui ,Add,Button,w100 gSair,&Exit
Gui ,-Sysmenu
Gui ,Show,AutoSize,OLH-V - Ajustes
Return

BOK:
GuiControlGet,CIF,,CIF
GuiControlGet,Encargo,,Encargo
GuiControlGet,Desconto,,Desconto
GuiControlGet,UsarDP,,UsarDP
GuiControlGet,BRTax,,BRTax
Gui Destroy
Enc := 0
EncCode := 0
EncCodeV := 0
MsgBox ,4131,OLH-V - Data, Proposta criada depois de 18/01/2016?
IfMsgBox Yes
	EncCodeV := 1
IfMsgBox No
	EncCodeV := 2
If (Encargo>0)
{	
	;SetTimer ,TipoEncButtonNames,-10
	;MsgBox ,4131,OLH-V - Tipo Encargo, Tipo de Encargo?
	;IfMsgBox Yes
		VarEncargo := 1
	;IfMsgBox No
	;	VarEncargo := 2
	If (VarEncargo = 1)
	{
		InputBox ,Enc,OLH-V - Encargo Único,Qual o valor de Encargo`na ser inserido em todos os itens?,,250,145
		If ErrorLevel
			Gosub Sair
		StringReplace ,Enc,Enc,`,,`.
	}
	EncCode := 46387011
	
}
If (CIF>0)
{
	SetTimer ,ChangeButtonNames,-10
	MsgBox ,4131,OLH-V - Destino UF,Qual o Estado (UF) da venda da Quote?
	IfMsgBox Yes
	{
		InputBox ,Enc,OLH-V - EncargoCIF,Qual o valor de EncargoCIF`na ser inserido em todos os itens?,,250,145
		If ErrorLevel
			Gosub Sair
		StringReplace ,Enc,Enc,`,,`.
		If (EncCodeV = 1)
			EncCode := 54151436
		Else If (EncCodeV = 2)
			EncCode := 46384226
		Else
		{
			MsgBox ,4132,OLH-V - Interrupção,Programa interrompido.`n`n***Entre em contado com o Pedro.***
			GoSub Sair2
		}
	}
	IfMsgBox No
	{
		InputBox ,Enc,OLH-V - EncargoCIF,Qual o valor de EncargoCIF`na ser inserido em todos os itens?,,250,145
		If ErrorLevel
			Gosub Sair
		StringReplace ,Enc,Enc,`,,`.
		If (EncCodeV = 1)
			EncCode := 54151441
		Else If (EncCodeV = 2)
			EncCode := 46384158
		Else
		{
			MsgBox ,4132,OLH-V - Interrupção,Programa interrompido.`n`n***Entre em contado com o Pedro.***
			GoSub Sair2
		}
	}
	IfMsgBox Cancel
		Gosub Sair
}

Nome := "OLH-V - Ajustes"
CriarExcel(Nome,0)

SetTimer ,IntroButtonNames,-10
If Desconto = 1
	MsgBox ,4161,OLH-V - Instruções,Preencha a planinha criada conforme abaixo:`n`nColuna A: LLP`nColuna B: Desconto (Número)`nOBS.: Escrever "Pula", na coluna B, para caso deseje não mexer no item da linha correspondente.`n`nDepois de pronto clique em "Continuar".
Else
	MsgBox ,4161,OLH-V - Instruções,Preencha a planinha criada conforme abaixo:`n`nColuna A: LLP`nOBS.: Escrever "Pula", na coluna B, para caso deseje não mexer no item da linha correspondente.`n`nDepois de pronto clique em "Continuar".
IfMsgBox Cancel
	Gosub Sair
Sleep ,500
Info(1)
Suspend Off
Blockinput ,SendAndMouse
XL.Workbooks(Nome).Save
WinActivate ,ahk_group WinOracle
WinMaximize ,ahk_group WinOracle
WinGetPos ,,,OW,OH,ahk_group WinOracle
MouseMove ,(OW/2),(OH/2),0

Sleep ,250
If (UsarDP = 1)
	DefPric(1)
Sleep ,250
Blockinput ,SendAndMouse
i := 1
SetFormat ,Float,0.0
Linha := XL.WorksheetFunction.CountA(XL.Workbooks(Nome).Sheets(1).Range("A:A"))
Conta := Linha
;While (XL.Workbooks("Planilha OLH - Ajustes.xls").Sheets(1).Range("A"i).Value <> "")
While (i <= Linha)
{	
	GuiControl ,Text,Contador,%i%/%Linha%
	WinActivate ,ahk_group WinOracle
	SetFormat ,Float,0.5
	ValorLLP := XL.Workbooks(Nome).Sheets(1).Range("A"i).Value
	ValorDesc := XL.Workbooks(Nome).Sheets(1).Range("B"i).Value
	;ValorEnc := XL.Workbooks(Nome).Sheets(1).Range("C"i).Value
	SetFormat ,Float,0.0
	If (ValorDesc <> "Pula")
	{
		ActionsV()
		Adj := 0
		While Adj<=CIF+Encargo+Desconto
		{
			MouseClick ,left,195,490
			Sleep ,1000
			MouseClick ,left,195,490
			Sleep ,1500
			If Adj = 0
				SendInput ,46384254
			Else If Adj = 1
				If Desconto = 1
				{
					If (EncCodeV = 1)
						SendInput ,54151466
						;SendInput ,49491030
					Else If (EncCodeV = 2)
						SendInput ,46382521
					Else
					{
						MsgBox ,4132,OLH-V - Interrupção,Programa interrompido.`n`n***Entre em contado com o Pedro.***
						GoSub Sair2
					}
				}
				Else
				{
					SendInput ,%EncCode%
					Adj := 2
				}
			Else 
				SendInput ,%EncCode%
			Sleep ,500
			Send ,{Tab}
			Sleep ,2000
			SetFormat ,Float,0.5
			If Adj=0
				SendInput ,%ValorLLP%
			Else If (Desconto=1 and Adj=1)
				SendInput ,%ValorDesc%
			Else
				SendInput ,%Enc%
				;SendInput ,%ValorEnc%
			Sleep ,500
			Adj++
		}
		Send ,!a
		Esperar(1000,1500)
	}
	Proximo(i,Linha)
	i := i + 1
}
XL.DisplayAlerts:=False
XL.Workbooks(Nome).Close SaveChanges:=True
XL.DisplayAlerts:=True
If (UsarDP = 1)
	DefPric(0)

CalculaImposto:
Sleep ,1000 ;Questão para calcular o imposto logo após inserir os ajustes
If (BRTax = 1)
{
	Sleep ,1000
	Send ,+{PgUp}
	Esperar(500,1500)
	SetKeyDelay ,100
	Send ,{ALT}fs
	SetKeyDelay ,10
	Esperar(500,1500)
	Sleep ,1000
	Send ,!a
	Sleep ,1500
	SendInput ,cal
	Sleep ,1500
	Send ,{Enter}
	Esperar(500,2000)
	; Clipboard := ""
	; Sleep ,100
	; Send ,^c
	; Sleep ,200
	; If (Clipboard = "")
	; {
		; MsgBox ,4112,OLH - Erro,Houve um erro ao calcular o "BR_Tax".`nA Rotina está sendo interrompida.
		; ExitApp
	; }
	Send ,{Enter}
	Sleep ,1000
	SetKeyDelay ,100
	Send ,{ALT}fs
	Esperar(500,1500)
	Send ,+{PgUp}
	SetKeyDelay ,10
	Esperar(500,500)
	Gosub B3	
}
Sleep ,1000
SalvaESai(Nome)

;**********************************
;Botão 3 - Rotina Incluir Impostos
;**********************************
B3:
Linha:=0
If BRTax = 0
{	
	Gui Destroy
	
	MsgBox ,4131,OLH-V - Data, Proposta criada depois de 18/01/2016?
	IfMsgBox Yes
		EncCodeV := 1
	IfMsgBox No
		EncCodeV := 2
	
	Nome := "OLH-V - Inc Imposto"
	InputBox ,Linha,%Nome%,Repetir para quantos itens?,,250,130
	If ErrorLevel
		Gosub Sair
	Sleep ,500
	Info(1)
	Suspend Off
	Blockinput ,SendAndMouse
	XL.Workbooks(Nome).Save
	WinActivate ,ahk_group WinOracle
	WinMaximize ,ahk_group WinOracle
	WinGetPos ,,,OW,OH,ahk_group WinOracle
	MouseMove ,(OW/2),(OH/2),0
	Sleep ,1000
}
Else
	Linha := Conta
i := 1


While (i <= Linha)
{
	GuiControl ,Text,Contador,%i%/%Linha%
	WinActivate ,ahk_group WinOracle
	ActionsV()
	MouseClick ,left,195,490
	Sleep ,1500
	MouseClick ,left,195,490
	Sleep ,1500
	If (EncCodeV = 1)
		SendInput ,54151483
		;SendInput ,49491049
	Else If (EncCodeV = 2)
		SendInput ,4638248
	Else
	{
		MsgBox ,4132,OLH-V - Interrupção,Programa interrompido.`n`n***Entre em contado com o Pedro.***
		GoSub Sair2
	}
	Sleep ,1000
	Send ,{Tab}
	Sleep ,1000
	Send ,!a
	Esperar(1000,1500)
	Proximo(i,Linha)
	i := i + 1
}
SalvaESai(Nome)
	
;*********************************
;Botão 4 - Rotina Excluir Imposto
;*********************************
B4:
Gui Destroy
Nome := "OLH-V - Exc Imposto"
InputBox ,Linha,%Nome%,Repetir para quantos itens?,,250,130
If ErrorLevel
	Gosub Sair
Sleep ,500
Info(1)
Suspend Off
Blockinput ,SendAndMouse
WinActivate ,ahk_group WinOracle
WinMaximize ,ahk_group WinOracle
WinGetPos ,,,OW,OH,ahk_group WinOracle
MouseMove ,(OW/2),(OH/2),0
Sleep ,1000
i := 1
While i <= Linha
{
	GuiControl ,Text,Contador,%i%/%Linha%
	WinActivate ,ahk_group WinOracle
	ActionsV()
	SetKeyDelay ,100
	Send ,{ALT}vdl
	SetKeyDelay ,10
	Esperar(250,500)
	Send ,{TAB}
	Esperar(250,250)
	Send ,^c
	Sleep ,200
	If ((Clipboard = 46382484) or (Clipboard = 46382486) or (Clipboard = 54151483))
	{
		SetKeyDelay ,100
		Send ,{ALT}ed
		SetKeyDelay ,10
		Esperar(250,250)
		Send ,!a
	}
	Else
		Send ,!c
	Esperar(1000,1500)
	Proximo(i,Linha)
	i := i + 1
}
SalvaESai(Nome)

;******************************
;Botão 5 - Rotina Excluir Item
;******************************
B5:
Gui Destroy
Nome := "OLH-V - Exc Itens"
InputBox ,Linha,%Nome%,Repetir para quantos itens?,,250,130
If ErrorLevel
	Gosub Sair
Sleep ,500
Info(1)
Suspend Off
Blockinput ,SendAndMouse
WinActivate ,ahk_group WinOracle
WinMaximize ,ahk_group WinOracle
WinGetPos ,,,OW,OH,ahk_group WinOracle
MouseMove ,(OW/2),(OH/2),0
Sleep ,1000
i := 1
While i <= Linha
{
	GuiControl ,Text,Contador,%i%/%Linha%
	WinActivate ,ahk_group WinOracle
	SetKeyDelay ,100
	Send ,{ALT}ed
	SetKeyDelay ,10
	Sleep ,2500
	Send ,{ENTER}
	Esperar(1000,1500)
	i := i + 1
}
SalvaESai("Exclusão de Item")

;******************************
;Botão 6 - Rotina Resetar Item
;******************************
B6:
Gui Destroy
Nome := "OLH-V - Reset Item"
InputBox ,Linha,%Nome%,Repetir para quantos itens?,,250,130
If ErrorLevel
	Gosub Sair
Sleep ,500
Info(1)
Suspend Off
Blockinput ,SendAndMouse
WinActivate ,ahk_group WinOracle
WinMaximize ,ahk_group WinOracle
WinGetPos ,,,OW,OH,ahk_group WinOracle
MouseMove ,(OW/2),(OH/2),0
Sleep ,1000
i := 1
While i<=Linha
{
	GuiControl ,Text,Contador,%i%/%Linha%
	WinActivate ,ahk_group WinOracle
	ActionsV()
	Send ,!r
	Esperar(1000,1500)
	Send ,!a
	Esperar(1000,1500)
	Proximo(i,Linha)
	i := i + 1
}
SalvaESai("Reset Item")

;*****************************
;Botão 7 - Rotina Incluir TAG
;*****************************
B7:
Gui Destroy
Nome := "OLH-V - TAG"
CriarExcel(Nome,0)
SetTimer ,IntroButtonNames,-10
MsgBox ,4161,OLH-V - Instruções,Preencha a planinha criada conforme abaixo:`n`nColuna A: TAG.`n`nDepois de pronto clique em "Continuar".
IfMsgBox Cancel
	Gosub Sair
Sleep ,500
Info(1)
Suspend Off
Blockinput ,SendAndMouse
XL.Workbooks(Nome).Save
WinActivate ,ahk_group WinOracle
WinMaximize ,ahk_group WinOracle
WinGetPos ,,,OW,OH,ahk_group WinOracle
MouseMove ,(OW/2),(OH/2),0
Sleep ,1000
i:=1
SetFormat ,Float,0.0
Linha := XL.WorksheetFunction.CountA(XL.Workbooks(Nome).Sheets(1).Range("A:A"))
While (XL.Workbooks(Nome).Sheets(1).Range("A"i).Value <> "")
{
	GuiControl ,Text,Contador,%i%/%Linha%
	WinActivate ,ahk_group WinOracle
	ValorCellA:=XL.Workbooks(Nome).Sheets(1).Range("A"i).Value
	SetKeyDelay ,100
	Send ,{ALT}va
	SetKeyDelay ,10
	Esperar(1000,1500)
	Send ,{Tab}
	Sleep ,500
	SendInput ,JDI_TAG_INFO
	Sleep ,500
	Send ,{Tab}
	Sleep ,1000
	SendInput ,TAG
	Sleep ,500
	Send ,{TAB}{TAB}
	Sleep ,1000
	Send ,+{END}
	Sleep ,200
	SendMode Input
	SendRaw ,%ValorCellA%
	SendMode Event
	Sleep ,500
	SetKeyDelay ,100
	Send ,{ALT}fs
	Esperar(250,500)
	Send ,{ALT}fc
	SetKeyDelay ,10
	Esperar(1000,1500)
	Proximo(i,Linha)
	i := i + 1
}
XL.DisplayAlerts:=False
XL.Workbooks(Nome).Close SaveChanges:=True
XL.DisplayAlerts:=True
SalvaESai(Nome)

;*********************************
;Botão 8 - Rotina Alterar Desconto
;*********************************
B8:
Gui Destroy
Nome := "OLH-V - Alt Desconto"
CriarExcel(Nome,0)
SetTimer ,IntroButtonNames,-10
MsgBox ,4161,OLH-V - Instruções,Preencha a planinha criada conforme abaixo:`n`nColuna A: Novos descontos`nOBS.: Escrever "Pula" para caso deseje não mexer no item da linha correspondente.`n`nDepois de pronto clique em "Continuar".
IfMsgBox Cancel
	Gosub Sair
Sleep ,500
Info(1)
Suspend Off
Blockinput ,SendAndMouse
XL.Workbooks(Nome).Save
WinActivate ,ahk_group WinOracle
WinMaximize ,ahk_group WinOracle
WinGetPos ,,,OW,OH,ahk_group WinOracle
MouseMove ,(OW/2),(OH/2),0
Sleep ,250
i := 1
SetFormat ,Float,0.0
Linha := XL.WorksheetFunction.CountA(XL.Workbooks(Nome).Sheets(1).Range("A:A"))
Conta := Linha
While (XL.Workbooks(Nome).Sheets(1).Range("A"i).Value <> "")
{
	GuiControl ,Text,Contador,%i%/%Linha%
	WinActivate ,ahk_group WinOracle
	SetFormat ,Float,0.5
	ValorCellA := XL.Workbooks(Nome).Sheets(1).Range("A"i).Value
	SetFormat ,Float,0.0
	Sleep ,500
	ActionsV()
	SetKeyDelay ,100
	Send ,{ALT}vdl
	SetKeyDelay ,10
	Esperar(250,500)
	Send ,{Tab}
	Esperar(250,250)
	Encontrado := 0
	Loop ,7
	{
		Send ,^c
		Sleep ,200
		If ((Clipboard = 46382484) or (Clipboard = 46382486) or (Clipboard = 54151483))
		{
			SetKeyDelay ,100
			Send ,{ALT}ed
			SetKeyDelay ,10
			Esperar(250,250)
			Send ,{Tab}
			Esperar(250,250)
			Continue
		}
		If ((Clipboard = 46382521) or (Clipboard = 54151466) or (ValorCellA = "Pula"))
		{	
			Encontrado := 1
			Sleep ,500
			Break
		}
		Else
		{
			Send ,{Up}
			Esperar(250,250)
		}
	}
	If (Encontrado := 0)
	{
		SetTimer ,ContOUCanc,-10
		MsgBox ,4132,OLH-V - Aviso,Ajuste de Desconto não encontrado.`nDeseja continuar e ir para o próximo item ou cancelar a Rotina?
		IfMsgBox No
			Gosub Sair
		Sleep ,500
		Send ,!c
		Esperar(250,250)
		Proximo(i,Linha)
		i := i + 1
		Continue
	}
	Else
	{
		If (ValorCellA = "Pula")
		{
			Send ,!a
			Esperar(250,250)
		}
		Else
		{
			Send ,{Tab}
			Sleep ,500
			SendMode Input
			SendRaw ,%ValorCellA%
			SendMode Event
			Sleep ,500
			Send ,!a
			Esperar(250,250)
		}
		Proximo(i,Linha)
		i := i + 1
	}
}

XL.DisplayAlerts:=False
XL.Workbooks(Nome).Close SaveChanges:=True
XL.DisplayAlerts:=True

GoSub CalculaImposto

;******************************
;Botão A - Rotina Criar Item FCI
;******************************
BA:
Suspend Off
Gui Destroy
Sleep ,500

SetTimer ,TipoCadButtonNames,-20
MsgBox ,4,OLH-V - Cadastro FCI,Tipo de cadastro?
IfMsgBox ,Yes
{
	Suspend On
	Nome := ""
	Gui ,Name: New
	Gui ,Add,Text,w425 R2 cBlue Center,Preencher formulário abaixo:
	Gui ,Add,Text,w174 R1 Section,Descrição Curta em Inglês:
	Gui ,Add,Text,wp R1,Descrição Curta em Português:
	Gui ,Add,Text,wp R1,Product Numbers do Composto:
	Gui ,Add,Text,wp R1,EMR DIVISION ARG:
	Gui ,Add,Text,wp R1,
	Gui ,Add,Edit,w225 ys R1 -Wrap vDescrUSA,
	Gui ,Add,Edit,wp R1 -Wrap vDescrBR,
	Gui ,Add,Edit,wp R1 -Wrap vLDescr,Separados por "/"
	Gui ,Add,Edit,wp R1 -Wrap vDIVARG,Ex.: VLS.CAA
	Gui ,Add,Text,wp R1,
	Gui ,Add,Button,w70 Section gContFCI,&Continuar
	Gui ,Add,Button,w70 ys gSair,&Sair
	Gui ,-Sysmenu
	Gui ,Show,w431,OLH - Formulário de Cadastro
	Return
	
	ContFCI:
	Suspend Off
	SetFormat ,Float,0.0
	Linha := 1
	i := 1
	GuiControlGet ,DescrUSA
	GuiControlGet ,DescrBR
	GuiControlGet ,LDescr
	GuiControlGet ,DIVARG
	Gui Destroy
	Gosub ,ContinuacaoFCI
	Return
}
IfMsgBox ,No
{
	Nome := "OLH-V - Cadastrar Novo Item FCI"
	CriarExcel(Nome,1)
	Suspend On
	SetTimer ,IntroButtonNames, -20
	MsgBox ,65,OLH-V - Instruções,Preencha a planilha criada conforme abaixo:`n`n- Coluna A: Descrição Curta em Inglês do item.`n- Coluna B: Descrição Curta em Português do item.`n- Coluna C: Product Numbers do Composto (Separados por barra "/").`n- Coluna D: Código do EMR DIVISION ARG.`n`nDepois de preenchida a planilha clique em "Continuar".
	IfMsgBox Cancel
		GoSub Sair
	Suspend Off
	Linha := XL.WorksheetFunction.CountA(XL.Workbooks(Nome).ActiveSheet.Range("A:A"))
	i:=2
}
ContinuacaoFCI:
Info(0)
Sleep ,500

iini := i
	
wb := IEGet()
WinActivate ,ahk_class IEFrame
WinWaitActive ,ahk_class IEFrame
WinMaximize ,ahk_class IEFrame
WinGetPos ,,,OW,OH,ahk_class IEFrame
MouseMove ,(OW/2),(OH/2),0
wb.Navigate(PagOracle)
IELoad(wb)
Sleep ,1000
TrocarOrg("IMO")
Sleep ,1000
TrocarLingua("US")
Sleep ,1000
LinkOracle("Create Production Items (Item Catalog)",1)

CriaContador()
Lingua := "ENG"

If (Lingua = "ENG")
{
	While (i <= Linha)
	{	
		If (Nome <> "")
		{
			DescrUSA := XL.Workbooks(Nome).ActiveSheet.Range("A"i).Value
			LDescr := XL.Workbooks(Nome).ActiveSheet.Range("C"i).Value
			DIVARG := XL.Workbooks(Nome).ActiveSheet.Range("D"i).Value
		}
		
		StringReplace ,DIVARG,DIVARG,%A_Space%,,All
		StringUpper ,DIVARG,DIVARG
		
		INVCLASS := 8
		
		If (Nome <> "")
		{
			NAtual := i - 1
			NTotal := Linha - 1
			GuiControl ,Text,Contador,%Lingua%: %NAtual%/%NTotal%
		}
		Else
			GuiControl ,Text,Contador,%Lingua%: %i%/%Linha%
		WinActivate ,ahk_class IEFrame
		WinWaitActive ,ahk_class IEFrame
		
		wb.Navigate(PagIni)
		IELoad(wb)
		Sleep ,500
			
		wb.Document.All.EgoItemCatalogCategory.Focus()
		Sleep ,200
		SendInput ,EMR_VLVS
		Sleep ,200
		Send ,{TAB}
		WinWait ,Search and Select List of Values
		Sleep ,1000
		WinActivate ,Search and Select List of Values
		Loop ,9
		{
			Send ,{TAB}
			Sleep ,100
		}
		Send ,{Enter}
		While WinExist("Search and Select List of Values")
			Sleep ,100
		Loop
			Sleep ,100
		Until (!wb.busy)
		Sleep ,2000
		Send ,!c
		IELoad(wb)
		Sleep ,1000
		wb.Document.All.EgoItemDescription_UE.Focus()
		Sleep ,300
		Clipboard := DescrUSA
		Sleep ,200
		Send ,^v
		Sleep ,300
		Send ,{TAB}
		Sleep ,300
		Clipboard = %DescrUSA% // %LDescr%
		Sleep ,200
		Send ,^v
		Sleep ,300
		Send ,{TAB}
		Sleep ,300
		Clipboard := "EMR`%BR" ;%
		Sleep ,200
		Send ,^v
		; Finished Good BR
		Sleep ,200
		Send ,{TAB}
		WinWait ,Search and Select List of Values
		Sleep ,1000
		WinActivate ,Search and Select List of Values
		Loop ,8
		{
			Send ,{TAB}
			Sleep ,100
		}
		Send ,{Enter}
		Sleep ,3000
		Esperar(500,2000)
		Send ,!c
		IELoad(wb)
		Sleep ,1500
		Send ,!i
		IELoad(wb)
		Sleep ,500
		
		Links := wb.Document.Links
		sleep 500
		p := 0
		Loop % Links.Length ;%
		{	
			NovoFCI := Links[p].InnerText
			StringLeft ,CodFCI,NovoFCI,6
			If (CodFCI = "FCI_01")
			{
				If (Nome <> "")
					XL.Workbooks(Nome).ActiveSheet.Range("E"i).Value := NovoFCI
				Links[p].Click()
				Break
			}
			p := p + 1
		}
		IELoad(wb)
		Sleep ,500
		;Configurando Classification and Categories
		LinkOracle("Classification and Categories",0)
		ClicaBotao("Update",1)
		wb.Document.All("N17:EgoCategoryAttr:2").Focus()
		Sleep ,200
		Send ,{HOME}{DEL}VLS
		Sleep ,300
		Send ,{TAB}
		Esperar(500,1000)
		ClicaBotao("Add",0)
		wb.Document.All("N17:EgoCatalogAttr:4").Focus()
		Sleep ,200
		SendInput ,EMR DIVISION ARG
		Sleep ,300
		Send ,{TAB}
		Esperar(500,1000)
		Send ,{TAB}
		Sleep ,200
		Clipboard := DIVARG
		Sleep ,200
		Send ,^v
		Sleep ,300
		Send ,{TAB}
		Esperar(500,1000)
		Send ,!p
		IELoad(wb)
		Sleep ,1500

		i := i + 1
	}
}
	
;Trocar linguagem para PT-BR
WinActivate ,ahk_class IEFrame
WinWaitActive ,ahk_class IEFrame
TrocarLingua("BR")
i := iini
LinkOracle("Item Simple Search (Item Catalog)",1)
Lingua := "PTBR"

While (i <= Linha)
{	
	If (Nome <> "")
	{
		NAtual := i - 1
		NTotal := Linha - 1
		GuiControl ,Text,Contador,%Lingua%: %NAtual%/%NTotal%
	}
	Else
		GuiControl ,Text,Contador,%Lingua%: %i%/%Linha%
	Sleep ,200
	WinActivate ,ahk_class IEFrame
	
	If (Nome <> "")
	{
		DescrBR := XL.Workbooks(Nome).ActiveSheet.Range("B"i).Value
		LDescr := XL.Workbooks(Nome).ActiveSheet.Range("C"i).Value
		NovoFCI := XL.Workbooks(Nome).ActiveSheet.Range("E"i).Value
	}
	
	wb.Navigate(PagIni)
	IELoad(wb)
	Sleep ,500
	PesquisarItem(NovoFCI,1)
		
	wb.Document.All.EgoItemActionPopList.Focus()
	Sleep ,200
	Send ,{TAB}
	Sleep ,200
	Send ,{ENTER}
	IELoad(wb)
	wb.Document.All.EgoInputItemDescAttr.Focus()
	Sleep ,200
	Send ,+{END}
	Sleep ,200
	Clipboard := DescrBR
	Sleep ,200
	Send ,^v
	Sleep ,300
	Send ,{TAB}
	Sleep ,200
	Send ,^+{HOME}
	Sleep ,200
	Clipboard = %DescrBR% // %LDescr%
	Sleep ,200
	Send ,^v
	Sleep ,300
	Send ,{TAB}
	Sleep ,200
	Send ,!p
	IELoad(wb)
	Sleep ,1000
	
	;Associando o item
	LinkOracle("Organizations",0)
	ClicaBotao("AssToOrg",1)
	CMVO := 0
	CS01 := 0
	CSOR := 0
	Loop ,15
	{
		j := 0
		Loop ,25
		{	
			If ((wb.Document.All("N11:EgoOrganizationCode:" . j).InnerText = "MVO") or (wb.Document.All("N12:EgoOrganizationCode:" . j).InnerText = "MVO"))
			{
				If ((wb.Document.All("N11:EgoAssigned:" . j).checked <> True) or (wb.Document.All("N12:EgoAssigned:" . j).checked <> True))
				{
					wb.Document.All("N11:EgoAssigned:" . j).checked := True
					wb.Document.All("N12:EgoAssigned:" . j).checked := True
					Sleep ,100
				}
				CMVO := 1
			}
			;If ((wb.Document.All("N11:EgoOrganizationCode:" . j).InnerText = "S06") or (wb.Document.All("N12:EgoOrganizationCode:" . j).InnerText = "S06"))
			If ((wb.Document.All("N11:EgoOrganizationCode:" . j).InnerText = "S01") or (wb.Document.All("N12:EgoOrganizationCode:" . j).InnerText = "S01"))
			{
				If ((wb.Document.All("N11:EgoAssigned:" . j).checked <> True) or (wb.Document.All("N12:EgoAssigned:" . j).checked <> True))
				{
					wb.Document.All("N11:EgoAssigned:" . j).checked := True
					wb.Document.All("N12:EgoAssigned:" . j).checked := True
					Sleep ,100
				}
				CS01 := 1
			}
			If ((wb.Document.All("N11:EgoOrganizationCode:" . j).InnerText = "SOR") or (wb.Document.All("N12:EgoOrganizationCode:" . j).InnerText = "SOR"))
			{
				If ((wb.Document.All("N11:EgoAssigned:" . j).checked <> True) or (wb.Document.All("N12:EgoAssigned:" . j).checked <> True))
				{
					wb.Document.All("N11:EgoAssigned:" . j).checked := True
					wb.Document.All("N12:EgoAssigned:" . j).checked := True
					Sleep ,100
				}
				CSOR := 1
			}
			j := j + 1
			If (j = 26)
				MsgBox ,4132,OLH-V - Erro,Erro detectado!`nEntrar em contato com Pedro Mendonça.
		}
		Sleep ,1000
		If ((CMVO = 1) and (CS01 = 1) and (CSOR = 1))
			Break
		LinkOracle("Próximo 25",2)
	}
	Sleep ,1000
	Send ,!p
	IELoad(wb)
	
	i := i + 1
}

;Trocar linguagem para ENG
TrocarLingua("US")

wb.Navigate(PagOracle)

If (Nome <> "")
{
	XL.Workbooks(Nome).Save
}
Else
{
	Suspend On
	Gui ,Name: New
	Gui ,Add,Edit,w92 R1 Center,%NovoFCI%
	Gui ,Add,Button,w92 gSair2,&Sair
	Gui ,-Sysmenu
	Gui ,Show,w112,Rotina Finalizada!
	Return
}

MsgBox ,4160,OLH-V - Cadastro,Rotina Finalizada!
ExitApp

;*******************************
;Botão B - Rotina Criar Item _BR
;*******************************
BB:
Suspend Off
Gui Destroy
Sleep ,500

SetTimer ,TipoCadButtonNames,-20
MsgBox ,4,OLH-V - Cadastro _BR,Tipo de cadastro?
IfMsgBox ,Yes
{
	Suspend On
	Nome := ""
	Gui ,Name: New
	Gui ,Add,Text,w300 R2 cBlue Center,Preencher formulário abaixo:
	Gui ,Add,Text,w124 R1 Section,Novo código Oracle:
	Gui ,Add,Text,wp R1,Descrição em Inglês:
	Gui ,Add,Text,wp R1,Descrição em Português:
	Gui ,Add,Text,wp R1,EMR DIVISION ARG:
	Gui ,Add,Text,wp R1,
	Gui ,Add,Edit,w150 ys R1 vNovoCod,
	Gui ,Add,Edit,wp R1 -Wrap vDescrUSA
	Gui ,Add,Edit,wp R1 -Wrap vDescrBR
	Gui ,Add,Edit,wp R1 -Wrap vDIVARG,Ex.: VLS.CAA
	Gui ,Add,Text,wp R1,
	Gui ,Add,Button,w70 Section gCont_BR,&Continuar
	Gui ,Add,Button,w70 ys gSair,&Sair
	Gui ,-Sysmenu
	Gui ,Show,w306,OLH - Formulário de Cadastro
	Return
	
	Cont_BR:
	Suspend Off
	Linha := 1
	i := 1
	GuiControlGet ,NovoCod
	GuiControlGet ,DescrUSA
	GuiControlGet ,DescrBR
	GuiControlGet ,DIVARG
	GuiControlGet ,DRICode
	GuiControlGet ,PesoCod
	Gui Destroy
	Gosub ,Continuacao_BR
	Return
}
IfMsgBox ,No
{
	Nome := "OLH-V - Cadastrar Novo Item _BR"
	CriarExcel(Nome,2)
	Suspend On
	SetTimer, IntroButtonNames, -20
	MsgBox ,65,OLH-V - Instruções,Preencha a planilha criada conforme abaixo:`n`n- Coluna A: Novo código Oracle.`n- Coluna B: Descrição em Inglês do item.`n- Coluna C: Descrição em Português do item.`n- Coluna D: Código do EMR DIVISION ARG.`n`nDepois de preenchida a planilha clique em "Continuar".
	IfMsgBox Cancel
		GoSub Sair
	Suspend Off
	SetFormat ,Float,0.0
	Linha := XL.WorksheetFunction.CountA(XL.Workbooks(Nome).ActiveSheet.Range("A:A"))
	i := 2
}
Continuacao_BR:
Info(0)
Sleep ,500

iini := i
	
wb := IEGet()
WinActivate ,ahk_class IEFrame
WinWaitActive ,ahk_class IEFrame
WinMaximize ,ahk_class IEFrame
WinGetPos ,,,OW,OH,ahk_class IEFrame
MouseMove ,(OW/2),(OH/2),0
wb.Navigate(PagOracle)
IELoad(wb)
Sleep ,1000
TrocarOrg("IMO")
Sleep ,1000
TrocarLingua("US")
Sleep ,1000
LinkOracle("Create Production Items (Item Catalog)",1)

CriaContador()
Lingua := "ENG"

If (Lingua = "ENG")
{
	While (i <= Linha)
	{	
		If (Nome <> "")
		{
			NovoCod := XL.Workbooks(Nome).ActiveSheet.Range("A"i).Value
			DescrUSA := XL.Workbooks(Nome).ActiveSheet.Range("B"i).Value
			DIVARG := XL.Workbooks(Nome).ActiveSheet.Range("D"i).Value
			DRICode := XL.Workbooks(Nome).ActiveSheet.Range("E"i).Value
			PesoCod := XL.Workbooks(Nome).ActiveSheet.Range("F"i).Value
			LDescrUSA := XL.Workbooks(Nome).ActiveSheet.Range("G"i).Value
		}
		
		StringReplace ,NovoCod,NovoCod,%A_Space%,,All
		StringUpper ,NovoCod,NovoCod
		StringReplace ,DIVARG,DIVARG,%A_Space%,,All
		StringUpper ,DIVARG,DIVARG
		StringReplace ,DRICode,DRICode,%A_Space%,,All
		StringUpper ,DRICode,DRICode
		StringReplace ,PesoCod,PesoCod,%A_Space%,,All
		
		UnderBR := 0
		
		StringRight ,UnderBR,NovoCod,3
		INVCLASS := 8
		If (UnderBR = "_BR")
			StringTrimRight ,NovoCod,NovoCod,3
			
		If (Nome <> "")
		{
			NAtual := i - 1
			NTotal := Linha - 1
			GuiControl ,Text,Contador,%Lingua%: %NAtual%/%NTotal%
		}
		Else
			GuiControl ,Text,Contador,%Lingua%: %i%/%Linha%
		WinActivate ,ahk_class IEFrame
		WinWaitActive ,ahk_class IEFrame
		
		wb.Navigate(PagIni)
		IELoad(wb)
		Sleep ,500
			
		wb.Document.All.EgoItemCatalogCategory.Focus()
		Sleep ,200
		SendInput ,EMR_VLVS
		Sleep ,200
		Send ,{TAB}
		WinWait ,Search and Select List of Values
		Sleep ,1000
		WinActivate ,Search and Select List of Values
		Loop ,11
		{
			Send ,{TAB}
			Sleep ,100
		}
		Send ,{Enter}
		While WinExist("Search and Select List of Values")
			Sleep ,100
		Loop
			Sleep ,100
		Until (!wb.busy)
		Sleep ,2000
		Send ,!c
		IELoad(wb)
		Sleep ,1000
		wb.Document.All.EgoItemNumber_UE.Focus()
		Sleep ,500
		Clipboard = %NovoCod%_BR
		Sleep ,200
		Send ,^v
		Sleep ,300
		Send ,{TAB}{TAB}
		Sleep ,300
		Clipboard := DescrUSA
		Sleep ,200
		Send ,^v
		Sleep ,300
		Send ,{TAB}
		Sleep ,300
		If (LDescrUSA = "")
		{
			Clipboard = %DescrUSA% // %NovoCod%
			Sleep ,200
			Send ,^v
		}
		Else
		{
			Clipboard = %LDescrUSA% // %NovoCod%
			Sleep ,200
			Send ,^v
		}
		Sleep ,300
		Send ,{TAB}
		Sleep ,300
		Clipboard := "EMR`%BR" ;%
		Sleep ,200
		Send ,^v
		;Finished Good BR
		Sleep ,200
		Send ,{TAB}
		WinWait ,Search and Select List of Values
		Sleep ,1000
		WinActivate ,Search and Select List of Values
		Loop ,10
		{
			Send ,{TAB}
			Sleep ,100
		}
		Send ,{Enter}
		Sleep ,3000
		Esperar(500,2000)
		Send ,!c
		IELoad(wb)
		Sleep ,1500
		Send ,!i
		IELoad(wb)
		Sleep ,500
		
		Links := wb.Document.Links
		sleep 500
		p := 0
		Loop % Links.Length ;%
		{	
			If (Links[p].InnerText = NovoCod . "_BR")
			{
				Links[p].Click()
				Break
			}
			p := p + 1
		}
		IELoad(wb)
		Sleep ,500
		;Configurando Classification and Categories
		LinkOracle("Classification and Categories",0)
		ClicaBotao("Update",1)
		wb.Document.All("N17:EgoCategoryAttr:2").Focus()
		Sleep ,200
		Send ,{HOME}{DEL}VLS
		Sleep ,300
		Send ,{TAB}
		Esperar(500,1000)
		ClicaBotao("Add",0)
		wb.Document.All("N17:EgoCatalogAttr:4").Focus()
		Sleep ,200
		SendInput ,EMR DIVISION ARG
		Sleep ,300
		Send ,{TAB}
		Esperar(500,1000)
		Send ,{TAB}
		Sleep ,200
		Clipboard := DIVARG
		Sleep ,200
		Send ,^v
		Sleep ,300
		Send ,{TAB}
		Esperar(500,1000)
		Send ,!p
		IELoad(wb)
		Sleep ,1500

		i := i + 1
	}
}
	
;Trocar linguagem para PT-BR
WinActivate ,ahk_class IEFrame
WinWaitActive ,ahk_class IEFrame
TrocarLingua("BR")
i := iini
LinkOracle("Item Simple Search (Item Catalog)",1)
Lingua := "PTBR"

While (i <= Linha)
{	
	If (Nome <> "")
	{
		NAtual := i - 1
		NTotal := Linha - 1
		GuiControl ,Text,Contador,%Lingua%: %NAtual%/%NTotal%
	}
	Else
		GuiControl ,Text,Contador,%Lingua%: %i%/%Linha%
	Sleep ,200
	WinActivate ,ahk_class IEFrame
	
	If (Nome <> "")
	{
		NovoCod := XL.Workbooks(Nome).ActiveSheet.Range("A"i).Value
		DescrBR := XL.Workbooks(Nome).ActiveSheet.Range("C"i).Value
		LDescrBR := XL.Workbooks(Nome).ActiveSheet.Range("H"i).Value

		StringReplace ,NovoCod,NovoCod,%A_Space%,,All
		StringUpper ,NovoCod,NovoCod
		StringReplace ,INVCLASS,INVCLASS,%A_Space%,,All
		
		UnderBR := 0
		
		StringRight ,UnderBR,NovoCod,3
		INVCLASS := 8
		If (UnderBR = "_BR")
			StringTrimRight ,NovoCod,NovoCod,3
	}
	
	wb.Navigate(PagIni)
	IELoad(wb)
	Sleep ,500
	PesquisarItem(NovoCod . "_BR",1)
		
	wb.Document.All.EgoItemActionPopList.Focus()
	Sleep ,200
	Send ,{TAB}
	Sleep ,200
	Send ,{ENTER}
	IELoad(wb)
	wb.Document.All.EgoInputItemDescAttr.Focus()
	Sleep ,200
	Send ,+{END}
	Sleep ,200
	Clipboard := DescrBR
	Sleep ,200
	Send ,^v
	Sleep ,300
	Send ,{TAB}
	Sleep ,200
	Send ,^+{HOME}
	Sleep ,200
	If (LDescrBR = "")
	{
		Clipboard = %DescrBR% // %NovoCod%
		Sleep ,200
		Send ,^v
	}
	Else
	{
		Clipboard = %LDescrBR% // %NovoCod%
		Sleep ,200
		Send ,^v
	}
	Sleep ,300
	Send ,{TAB}
	Sleep ,200
	Send ,!p
	IELoad(wb)
	Sleep ,1000
	
	;Associando o item
	LinkOracle("Organizations",0)
	ClicaBotao("AssToOrg",1)
	CMVO := 0
	CS01 := 0
	CSOR := 0
	Loop ,15
	{
		j := 0
		Loop ,25
		{	
			If ((wb.Document.All("N11:EgoOrganizationCode:" . j).InnerText = "MVO") or (wb.Document.All("N12:EgoOrganizationCode:" . j).InnerText = "MVO"))
			{
				If ((wb.Document.All("N11:EgoAssigned:" . j).checked <> True) or (wb.Document.All("N12:EgoAssigned:" . j).checked <> True))
				{
					wb.Document.All("N11:EgoAssigned:" . j).checked := True
					wb.Document.All("N12:EgoAssigned:" . j).checked := True
					Sleep ,100
				}
				CMVO := 1
			}
			If ((wb.Document.All("N11:EgoOrganizationCode:" . j).InnerText = "S01") or (wb.Document.All("N12:EgoOrganizationCode:" . j).InnerText = "S01"))
			{
				If ((wb.Document.All("N11:EgoAssigned:" . j).checked <> True) or (wb.Document.All("N12:EgoAssigned:" . j).checked <> True))
				{
					wb.Document.All("N11:EgoAssigned:" . j).checked := True
					wb.Document.All("N12:EgoAssigned:" . j).checked := True
					Sleep ,100
				}
				CS01 := 1
			}
			If ((wb.Document.All("N11:EgoOrganizationCode:" . j).InnerText = "SOR") or (wb.Document.All("N12:EgoOrganizationCode:" . j).InnerText = "SOR"))
			{
				If ((wb.Document.All("N11:EgoAssigned:" . j).checked <> True) or (wb.Document.All("N12:EgoAssigned:" . j).checked <> True))
				{
					wb.Document.All("N11:EgoAssigned:" . j).checked := True
					wb.Document.All("N12:EgoAssigned:" . j).checked := True
					Sleep ,100
				}
				CSOR := 1
			}
			j := j + 1
		}
		Sleep ,1000
		If ((CMVO = 1) and (CS01 = 1) and (CSOR = 1))
			Break
		LinkOracle("Próximo 25",2)
	}
	Sleep ,1000
	Send ,!p
	IELoad(wb)
	
	i := i + 1
}

;Trocar linguagem para ENG
TrocarLingua("US")

wb.Navigate(PagOracle)

If (Nome <> "")
{
	XL.Workbooks(Nome).Save
	XL.Quit()
}

MsgBox ,4160,OLH-V - Cadastro,Rotina Finalizada!
ExitApp

;****************************************
;Botão C - Rotina Cadastrar Novo Workflow
;****************************************
BC:
Suspend Off
Gui Destroy
Sleep ,250

SetTimer ,TipoCadButtonNames,-20
MsgBox ,4,OLH-V - Cadastro Workflow,Tipo de cadastro?
IfMsgBox ,Yes
{
	Suspend On
	Nome := ""
	Gui ,Name: New
	Gui ,Add,Text,w350 R2 cBlue Center,Preencher formulário abaixo:
	Gui ,Add,Text,w124 R1 section,Código Oracle:
	Gui ,Add,Text,wp R1,Número do Worflow:
	Gui ,Add,Text,wp R1,Nome do Workflow:
	Gui ,Add,Text,wp R1,Descrição Genérica:
	Gui ,Add,Text,wp R1,Origem:
	Gui ,Add,Text,wp R1,Uso Militar ou Nuclear?:
	Gui ,Add,Text,wp R1,
	Gui ,Add,Edit,w210 R1 ys section -Wrap vCodOracle
	Gui ,Add,Edit,wp R1 -Wrap vNumWF
	Gui ,Add,Edit,wp R1 -Wrap vNomeWF
	Gui ,Add,DropDownList,wp vDescWF,Válvula Globo|Válvula Esfera|Válvula Borboleta|Atuador Pneumático|Posicionador|Filtro Regulador/Booster/Lock-up/Trip|Solenóide|Monitor TopWorx|Controlador Pneumático C1|Transdutor|Displacer|Dessuperaquecedor|Partes/Peças - Dessuper|Gateway/THUM 775|Hart Modem
	;Gui ,Add,DropDownList,wp vDescWF,Válvula Globo|Válvula Esfera|Válvula Borboleta|Partes/Peças - Válvula|Atuador Pneumático|Partes/Peças - Atuador Pneumático|Posicionador|Partes/Peças - Posicionador|Filtro Regulador/Booster/Lock-up/Trip|Solenóide|Monitor TopWorx|Controlador Pneumático C1|Transdutor|Displacer|Dessuperaquecedor|Partes/Peças - Dessuper|Gateway/THUM 775|Hart Modem
	Gui ,Add,DropDownList,wp AltSubmit vOriWF,0 - Nacional|1 - 100`% Importado|3 - >40`% Importado|5 - <40`% Importado ;%
	Gui ,Add,DropDownList,wp AltSubmit vMiNu,Não|Militar|Nuclear
	Gui ,Add,Text,wp R1,
	Gui ,Add,Button,w70 section gContCadWF,&Continuar
	Gui ,Add,Button,w70 ys gSair2,&Sair
	Gui ,-Sysmenu
	Gui ,Show,w366,OLH-V - Formulário de Workflow
	Return
	
	ContCadWF:
	Suspend Off
	Linha := 1
	i := 1
	GuiControlGet ,CodOracle
	GuiControlGet ,TipoWF
	GuiControlGet ,NumWF
	GuiControlGet ,NomeWF
	GuiControlGet ,DescWF
	GuiControlGet ,OriWF
	GuiControlGet ,MiNu
	
	Gui Destroy
	Gosub ,ContinuacaoWF
	Return
}
IfMsgBox ,No
{
	Nome := "OLH-V - Criar Workflow"
	CriarExcel(Nome,3)
	Suspend On
	SetTimer, IntroButtonNames, -20
	MsgBox ,65,OLH-V - Instruções,Preencha a planilha criada conforme abaixo:`n`nColuna A: Código Oracle`nColuna B: "Associação" / "Purchase/BuyOut" / "Finished Good"`nColuna C: Número do Workflow`nColuna D: Nome do Workflow`nColuna E: Descrição Genérica do Workflow`nColuna F: Origem do Item`nColuna G: Uso Militar ou Nuclear?`n`nDepois de pronto clique em "Continuar".
	IfMsgBox Cancel
		GoSub Sair
	Suspend Off
	Linha := XL.WorksheetFunction.CountA(XL.Workbooks(Nome).ActiveSheet.Range("A:A"))
	i := 2
}

ContinuacaoWF:
Info(0)
Sleep ,500

wb := IEGet()
WinActivate ,ahk_class IEFrame
WinWaitActive ,ahk_class IEFrame
WinGetPos ,,,OW,OH,ahk_class IEFrame
MouseMove ,(OW/2),(OH/2),0

wb.Navigate(PagOracle)
IELoad(wb)
Sleep ,1000
TrocarOrg("S01")
Sleep ,1000
iini := i
LinkOracle("Create Request for Specifying New Item Attributes (Change Management)",1)
	
Erro := 0
CriaContador()

While (i <= Linha)
{	
	If (Nome <> "")
	{
		NAtual := i - 1
		NTotal := Linha - 1
		GuiControl ,Text,Contador,%NAtual%/%NTotal%
	}
	Else
		GuiControl ,Text,Contador,%i%/%Linha%
	WinActivate ,ahk_class IEFrame
	WinWaitActive ,ahk_class IEFrame
	
	wb.Navigate(PagIni)
	IELoad(wb)
	Sleep ,500
	
	If (Nome <> "")
	{
		CodOracle := XL.Workbooks(Nome).ActiveSheet.Range("A"i).Value ;Código para que está sendo gerado o Workflow
		TipoWF := XL.Workbooks(Nome).ActiveSheet.Range("B"i).Value ;Tipo de Workflow: 1 - Associação / 2 - BuyOut / 3 - Finished Good
		NumWF := XL.Workbooks(Nome).ActiveSheet.Range("C"i).Value ;Número do Workflow
		NomeWF := XL.Workbooks(Nome).ActiveSheet.Range("D"i).Value ;Nome do Workflow
		DescWF := XL.Workbooks(Nome).ActiveSheet.Range("E"i).Value ;Descrição do Workflow
		OriWF := XL.Workbooks(Nome).ActiveSheet.Range("F"i).Value ;Origem do Item
		Minu := XL.Workbooks(Nome).ActiveSheet.Range("G"i).Value ;Militar ou Nuclear?
	}
	
	;StringReplace ,TipoWF,TipoWF,%A_Space%,,All
	StringReplace ,NumWF,NumWF,%A_Space%,,All
	StringReplace ,CodOracle,CodOracle,%A_Space%,,All
	StringUpper ,CodOracle,CodOracle
	
	If (Nome = "")
	{
		If (OriWF = 1)
			OriWF := "Produto 100`% Nacional" ;%
		If (OriWF = 2)
			OriWF := "Produto 100`% Importado" ;%
		If (OriWF = 3)
			OriWF := "Produto nacional com conteúdo importado superior a 40`% (Origem 3)" ;% 
		If (OriWF = 4)
			OriWF := "Produto nacional com conteúdo importado inferior a 40`% (Origem 5)" ;%
			
		If (MiNu = "1")
			MiNu := "Uso Militar ou Nuclear?: Não"
		If (MiNu = "2")
			MiNu := "Uso Militar ou Nuclear?: Militar"
		If (MiNu = "3")
			MiNu := "Uso Militar ou Nuclear?: Nuclear"
	}
	Else
	{
		If (OriWF = "0 - 100`% Nacional") ;%
			OriWF := "Produto 100`% Nacional" ;%
		If (OriWF = "1 - 100`% Importado") ;%
			OriWF := "Produto 100`% Importado" ;%
		If (OriWF = "3 - >40`% Importado") ;%
			OriWF := "Produto nacional com conteúdo importado superior a 40`% (Origem 3)" ;% 
		If (OriWF = "5 - <40`% Importado") ;%
			OriWF := "Produto nacional com conteúdo importado inferior a 40`% (Origem 5)" ;%
			
		If (TipoWF = "Associação")
			TipoWF := 1
		If (TipoWF = "Purchase/BuyOut")
			TipoWF := 2
		If (TipoWF = "Finished Good")
			TipoWF := 3
			
		If (MiNu = "Não")
			MiNu := "Uso Militar ou Nuclear?: Não"
		If (MiNu = "Militar")
			MiNu := "Uso Militar ou Nuclear?: Militar"
		If (MiNu = "Nuclear")
			MiNu := "Uso Militar ou Nuclear?: Nuclear"
	}
		
	If (DescWF = "Válvula Globo")
	{
		DescWF := "Válvula de Controle/ON-OFF Tipo Globo"
		NCMWF := "8481.80.94"
	}
	If (DescWF = "Válvula Esfera")
	{
		DescWF := "Válvula de Controle/ON-OFF Tipo Esfera"
		NCMWF := "8481.80.95"
	}
	If (DescWF = "Válvula Borboleta")
	{
		DescWF := "Válvula de Controle/ON-OFF Tipo Borboleta"
		NCMWF := "8481.80.97"
	}
	If (DescWF = "Partes/Peças - Válvula")
	{
		DescWF := "Partes/Peças para Válvula de Controle/ON-OFF"
		NCMWF := "8481.90.90"
	}
	If (DescWF = "Atuador Pneumático")
		NCMWF := "8412.39.00"
	If (DescWF = "Partes/Peças - Atuador Pneumático")
	{
		DescWF := "Partes/Peças para Atuador Pneumático"
		NCMWF := "8412.90.90"
	}
	If (DescWF = "Posicionador")
	{
		DescWF := "Posicionador para Válvula de Controle"
		NCMWF := "9032.81.00"
	}
	If (DescWF = "Partes/Peças - Posicionador")
	{
		DescWF := "Partes/Peças para Posicionador"
		NCMWF := "8412.90.90"
	}
	If (DescWF = "Filtro Regulador/Booster/Lock-up/Trip")
		NCMWF := "8481.10.00"
	If (DescWF = "Solenóide")
		NCMWF := "8481.80.92"
	If (DescWF = "Monitor TopWorx")
	{	
		DescWF := "Monitor de Posição TopWorx"
		NCMWF := "8536.50.90"
	}
	If (DescWF = "Controlador Pneumático C1")
		NCMWF := "9026.20.90"
	If (DescWF = "Transdutor")
		NCMWF := "8454.10.00"
	If (DescWF = "Displacer")
		NCMWF := "9026.10.29"
	If (DescWF = "Dessuperaquecedor")
		NCMWF := "8404.10.10"
	If (DescWF = "Partes/Peças - Dessuper")
	{
		DescWF := "Partes/Peças para Dessuperaquecedor"
		NCMWF := "8404.90.90"
	}
	If (DescWF = "Gateway/THUM 775")
	{
		DescWF := "Equipamento para Wireless HART (Gateway ou THUM 775)"
		NCMWF := "9026.90.90"
	}
	If (DescWF = "Hart Modem")
	{
		NCMWF := "9026.90.90"
	}
		
	wb.Document.All.changeType_poplist.value := 4880
			
	Sleep ,250
	Send ,!c
	IELoad(wb)
	Sleep ,1000
	wb.Document.All.ChangeNumber.Focus()
	Sleep ,200
	Clipboard := NumWF
	Sleep ,200
	Send ,^v
	Sleep ,200
	Send ,{TAB}
	wb.Document.All.ChangeName.Focus()
	Sleep ,200
	Clipboard := NomeWF
	Sleep ,200
	Send ,^v
	Sleep ,200
	Send ,{TAB}
	wb.Document.All.Description.Focus()
	Sleep ,200
	If (TipoWF = 1)
		Clipboard = %DescWF%`n%OriWF%`n%Minu%`nSugestão NCM: %NCMWF%`nItem associado em %MyBU% e SOR
	Else
		Clipboard = %DescWF%`n%OriWF%`n%Minu%`nSugestão NCM: %NCMWF%
	Sleep ,250
	Send ,^v
	Sleep ,250
	Send ,{TAB}
	wb.Document.All.SubjectName1.Focus()
	Sleep ,200
	Clipboard := CodOracle
	Sleep ,200
	Send ,^v
	Sleep ,200
	Send ,{TAB}
	Sleep ,1000
	WinWait ,Search and Select List of Values,,3
	If (ErrorLevel = 0)
	{
		WinActivate ,Search and Select List of Values
		Loop ,8
		{
			Send ,{TAB}
			Sleep ,100
		}
		Sleep ,250
		Send ,{Enter}
		While WinExist("Search and Select List of Values")
			Sleep ,100
		Sleep ,3000
	}
	Send ,!m
	IELoad(wb)
	Sleep ,2000
	i := i + 1
}
wb.Navigate(PagOracle)

XL.Workbooks(Nome).Save

If (Nome <> "")
{
	XL.Workbooks(Nome).Save
	XL.Quit()
}

MsgBox ,4160,Workflow Oracle,Rotina Finalizada!
ExitApp

;***********************************************
;Botão D - Rotina Verificador de Cadastro Oracle
;***********************************************
BD:
Suspend Off
Gui Destroy
Sleep ,500

Nome := "OLH-E - Verificar Cadastro"
CriarExcel(Nome,0)

Suspend On
SetTimer ,IntroButtonNames,-20
MsgBox ,65,OLH-V - Instruções,Preencha a planilha criada conforme abaixo:`n`nColuna A: Código Oracle a ser verificado.`n`nDepois de pronto clique em "Continuar".
IfMsgBox Cancel
	GoSub Sair
Suspend Off
Info(0)
Sleep ,500

i := 1
iini := i
Erro := 0	
Linha := XL.WorksheetFunction.CountA(XL.Workbooks(Nome).ActiveSheet.Range("A:A"))
wb := IEGet()
WinActivate ,ahk_class IEFrame
WinWaitActive ,ahk_class IEFrame
WinMaximize ,ahk_class IEFrame
WinGetPos ,,,OW,OH,ahk_class IEFrame
MouseMove ,(OW/2),(OH/2),0
wb.Navigate(PagOracle)
IELoad(wb)
Sleep ,1000
TrocarOrg("S01")
LinkOracle("Item Simple Search (Item Catalog)",1)

CriaContador()

While (XL.Workbooks(Nome).ActiveSheet.Range("A"i).Value <> "")
{	
	GuiControl ,Text,Contador,%i%/%Linha%
	WinActivate ,ahk_class IEFrame
	CodOracle := XL.Workbooks(Nome).ActiveSheet.Range("A"i).Value
	StringReplace ,CodOracle,CodOracle,%A_Space%,,All
	StringUpper ,CodOracle,CodOracle
	wb.Navigate(PagIni)
	IELoad(wb)
	Sleep ,500
	PesquisarItem(CodOracle,0)
	i := i + 1
}

wb.Navigate(PagOracle)

If (Erro = 0)
{
	XL.Workbooks(Nome).Save
	XL.Quit()

	MsgBox ,4160,Verificar Cadastro,Rotina Finalizada!`nTodos os códigos estavam cadastrados.
	ExitApp
}
Else
{
	XL.Workbooks(Nome).Save

	MsgBox ,4144,Verificar Cadastro,Rotina Finalizada!`nQuantidade de item(s) não encontrado(s): %Erro%`nVerifique na planilha os item(s) pintado(s) em vermelho.
	ExitApp
}

;*******Funções*******
ContOUCanc:
IfWinNotExist ,OLH-V - Aviso
    Return
WinActivate ,OLH-V - Aviso
ControlSetText, Button1, &Continuar
ControlSetText, Button2, &Cancelar
Return

TimerButtonName:
IfWinNotExist ,OLH-V - Informações
	Return
SetTimer ,TimerButtonName,950
ControlSetText, Button1, &OK (%TicTac%)
TicTac := TicTac - 1
Return
	
ChangeButtonNames: 
IfWinNotExist ,OLH-V - Destino UF
    Return
WinActivate ,OLH-V - Destino UF
ControlSetText, Button1, &São Paulo
ControlSetText, Button2, &Outros
ControlSetText, Button3, &Sair
Return

IntroButtonNames: 
IfWinNotExist ,OLH-V - Instruções
    Return
WinActivate ,OLH-V - Instruções
ControlSetText, Button1, &Continuar
ControlSetText, Button2, C&ancelar
Return

TipoCadButtonNames: 
IfWinNotExist ,OLH-V - Cadastro
    Return
WinActivate ,OLH-V - Cadastro
ControlSetText, Button1, Ú&nico
ControlSetText, Button2, &Múltiplo
Return

TipoEncButtonNames: 
IfWinNotExist ,OLH-V - Tipo Encargo
    Return
WinActivate ,OLH-V - Tipo Encargo
ControlSetText, Button1, Ú&nico
ControlSetText, Button2, &Variado
Return

ImpostoButtonName: 
IfWinNotExist ,OLH-V - Opção Imposto
    Return
WinActivate ,OLH-V - Opção Imposto
SetTimer ,ImpostoButtonName,950
ControlSetText, Button2, &No (%TicTac%)
TicTac := TicTac - 1
Return

CriarExcel(NomeF,Num)
{
	XL:=ComObjCreate("Excel.Application")
	XL.Workbooks.Add
	XL.DisplayAlerts:=False
	XL.ActiveWorkbook.SaveAs("C:\Temp\" . NomeF . ".xlsx")
	XL.DisplayAlerts:=True
	If (Num = 1)
	{
		XL.Workbooks(NomeF).ActiveSheet.Range("A1").Value := "Descrição (US)"
		XL.Workbooks(NomeF).ActiveSheet.Range("B1").Value := "Descrição (BR)"
		XL.Workbooks(NomeF).ActiveSheet.Range("C1").Value := "Product Numbers"
		XL.Workbooks(NomeF).ActiveSheet.Range("D1").Value := "EMR DIVISION ARG"
		XL.Workbooks(NomeF).ActiveSheet.Range("A1:D1").Interior.Colorindex := 6
		XL.Workbooks(NomeF).ActiveSheet.Range("A1:D1").Font.Bold := True
		XL.Workbooks(NomeF).ActiveSheet.Range("C:C").Select
		XL.Workbooks(NomeF).ActiveSheet.Range("C:C").FormatConditions.AddUniqueValues
		XL.Workbooks(NomeF).ActiveSheet.Range("C:C").FormatConditions(XL.Workbooks(NomeF).ActiveSheet.Range("C:C").FormatConditions.Count).SetFirstPriority
		XL.Workbooks(NomeF).ActiveSheet.Range("C:C").FormatConditions(1).DupeUnique := 1
		XL.Workbooks(NomeF).ActiveSheet.Range("C:C").FormatConditions(1).Font.Bold := True
		XL.Workbooks(NomeF).ActiveSheet.Range("C:C").FormatConditions(1).Font.Italic := False
		XL.Workbooks(NomeF).ActiveSheet.Range("C:C").FormatConditions(1).Font.TintAndShade := 0
		XL.Workbooks(NomeF).ActiveSheet.Range("C:C").FormatConditions(1).Borders(-4131).LineStyle := -4142
		XL.Workbooks(NomeF).ActiveSheet.Range("C:C").FormatConditions(1).Borders(-4152).LineStyle := -4142
		XL.Workbooks(NomeF).ActiveSheet.Range("C:C").FormatConditions(1).Borders(-4160).LineStyle := -4142
		XL.Workbooks(NomeF).ActiveSheet.Range("C:C").FormatConditions(1).Borders(-4107).LineStyle := -4142
		XL.Workbooks(NomeF).ActiveSheet.Range("C:C").FormatConditions(1).Interior.PatternColorIndex := -4105
		XL.Workbooks(NomeF).ActiveSheet.Range("C:C").FormatConditions(1).Interior.Color := 255
		XL.Workbooks(NomeF).ActiveSheet.Range("C:C").FormatConditions(1).Interior.TintAndShade := 0
		XL.Workbooks(NomeF).ActiveSheet.Range("C:C").FormatConditions(1).StopIfTrue := False
		XL.Workbooks(NomeF).ActiveSheet.Cells.Select
		XL.Workbooks(NomeF).ActiveSheet.Cells.EntireColumn.AutoFit
		XL.Workbooks(NomeF).ActiveSheet.Range("A2").Select
	}
	If (Num = 2)
	{
		XL.Workbooks(NomeF).ActiveSheet.Range("A1").Value := "Código Oracle"
		XL.Workbooks(NomeF).ActiveSheet.Range("B1").Value := "Descrição (US)"
		XL.Workbooks(NomeF).ActiveSheet.Range("C1").Value := "Descrição (BR)"
		XL.Workbooks(NomeF).ActiveSheet.Range("D1").Value := "EMR DIVISION ARG"
		XL.Workbooks(NomeF).ActiveSheet.Range("A1:D1").Interior.Colorindex := 6
		XL.Workbooks(NomeF).ActiveSheet.Range("A1:D1").Font.Bold := True
		XL.Workbooks(NomeF).ActiveSheet.Cells.Select
		XL.Workbooks(NomeF).ActiveSheet.Cells.EntireColumn.AutoFit
		XL.Workbooks(NomeF).ActiveSheet.Range("A2").Select
	}
	If (Num = 3)
	{
		XL.Workbooks(NomeF).ActiveSheet.Range("A1").Value := "Código Oracle"
		XL.Workbooks(NomeF).ActiveSheet.Range("B1").Value := "Tipo"
		XL.Workbooks(NomeF).ActiveSheet.Range("B2").Select
		Xl.Selection.Validation.Delete
        Xl.Selection.Validation.Add(3,1,1,"Associação`;Purchase/BuyOut`;Finished Good")
        Xl.Selection.Validation.IgnoreBlank := True
        Xl.Selection.Validation.InCellDropdown := True
        Xl.Selection.Validation.InputTitle := ""
        Xl.Selection.Validation.ErrorTitle := ""
        Xl.Selection.Validation.InputMessage := ""
        Xl.Selection.Validation.ErrorMessage := ""
        Xl.Selection.Validation.ShowInput := True
        Xl.Selection.Validation.ShowError := True
		XL.Workbooks(NomeF).ActiveSheet.Range("C1").Value := "Número"
		XL.Workbooks(NomeF).ActiveSheet.Range("D1").Value := "Nome"
		XL.Workbooks(NomeF).ActiveSheet.Range("E1").Value := "Descrição Genérica"
		XL.Workbooks(NomeF).ActiveSheet.Range("E2").Select
		Xl.Selection.Validation.Delete
        Xl.Selection.Validation.Add(3,1,1,"Válvula Globo`;Válvula Esfera`;Válvula Borboleta`;Atuador Pneumático`;Posicionador`;Filtro Regulador/Booster/Lock-up/Trip`;Solenóide`;Monitor TopWorx`;Controlador Pneumático C1`;Transdutor`;Displacer`;Dessuperaquecedor`;Partes/Peças - Dessuper`;Gateway/THUM 775`;Hart Modem")
		;Xl.Selection.Validation.Add(3,1,1,"Válvula Globo`;Válvula Esfera`;Válvula Borboleta`;Partes/Peças - Válvula`;Atuador Pneumático`;Partes/Peças - Atuador Pneumático`;Posicionador`;Partes/Peças - Posicionador`;Filtro Regulador/Booster/Lock-up/Trip`;Solenóide`;Monitor TopWorx`;Controlador Pneumático C1`;Transdutor`;Displacer`;Dessuperaquecedor`;Partes/Peças - Dessuper`;Gateway/THUM 775`;Hart Modem")
        Xl.Selection.Validation.IgnoreBlank := True
        Xl.Selection.Validation.InCellDropdown := True
        Xl.Selection.Validation.InputTitle := ""
        Xl.Selection.Validation.ErrorTitle := ""
        Xl.Selection.Validation.InputMessage := ""
        Xl.Selection.Validation.ErrorMessage := ""
        Xl.Selection.Validation.ShowInput := True
        Xl.Selection.Validation.ShowError := True
		XL.Workbooks(NomeF).ActiveSheet.Range("F1").Value := "Origem"
		XL.Workbooks(NomeF).ActiveSheet.Range("F2").Select
		Xl.Selection.Validation.Delete
        Xl.Selection.Validation.Add(3,1,1,"0 - Nacional`;1 - 100`% Importado`;3 - >40`% Importado`;5 - <40`% Importado")
        Xl.Selection.Validation.IgnoreBlank := True
        Xl.Selection.Validation.InCellDropdown := True
        Xl.Selection.Validation.InputTitle := ""
        Xl.Selection.Validation.ErrorTitle := ""
        Xl.Selection.Validation.InputMessage := ""
        Xl.Selection.Validation.ErrorMessage := ""
        Xl.Selection.Validation.ShowInput := True
        Xl.Selection.Validation.ShowError := True
		XL.Workbooks(NomeF).ActiveSheet.Range("G1").Value := "Uso Militar ou Nuclear?"
		XL.Workbooks(NomeF).ActiveSheet.Range("G2").Select
		Xl.Selection.Validation.Delete
        Xl.Selection.Validation.Add(3,1,1,"Não`;Militar`;Nuclear")
        Xl.Selection.Validation.IgnoreBlank := True
        Xl.Selection.Validation.InCellDropdown := True
        Xl.Selection.Validation.InputTitle := ""
        Xl.Selection.Validation.ErrorTitle := ""
        Xl.Selection.Validation.InputMessage := ""
        Xl.Selection.Validation.ErrorMessage := ""
        Xl.Selection.Validation.ShowInput := True
        Xl.Selection.Validation.ShowError := True
		XL.Workbooks(NomeF).ActiveSheet.Range("A1:G1").Interior.Colorindex := 6
		XL.Workbooks(NomeF).ActiveSheet.Range("A1:G1").Font.Bold := True
		XL.Workbooks(NomeF).ActiveSheet.Range("2:2").Copy(XL.Workbooks(NomeF).ActiveSheet.Range("3:11"))
		XL.Workbooks(NomeF).ActiveSheet.Range("1:1").HorizontalAlignment := -4108
		XL.Workbooks(NomeF).ActiveSheet.Range("1:1").VerticalAlignment := -4108
		XL.Workbooks(NomeF).ActiveSheet.Range("1:1").WrapText := False
		XL.Workbooks(NomeF).ActiveSheet.Range("1:1").Orientation := 0
		XL.Workbooks(NomeF).ActiveSheet.Range("1:1").AddIndent := False
		XL.Workbooks(NomeF).ActiveSheet.Range("1:1").IndentLevel := 0
		XL.Workbooks(NomeF).ActiveSheet.Range("1:1").ShrinkToFit := False
		XL.Workbooks(NomeF).ActiveSheet.Range("1:1").ReadingOrder := -5002
		XL.Workbooks(NomeF).ActiveSheet.Range("1:1").MergeCells := False
		XL.Workbooks(NomeF).ActiveSheet.Range("A:A").ColumnWidth := 17
		XL.Workbooks(NomeF).ActiveSheet.Range("B:B").ColumnWidth := 13
		XL.Workbooks(NomeF).ActiveSheet.Range("C:C").ColumnWidth := 15
		XL.Workbooks(NomeF).ActiveSheet.Range("D:D").ColumnWidth := 20
		XL.Workbooks(NomeF).ActiveSheet.Range("E:E").ColumnWidth := 30
		XL.Workbooks(NomeF).ActiveSheet.Range("F:F").ColumnWidth := 15
		XL.Workbooks(NomeF).ActiveSheet.Range("G:G").ColumnWidth := 25
		XL.Workbooks(NomeF).ActiveSheet.Range("A2").Select
	}
	XL.Visible:=True
	WinWait ,%NomeF%
	WinMaximize ,%NomeF%
	Sleep ,1000
}

Info(Temp)
{	
	TicTac :=15
	SetTimer ,TimerButtonName,-10
	MsgBox ,4160,OLH-V - Informações,Evitar mexer no computador enquanto a Rotina é executada`n`nTecla "ESC" = Interrompe a Rotina.`n`nCertifique-se que o item 1 da Quote esteja selecionado antes de dar "OK" nessa mensagem.,15
	Sleep ,500
	SetTimer ,TimerButtonName,OFF
	If (Temp = 1)
		TempoIni()
	SetCapsLockState ,Off
	Sleep ,500
}

CriaContador()
{
	PosX := A_ScreenWidth - 211
	PosY := A_ScreenHeight - 126
	Gui ,New
	Gui ,Font,s8
	Gui ,Add,Text,w175 Center,Progresso:
	Gui ,Font,s16
	Gui ,Add,Text,w175 vContador Center,...
	Gui ,Font,s8
	Gui ,-Sysmenu -Border +AlwaysOnTop +Owner
	Gui ,Show,x%PosX% y%PosY%,Status OLH
}

DefPric(DPS)
{
	If DPS = 1
	{
		Sleep ,1500
		MouseClick ,left,530,180
		Esperar(250,1000)
		Send ,{Enter}
		Sleep ,1500
		Send ,+{PgUp}
		Esperar(250,1000)
	}
	Else
	{	
		Sleep ,1000
		Send ,!p
		Esperar(1000,1500)
		Send ,!p
		Sleep ,750
		Send ,!m	
		Esperar(1000,1500)
	}
}

Proximo(z,w)
{
	If (z=w)
		Send ,{Tab}
	Else
		Send ,{Down}
	Esperar(500,1000)
}

Esperar(EspA,EspD)
{
	Sleep ,%EspA%
	While A_Cursor = "Wait"
		Sleep ,500
	Sleep ,%EspD%
}

ActionsV()
{
	Send ,!a
	Sleep ,1500
	SendInput ,View
	Sleep ,1500
	Send ,{Enter}
	Esperar(250,1000)
}

SalvaESai(Rot)
{
	SetKeyDelay ,100
	Send ,{ALT}fs
	SetKeyDelay ,10
	Sleep ,1000
	Blockinput ,Off
	Clipboard = %CopiaeCola%
	TempoFim()
	TempoTotal()
	MsgBox ,4144,%Rot%,Rotina Finalizada!
	Exitapp
}

TempoIni()
{	
	PosX := A_ScreenWidth - 211
	PosY := A_ScreenHeight - 360
	Gui ,New
	Gui ,Font,s8
	Gui ,Add,Text,w175 Center,Horário de Inicio da Rotina
	Gui ,Font,s20
	FormatTime ,T1h,T12,hh
	FormatTime ,T1m,T12,mm
	FormatTime ,T1s,T12,ss
	Gui ,Add,Text,w175 Center,%T1h%:%T1m%:%T1s%
	Gui ,Font,s8
	Gui ,-Sysmenu -Border +AlwaysOnTop +Owner
	Gui ,Show,x%PosX% y%PosY%,IniTime
	
	PosY := A_ScreenHeight - 282
	Gui ,New
	Gui ,Font,s8
	Gui ,Add,Text,w175 Center,Horário de Finalização da Rotina
	Gui ,Font,s20
	FormatTime ,T2h,T12,hh
	FormatTime ,T2m,T12,mm
	FormatTime ,T2s,T12,ss
	Gui ,Add,Text,w175 Center,  :  :  
	Gui ,Font,s8
	Gui ,-Sysmenu -Border +AlwaysOnTop +Owner
	Gui ,Show,x%PosX% y%PosY%,FimTime
	
	PosY := A_ScreenHeight - 204
	Gui ,New
	Gui ,Font,s8
	Gui ,Add,Text,w175 Center,Tempo Total da Rotina
	Gui ,Font,s20
	Gui ,Add,Text,w175 Center,  :  :  
	Gui ,Font,s8
	Gui ,-Sysmenu -Border +AlwaysOnTop +Owner
	Gui ,Show,x%PosX% y%PosY%,TotalTime
	
	PosY := A_ScreenHeight - 126
	Gui ,New
	Gui ,Font,s8
	Gui ,Add,Text,w175 Center,Progresso
	Gui ,Font,s20
	Gui ,Add,Text,w175 vContador Center,  /  
	Gui ,Font,s8
	Gui ,-Sysmenu -Border +AlwaysOnTop +Owner
	Gui ,Show,x%PosX% y%PosY%,ProgressoTotal
}
	
TempoFim()
{
	PosX := A_ScreenWidth - 211
	PosY := A_ScreenHeight - 282
	Gui ,FimTime:New
	Gui ,Font,s8
	Gui ,Add,Text,w175 Center,Horário de Finalização da Rotina
	Gui ,Font,s20
	FormatTime ,T2h,T12,hh
	FormatTime ,T2m,T12,mm
	FormatTime ,T2s,T12,ss
	Gui ,Add,Text,w175 Center,%T2h%:%T2m%:%T2s%
	Gui ,Font,s8
	Gui ,-Sysmenu -Border +AlwaysOnTop +Owner
	Gui ,Show,x%PosX% y%PosY%,FimTime
}

TempoTotal()
{
	PosX := A_ScreenWidth - 211
	PosY := A_ScreenHeight - 204
	Gui ,TotalTime:New
	Gui ,Font,s8
	Gui ,Add,Text,w175 Center,Tempo Total da Rotina
	Gui ,Font,s20
	T1h := T1h*3600
	T2h := T2h*3600
	T1m := T1m*60
	T2m := T2m*60
	T1s := T1h+T1m+T1s
	T2s := T2h+T2m+T2s
	T3s := T2s-T1s
	T3h := T3s//3600
	T3m := (T3s - T3h*3600)//60
	T3s := T3s - T3m*60 - T3h*3600
	If T3h<10
		T3h = 0%T3h%
	If T3m<10
		T3m = 0%T3m%
	If T3s<10
		T3s = 0%T3s%
	Gui ,Add,Text,w175 Center,%T3h%:%T3m%:%T3s%
	Gui ,Font,s8
	Gui ,-Sysmenu -Border +AlwaysOnTop +Owner
	Gui ,Show,x%PosX% y%PosY%,TotalTime
}

LinkOracle(Funcao,Guardar)
{
	If ((i = iini) or (Guardar = 0) or (Guardar = 2))
	{
		
		Links := wb.Document.Links
		p := 0
		Loop % Links.Length ;%
		{	
			If (Links[p].InnerText = Funcao)
			{
				If (Guardar = 1)
					PagIni := Links[p].href
				Else
				{
					If (Guardar = 2)
					{
						PagAss11 := wb.Document.All("N11:EgoOrganizationCode:0").InnerText
						PagAss12 := wb.Document.All("N12:EgoOrganizationCode:0").InnerText
					}
					Links[p].Click()
				}
				Break
			}
			p := p + 1
		}
	}
	Else
		Wb.Navigate(PagIni)
	If (Guardar = 0)
	{
		IELoad(wb)
		Sleep ,1000
	}
	Else If (Guardar = 2)
	{
		While (((wb.Document.All("N11:EgoOrganizationCode:0").InnerText = PagAss11) and (PagAss11 <> "")) or ((wb.Document.All("N12:EgoOrganizationCode:0").InnerText = PagAss12)  and (PagAss12 <> "")))
			Sleep ,100
		Sleep ,1000
	}
}

ClicaBotao(Texto,Tempo)
{
	If (Texto = "Update")
		LinkImg := "ITv_"
	If (Texto = "AssToOrg")
		LinkImg := "bEgoItemAssignedToOrgButton"
	If (Texto = "Add")
		LinkImg := "bAdd_N-N6"

	Images := wb.Document.getElementsByTagName("IMG")
	p := 0
	Loop % Images.Length ;%
	{	
		Imagem := Images[p].src
		StringGetPos ,Botao,Imagem,%LinkImg%
		If (Botao > 0)
		{
			Images[p].Click()
			Break
		}
		p := p + 1
	}
	If (Tempo = 1)
	{
		IELoad(wb)
		Sleep ,1000
	}
	Else
		Esperar(500,1000)
}

TrocarOrg(Org)
{
	p := 0
	Links := wb.Document.Links
	Loop % Links.Length ;%
	{	
		If (Links[p].InnerText = "Change Organization (EMR APC Item Creation BR)")
		{
			Links[p].Click()
			Break
		}
		p := p + 1
	}
	IELoad(wb)
	Sleep ,1000
	p := 0
	Loop ,20
	{
		If (wb.Document.All("N11:orgCode:" . p).InnerText = Org)
		{
			; wb.Document.All("N11:N31:" . p).Checked := True
			; Sleep ,1000
			; Send ,!p
			wb.Document.All("N11:N31:" . p).Focus
			SetKeyDelay ,200
			Send ,{TAB}{ENTER}
			SetKeyDelay ,10
			Break
		}
		p := p + 1
	}
	IELoad(wb)
	Sleep ,1000
}

;Função para trocar o idioma do Oracle
TrocarLingua(Ling)
{	
	wb.Navigate(PagOracle)
	IELoad(wb)
	Sleep ,500
	LinkOracle("Preferences",0)
	If (Ling = "US")
		wb.Document.All.CurrentLanguage.value := "US`;AMERICAN"
	Else If (Ling = "BR")
		wb.Document.All.CurrentLanguage.value := "PTB`;BRAZILIAN PORTUGUESE"
	Sleep ,500
	Ling1 := wb.Document.All.CurrentLanguage.value
	If (((Ling = "US") and(Ling1 <> "US`;AMERICAN")) or ((Ling = "BR") and(Ling1 <> "PTB`;BRAZILIAN PORTUGUESE")))
	{
		MsgBox ,4112,OLH - Erro,Foi encontrado um erro durante a execução da Rotina.`nPor favor, execute novamente.
		Gosub Sair2
	}
	Send ,!p
	IELoad(wb)
	Sleep ,500
	wb.Navigate(PagOracle)
	IELoad(wb)
	Sleep ,500
}

PesquisarItem(Cod,Clica)
{
	wb.Document.All.searchField1.Focus()
	Sleep ,200
	Clipboard := Cod
	Sleep ,200
	Send ,^v
	Sleep ,200
	Send ,{TAB}
	Sleep ,200
	Send ,{Enter}
	IELoad(wb)
	Sleep ,500
	Loop
	{
		Encontrado := 0
		ProPag := 0
		q := 0
		LinksN := wb.Document.Links
		Loop % LinksN.Length ;%
		{	
			If (LinksN[q].InnerText = "Next 25")
			{
				ProPag := 1
				Break
			}
			q := q + 1
		}
		p := 0
		Links := wb.Document.Links
		Loop % Links.Length ;%
		{	
			If (Links[p].InnerText = Cod)
			{
				Encontrado := 1
				Break
			}
			p := p + 1
		}
		If ((Encontrado = 1) and (Clica = 1))
		{
			Links[p].Click()
			IELoad(wb)
			Sleep ,500
			Break
		}
		Else
		{
			If (Encontrado = 0)
			{
				XL.Workbooks(Nome).ActiveSheet.Range("A"i).Interior.ColorIndex := 3
				Erro := Erro + 1
				Break
			}
			Else
				Break
			If (ProPag = 1)
			{
				LinksN[q].Click()
				Esperar(200,1000)
			}
		}
	}
}

IEGet(Name="")
{
	IfEqual ,Name,,WinGetTitle,Name,ahk_class IEFrame
	Name := (Name="New Tab - Windows Internet Explorer")? "about:Tabs":RegExReplace(Name, " - (Windows|Microsoft)? ?Internet Explorer$")
	For wbget in ComObjCreate("Shell.Application").Windows()
	If wbget.LocationName=Name and InStr(wbget.FullName, "iexplore.exe")
		Return wbget
}

IELoad(wbload)
{
    If !wbload
        Return False
    Loop
	{
        Sleep,100
		If (A_Index > 50)
			Send ,{F5}
    } Until (wbload.busy)
    Loop
        Sleep,100
    Until (!wbload.busy)
    Loop
        Sleep,100
    Until (wbload.Document.Readystate = "Complete")
	Return True
}

Pause::Pause

Sair2:
ExitApp
	
Esc::
Sair:
Suspend On
MsgBox ,4132,OLH-V - Interrupção,Programa interrompido.`nDeseja voltar para o Menu Principal?
IfMsgBox Yes
    Gosub MenuPHM
Else
{
	If (Nome <> "")
	{
		XL.Workbooks(Nome).Save
		XL.Quit()
	}
    ExitApp
}