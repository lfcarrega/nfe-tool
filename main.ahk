#Requires AutoHotkey v2.0

Estados := ["AC:12","AL:27","AM:13","AP:16","BA:29","CE:23","DF:53","ES:32","GO:52","MA:21","MG:31","MS:50","MT:51","PA:15","PB:25","PE:26","PI:22","PR:41","RJ:33","RN:24","RO:11","RR:14","RS:43","SC:42","SE:28","SP:35","TO:17"]
Modelos := ["55 - NFe", "65 - NFCe", "59 - SAT CFe", "57 - CTe", "58 - MDFe"]
Formas := ["1 = Emissão normal (não em contingência)", "2 = Contingência FS-IA, com impressão do DANFE em formulário de segurança", "3 = Contingência SCAN (Sistema de Contingência do Ambiente Nacional)", "4 = Contingência DPEC (Declaração Prévia da Emissão em Contingência)", "5 = Contingência FS-DA, com impressão do DANFE em formulário de segurança", "6 = Contingência SVC-AN (SEFAZ Virtual de Contingência do AN)", "7 = Contingência SVC-RS (SEFAZ Virtual de Contingência do RS)"]

Main := Gui(, "Chave de Acesso")
;
Tab := Main.Add("Tab3",, ["Gerador e Extrator"])

; Campos do Estado
Main.Add("Text", "Section", "Estado:")
Estado := Main.Add("DropDownList", "xs w60", Estados)
Estado.GetPos(&EstadoX, &EstadoY, &EstadoW, &EstadoH)

; Campos da data
Main.Add("Text", "xs", "Data:")
Date := Main.Add("DateTime", "xs w100", "ShortDate")

; Campos do CNPJ
Main.Add("Text", "xs", "CNPJ:")
CNPJ := Main.Add("Edit", "xs w120 Limit18", "")

; Campos do MODELO
Main.Add("Text", "xs", "Modelo:")
Modelo := Main.Add("DropDownList", "xs w100", Modelos)

; Campos da SÉRIE
Main.Add("Text", "xs", "Série:")
Serie := Main.Add("Edit", "Number xs w40 Limit3",)

; Campos do número da NF-e
Main.Add("Text", "xs", "Número:")
Numero := Main.Add("Edit", "Number xs w60 Limit9",)

; Campos da forma da emissão
Main.Add("Text", "xs", "Forma da emissão:")
Forma := Main.Add("DropDownList", "xs w380", Formas)

; Campos do código numérico
Main.Add("Text", "xs", "Código numérico (deixe em branco pra gerar um código aleatório):")
Codigo := Main.Add("Edit", "Number xs w60 Limit8",)

; Campos do dígito verificador (DV)
Main.Add("Text", "xs ", "Dígito de Verificação (DV):")
DV := Main.Add("Edit", "Number xs w20 Limit1", "")

; Campos da chave de acesso
Main.Add("Text", "xs", "Chave de Acesso:")
Chave := Main.Add("Edit", "Number xs w280 Number Limit44", "")

; Botão pra gerar a chave
Main.Add("Button", "xs", "Gerar").OnEvent("Click", GerarChave)

; Botão pra gerar o DV
Main.Add("Button", "yp", "Gerar DV").OnEvent("Click", CalculaDV2)

Main.Add("GroupBox", "xs w165 Section", "Extrair:")

; Botão pra extrair da chave de acesso
Main.Add("Button", "xs+7 ys+20", "Chave de Acesso").OnEvent("Click", Extrair)

; Botão pra extrair do XML
Main.Add("Button", "yp", "XML...").OnEvent("Click", ExtrairXML)

; Exibe a janela
Main.Show

XmlProcessor() {
; Exemplo
; 	For k, in doc.selectNodes("//det") ; este método seleciona o conjunto de nós que contém as informações de cada produto da nota fiscal
 ; {
 ;        xProd       := k.selectSingleNode("./prod/xProd").text ; Utilizando a notação do ponto na expressão Xpath não é necessário escrever a expressão completa 
 ;        NCM         := k.selectSingleNode("./prod/NCM").text
 ;        CFOP        := k.selectSingleNode("./prod/CFOP").text
 ;        CST           := k.selectSingleNode("./imposto/ICMS/*/CST | ./imposto/ICMS/*/CSOSN").text
 ;        qCom        := Format("{:.2f}", k.selectSingleNode("./prod/qCom").text)
 ;        vProd       := Format("{:.2f}", k.selectSingleNode("./prod/vProd").text) ; a função Format já retorna a quantidade em formato mais amigável
 
	xml := FileRead("xml.xml", "CP65001")

	xmlp := ComObject("MSXML2.DOMDocument.3.0")
	xmlp.async := false
	xmlp.loadXML(xml)

	MSGBOX xmlp.selectSingleNode("//infNFe/ide/cUF").text

}

ExtrairXML(*) {
	local SelXML := FileSelect(,,, "*.xml")
	If (SelXML) {
		Try {
			local XMLFile := FileRead(SelXML, "CP65001")
			xmlp := ComObject("MSXML2.DOMDocument.3.0")
			xmlp.async := false
			xmlp.loadXML(XMLFile)
			Chave.Value := xmlp.selectSingleNode("//protNFe/infProt/chNFe").text
		} Catch {
			MsgBox "Arquivo inválido!", "Erro", "IconX"
			Return
		}
		Extrair()
	} Else {
		Return
	}
}

ValidaCodNum(CodigoNum) {
	If CodigoNum != "" {
		If CodigoNum == Numero.Value or CodigoNum == 00000000 or CodigoNum == 11111111 or CodigoNum == 22222222 or CodigoNum == 33333333 or CodigoNum == 44444444 or CodigoNum == 55555555 or CodigoNum == 66666666 or CodigoNum == 77777777 or CodigoNum == 88888888 or CodigoNum == 99999999 or CodigoNum == 12345678 or CodigoNum == 23456789 or CodigoNum == 34567890 or CodigoNum == 45678901 or CodigoNum == 56789012 or CodigoNum == 67890123 or CodigoNum == 78901234 or CodigoNum == 89012345 or CodigoNum == 90123456 or CodigoNum == 01234567
			MsgBox "Código numérico inválido!", "Aviso", "Icon!"
	}
	Else {
		CodigoNum := ""
		Loop {
			Aleatorio := Random(9)
			CodigoNum .= Aleatorio
		} Until StrLen(CodigoNum) = 8 and CodigoNum != Numero.Value and CodigoNum != 00000000 and CodigoNum != 11111111 and CodigoNum != 22222222 and CodigoNum != 33333333 and CodigoNum != 44444444 and CodigoNum != 55555555 and CodigoNum != 66666666 and CodigoNum != 77777777 and CodigoNum != 88888888 and CodigoNum != 99999999 and CodigoNum != 12345678 and CodigoNum != 23456789 and CodigoNum != 34567890 and CodigoNum != 45678901 and CodigoNum != 56789012 and CodigoNum != 67890123 and CodigoNum != 78901234 and CodigoNum != 89012345 and CodigoNum != 90123456 and CodigoNum != 01234567
	}
	Return SubStr("00000000" . CodigoNum, -8)
}

Extrair(*) {
	local EstadoMatched := 0
	If StrLen(Chave.Value) != 44
		MsgBox "Chave de acesso inválida ou não preenchida!", "Erro", "IconX"
	Else {
		local ExtEstado := SubStr(Chave.Value, 1, 2)
		For K, V in Estados {
			If ExtEstado = SubStr(V, 4, 2) {
				Estado.Choose(K)
				local EstadoMatched := 1
			}
		}
		If EstadoMatched = 0
			MsgBox "Estado inválido!", "Erro", "IconX"
		local ExtAno := SubStr(Chave.Value, 3, 2)
		local ExtMes := SubStr(Chave.Value, 5, 2)
		Date.Value := "20" ExtAno ExtMes "01000000"
		CNPJ.Value := SubStr(Chave.Value, 7, 14)
		local ModeloMatched := 0
		local ExtModelo := SubStr(Chave.Value, 21, 2)
		For K, V in Modelos {
			If SubStr(V, 1, 2) = ExtModelo {
				Modelo.Choose(V)
				local ModeloMatched := 1
			} 
		}
		Serie.Value := SubStr(Chave.Value, 23, 3)
		Numero.Value := SubStr(Chave.Value, 26, 9)
		local FormaMatched := 0
		local ExtForma := SubStr(Chave.Value, 35, 1)
		For K, V in Formas {
			If SubStr(V, 1, 1) = ExtForma {
				Forma.Choose(V)
				local FormaMatched := 1
			} 
		}
		Codigo.Value := SubStr(Chave.Value, 36, 8)
		DV.Value := SubStr(Chave.Value, 44, 1)
	}
}

CalculaDV2(*) {
	If StrLen(Chave.Value) != 43 {
		MsgBox "Chave de acesso inválida ou com o DV já preenchido!", "Erro", "IconX"
	}
	Else
		Chave.Value .= CalculaDV(Chave.Value)
}

CalculaDV(ChaveSemDV) {
	local ChaveAcessoNum := Array()
	local ChaveAcessoMult := Array()
	local ChaveAcessoPos := 1
	local ChaveAcessoMultStart := 2
	local ChaveAcessoLenStrip := 1
	local ChaveAcessoMultSum := 0
	local DV_Calculado := 0
	If StrLen(ChaveSemDV) != 43 {
		MsgBox "Chave de acesso inválida!", "Erro", "IconX"
	}
	Else {
		Loop 43 {
			ChaveAcessoNum.Push(SubStr(ChaveSemDV, -ChaveAcessoLenStrip, 1))
			ChaveAcessoInvertida .= SubStr(ChaveSemDV, ChaveAcessoLenStrip, 1) " "
			ChaveAcessoLenStrip++
		}
		Loop {
			ChaveAcessoMult.Push(ChaveAcessoNum[ChaveAcessoPos]*ChaveAcessoMultStart)
			ChaveAcessoPos++
			ChaveAcessoMultStart++
			If ChaveAcessoMultStart > 9
				ChaveAcessoMultStart := 2
		} Until ChaveAcessoPos = 44
		For K, V in ChaveAcessoMult {
			ChaveAcessoMultSum += V
		}
		If Mod(ChaveAcessoMultSum, 11) = 0 or Mod(ChaveAcessoMultSum, 11) = 1
			DV_Calculado := 0
		Else
			DV_Calculado := 11-Mod(ChaveAcessoMultSum, 11)
	}
	Return DV_Calculado
}

GerarChave(*) {
	Chave.Value := ""
	If Estado.Text
		Chave.Value .= SubStr(Estado.Text, 4, 2)
	Else
		MsgBox "Estado inválido ou não preenchido!", "Aviso", "Icon!"
	Chave.Value .= FormatTime(Date.Value, "yy")
	Chave.Value .= FormatTime(Date.Value, "MM")
	If StrLen(RegExReplace(CNPJ.Value, "\D")) != 14
		MsgBox "CNPJ inválido ou não preenchido!", "Aviso", "Icon!"
	If CNPJ.Value
		Chave.Value .= RegExReplace(CNPJ.Value, "\D")
	If Modelo.Value 
		Chave.Value .= SubStr(Modelo.Text, 1, 2)
	Else
		MsgBox "Modelo não preenchido!", "Aviso", "Icon!"
	If Serie.Value {
		Serie.Value := SubStr("000" . Serie.Value, -3)
		Chave.Value .= Serie.Value
	} Else
		MsgBox "Série não preenchida!", "Aviso", "Icon!"
	If Numero.Value {
		Numero.Value := SubStr("000000000" . Numero.Value, -9)
		Chave.Value .= Numero.Value
	} Else
		MsgBox "Número não preenchido!", "Aviso", "Icon!"
	If Forma.Value 
		Chave.Value .= SubStr(Forma.Text, 1, 1)
	Else
		MsgBox "Forma da emissão não preenchida!", "Aviso", "Icon!"
	Codigo.Value := ValidaCodNum(Codigo.Value)
	Chave.Value .= Codigo.Value
	DV.Value := CalculaDV(Chave.Value)
	Chave.Value .= DV.Value
}