# NF-e Tool

Ferramenta desktop para geração e extração de chaves de acesso de documentos fiscais eletrônicos brasileiros (NF-e, NFC-e, SAT CF-e, CT-e e MDF-e), desenvolvida em AutoHotkey v2.

![Demonstração](https://github.com/lfcarrega/nfe-tool/blob/main/demo.gif)

---

## Funcionalidades

- **Gerar chave de acesso** a partir dos campos do documento (UF, data, CNPJ, modelo, série, número, forma de emissão e código numérico)
- **Calcular o DV** (dígito verificador) pelo módulo 11 conforme especificação da SEFAZ
- **Extrair campos** a partir de uma chave de acesso existente, preenchendo o formulário automaticamente
- **Extrair a partir de XML** — abre um arquivo `.xml` de NF-e e extrai a chave de acesso diretamente do nó `chNFe`
- **Geração de código numérico aleatório** com validação dos códigos proibidos pela especificação

---

## Modelos suportados

| Código | Documento |
|--------|-----------|
| 55 | NF-e |
| 65 | NFC-e |
| 59 | SAT CF-e |
| 57 | CT-e |
| 58 | MDF-e |

---

## Requisitos

- Windows
- [AutoHotkey v2](https://www.autohotkey.com/)

---

## Como usar

### Gerando uma chave

Preencha os campos na interface e clique em **Gerar**. O código numérico pode ser deixado em branco, a ferramenta gera um aleatório válido automaticamente. O DV é calculado e preenchido junto.

### Extraindo de uma chave existente

Cole a chave de acesso no campo correspondente e clique em **Chave de Acesso** na seção *Extrair*. Os campos são preenchidos automaticamente.

### Extraindo de um XML

Clique em **XML...**, selecione o arquivo `.xml` da NF-e e a ferramenta extrai a chave e preenche os campos.

---

## Referência

Baseado na especificação técnica da NF-e disponível no [Portal da Nota Fiscal Eletrônica](https://www.nfe.fazenda.gov.br/portal/listaConteudo.aspx?tipoConteudo=BMPFMBoln3w=).
