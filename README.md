# Guia patrimonial

Aplicativo desktop em Python para preencher automaticamente guias Word de entrega e recebimento.

- `Guia solano Entrega.docx`
- `Guia solano Recebimento2.docx`

## O que o programa preenche

- o campo de remetente na guia de entrega
- o campo de receptor na guia de recebimento
- a tabela com numero de patrimonio e descricao do item
- voce escolhe se quer gerar somente entrega, somente recebimento ou os dois

## Como usar

1. Execute:

```powershell
python main.py
```

2. Na interface:

- confira ou selecione os arquivos `.docx`
- escolha o tipo de guia que quer gerar
- informe o remetente da entrega quando precisar
- informe o receptor do recebimento quando precisar
- preencha os itens da tabela
- clique em `Gerar documentos`

3. Os arquivos finais serao salvos em `saida/`

## Limite atual

Os modelos enviados possuem 22 linhas para itens. Se voce preencher mais do que isso, o programa avisa.
