# Disparo de e-mails pelo Outlook

Aplicativo local em PowerShell/WinForms para enviar e-mails pelo Outlook instalado no Windows, usando uma planilha Excel/CSV e anexos encontrados automaticamente em uma pasta.

## Como usar

1. Abra `Abrir Disparo Outlook.cmd`.
2. Selecione a planilha, como `acre -email (2).xlsx`.
3. Confirme a aba e as colunas. Para o modelo enviado, o sistema detecta:
   - `Cidade`
   - `Nome`
   - `Email`
4. Em `Módulo de anexos`, escolha o fluxo:
   - `Resumo Município-UF`: usa arquivos soltos na pasta, encontrados pelo nome do município/UF.
   - `Distribuição por subpastas`: usa uma pasta raiz, como `C:\Users\jonatas.chaves\Downloads\RelatoriosDetalhado\AM`, e anexa todos os arquivos encontrados dentro da subpasta de cada município.
5. Selecione a pasta onde estão os documentos.
6. Selecione os dois arquivos de informativo quando eles também precisarem ser anexados em todos os e-mails.
7. No modo `Resumo Município-UF`, o nome dos arquivos por município deve conter o município e, quando possível, a UF, por exemplo:
   - `ACRELANDIA-AC.pdf`
   - `ACRELANDIA-AC-resumo.xlsx`
   - `ACRELANDIA-AC-informativo.pdf`
8. No modo `Distribuição por subpastas`, a pasta raiz deve conter uma subpasta por município, por exemplo:
   - `AM\ALVARAES\arquivo1.pdf`
   - `AM\ALVARAES\arquivo2.xlsx`
   - `AM\BARREIRINHA\arquivo1.pdf`
9. Confira a pré-visualização.
10. Use `Criar rascunhos` para revisar antes do envio ou `Disparar e-mails` para enviar diretamente.

## Regras principais

- O envio usa a sessão local do Outlook via COM.
- No modo `Resumo Município-UF`, um coordenador pode receber mais anexos quando houver mais de um arquivo Município-UF correspondente.
- No modo `Distribuição por subpastas`, todos os arquivos da subpasta do município são anexados, inclusive quando houver mais de um arquivo.
- Os informativos selecionados são anexados em todos os e-mails do lote.
- A pré-visualização mostra `Sem anexo` quando a linha da planilha não encontrou nenhum arquivo correspondente.
- O texto padrão do e-mail é enviado em HTML.
- A seção `Editor do e-mail` permite escrever e formatar visualmente o corpo da mensagem com negrito, itálico, sublinhado, listas, alinhamento, link, desfazer/refazer e limpeza de formatação.
- O botão `HTML` abre o código bruto para colar ou ajustar HTML manualmente quando necessário.
- O editor abre em branco. Ao colar texto já formatado, ele preserva a formatação quando o Windows/Outlook/Word disponibilizar HTML na área de transferência.
- Ao colar texto simples pelo botão `HTML`, marcações como `**negrito**`, `__sublinhado__`, `*itálico*`, links e quebras de linha são convertidas automaticamente para HTML.
- A assinatura da sessão atual do Outlook é obrigatória e será adicionada ao final de cada e-mail.

## Campos variáveis

No assunto e no HTML, podem ser usados:

- `{{Nome}}`
- `{{Coordenador}}`
- `{{Cidade}}`
- `{{Municipio}}`
- `{{Município}}`
- `{{UF}}`
- Qualquer coluna da planilha, no formato `{{NomeDaColuna}}`
