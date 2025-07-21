# XMLComparer

Programa desktop feito em **Windows Forms** para comparar e validar arquivos XML de documentos fiscais com os dados armazenados em banco de dados SQL Server.

---

## Funcionalidades

1. **Comparação automática de XMLs fiscais**  
   Compara arquivos XML de NFes, CFes e NFCes (Notas Fiscais eletrônicas, Cupom Fiscal Eletrônico e Nota Fiscal de Consumidor Eletrônica) com os dados dos mesmos XMLs guardados no banco de dados. O programa identifica automaticamente o tipo do documento fiscal.

2. **Validação de valores e quantidades**  
   Ao comparar os documentos, calcula o valor total dos XMLs e verifica se o valor total armazenado no banco está correto. Também identifica se a quantidade de XMLs emitidos é igual à quantidade registrada no banco de dados.

3. **Reemissão de XMLs de CFes**  
   Permite a reemissão dos XMLs de CFes, criando os arquivos XML com base nas informações armazenadas no banco de dados.

---

## Tecnologias utilizadas

- Windows Forms (C#)
- SQL Server

---

## Como usar

1. Abra o programa no Windows.
2. Selecione a pasta com os arquivos XML para comparação.
3. O programa irá identificar automaticamente os tipos de documentos fiscais e iniciar a comparação.
4. Veja o relatório de diferenças no valor total e na quantidade de documentos.
5. Para reemitir os XMLs de CFes, use a função específica no programa.

---

## Contribuições

Contribuições são bem-vindas! Para sugestões, abra uma issue ou envie um pull request.
