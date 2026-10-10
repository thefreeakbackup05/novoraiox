# Novo Raio-X de Investimentos

Aplicativo Windows para converter documentos Word `.docx` em PDF em lote usando o Microsoft Word instalado no computador. A exportação nativa do Word preserva a paginação, tabelas, gráficos e demais elementos do documento original.

## Requisitos
- Windows 10/11
- Microsoft Word instalado e ativado
- Executável compilado pelo GitHub Actions ou Python 3.11+ para executar o código-fonte

## Uso
1. Coloque os arquivos `.docx` na pasta `WORD` ao lado do executável (ou selecione outra pasta na interface).
2. Execute `NovoRaioX.exe`.
3. Selecione a pasta de saída, por padrão `PDF`.
4. Clique em **Converter documentos**. Os arquivos originais não são alterados; cada documento é exportado para um PDF com o mesmo nome.
5. Os logs são gravados na pasta `logs`.

Arquivos temporários do Word (com prefixo `~$`) são ignorados. Um erro em um documento não interrompe o lote.

## Executar pelo código-fonte
```powershell
py -m pip install -r requirements.txt
py app.py
```

## Build
O workflow `.github/workflows/build.yml` compila o executável para Windows e publica o EXE como artefato do GitHub Actions.

**Importante:** o computador que executa o aplicativo precisa ter o Microsoft Word instalado. O build automatizado valida o empacotamento, não substitui um teste funcional em uma máquina com Word.
