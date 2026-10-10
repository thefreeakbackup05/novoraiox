# Novo Raio-X de Investimentos — versão Word

Aplicativo Windows para gerar PDFs a partir de documentos Word `.docx`, com configurações de nome, hiperlink condicional e rodapé. **Não usa Excel como origem.** A exportação do conteúdo é feita pelo Microsoft Word instalado no computador, preservando a formatação que já existe no documento de origem.

## Requisitos
- Windows 10/11
- Microsoft Word instalado e ativado
- Para rodar o código: Python 3.11+

## Uso
1. Extraia o artefato `NovoRaioX-Windows` do GitHub Actions.
2. Coloque cópias dos documentos Word `.docx` na pasta `WORD` ao lado do executável (ou selecione outra pasta).
3. Execute `NovoRaioX.exe`.
4. Configure as pastas, o padrão do nome, os dados do hiperlink e o rodapé.
5. Clique em **GERAR PDFs**. Os documentos Word originais não são alterados.
6. Consulte o resultado na janela e os logs em `logs`.

Arquivos temporários do Word (prefixo `~$`) são ignorados. Um erro em um documento não interrompe o lote.

## Configurações disponíveis
- **Nome do PDF:** padrão configurável com tokens `[NOME DO PLANO]`, `[NOME DO PLANO NO WORD]`, `[MÊS DO RAIO X]`, `[ANO]` e `[ARQUIVO DE ORIGEM]`. Nome do plano e competência podem ser informados manualmente; vazios, o programa tenta detectá-los no texto do Word.
- **Hiperlink:** URL, palavras indicadoras de vídeo e texto clicável configuráveis. A tentativa de inserir o link só ocorre se um indicador estiver presente no documento.
- **Rodapé da última página:** texto institucional, rótulo e data `dd/mm/aaaa` configuráveis. A aplicação desenha o rodapé na última página do PDF.
- **Sobrescrita:** por padrão, não substitui PDFs existentes.

## Fontes e diagramação
A exportação usa o Word para preservar fontes, estilos, tabelas, gráficos, imagens e paginação presentes no DOCX. O gerador não recria nem corrige automaticamente elementos ausentes do documento de origem. Para manter a identidade visual acordada, os DOCX de entrada precisam estar corretamente formatados com Campton SemiBold nos títulos/destaques, Campton Light nos subtítulos e Open Sans no corpo; nos seis blocos “Ativos que compõem uma carteira”, Open Sans Bold Italic nos títulos e Open Sans Light Italic no corpo, sem negrito. O gráfico de evolução do patrimônio líquido, tabelas, notas e asteriscos precisam existir corretamente no documento de origem.

## Validação e limitações conhecidas
O programa verifica se o PDF foi criado, se tem páginas e se há texto extraível. Isso **não substitui a revisão visual** de cortes, sobreposições, gráficos e fontes. O Word precisa estar instalado para gerar o PDF. A detecção automática do plano e da competência é heurística; se o texto do documento não for reconhecido, preencha os campos correspondentes na interface. O hiperlink é aplicado à última ocorrência da palavra configurada quando há um indicador de vídeo; confira os resultados com um documento real antes de uso em produção.

## Executar pelo código-fonte
```powershell
py -m pip install -r requirements.txt
py app.py
```

## Build
O workflow `.github/workflows/build.yml` compila o executável para Windows e publica o ZIP/artefato do GitHub Actions. A compilação automatizada valida sintaxe e empacotamento; não substitui teste funcional em um computador com Microsoft Word e documentos reais.
