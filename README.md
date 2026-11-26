# Folio — Conversor Universal de Documentos

**Acesse online:** [https://engfranciscocamara-tech.github.io/folio/](https://engfranciscocamara-tech.github.io/folio/)

O **Folio** é um conversor universal de documentos **100% local e gratuito** que roda inteiramente no navegador — seus arquivos nunca saem do seu dispositivo.

## Formatos suportados (42 combinações)

**PDF ⇄ EPUB ⇄ DOCX ⇄ Markdown (.md) ⇄ HTML ⇄ RTF ⇄ TXT**

## Principais recursos

- **PDF → EPUB 3 de qualidade editorial:** mede fonte mediana, entrelinha real e margens para reconstruir parágrafos (inclusive com recuo de 1ª linha sem espaço extra), diálogos, negrito/itálico, 2 colunas, capítulos e sumário (`nav.xhtml` + `toc.ncx` compatível com Kindle, Kobo e Apple Books).
- **Extração real de imagens:** detecta e recorta figuras, gráficos e capas mantendo a posição exata no fluxo de leitura (embutidas em EPUB, DOCX, HTML e PDF).
- **Barra de progresso em tempo real:** mostra página atual, imagens extraídas, tempo estimado (ETA) e otimização de memória para PDFs grandes.
- **Modo Claro e Escuro:** interface minimalista inspirada no macOS, com prévia de leitura e inspeção do pipeline e da Representação Intermediária (IR).

## Motor Python local opcional (para PDFs gigantes 40+ MB)

```powershell
# Converter direto pelo terminal:
python converter_local.py "livro.pdf" "livro.epub"

# Ou conectar ao Folio em modo local turbinado (PyMuPDF em C):
python converter_local.py --serve
```
