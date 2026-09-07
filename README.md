# Gestão de Tolva — Goiás Verde Alimentos

Painel de controle de troca de filtros das tolvas, publicado via GitHub Pages.

## Como funciona

- A página (`index.html`) lê `data/latest.json` e calcula **Próxima troca**,
  **Dias restantes** e **Status** direto no navegador, usando a data de hoje —
  igual à fórmula da planilha Excel (`=PRÓXIMA-HOJE()`).
- Isso significa que o site se atualiza **sozinho todos os dias**, sem precisar
  reprocessar nada. Só é necessário subir a planilha de novo quando uma
  **troca de filtro realmente acontecer** (mudando a data "última troca").

## Fluxo do dia a dia (quando uma troca acontece)

1. Atualize e salve a planilha `GESTÃO_TOLVA` normalmente no Excel.
2. No GitHub, vá até a pasta `uploads/`.
3. Suba o arquivo `.xlsm`, **substituindo** o arquivo anterior (mesmo nome).
4. Escreva uma mensagem de commit e clique em **Commit changes**.
5. Aguarde ~1 minuto: o GitHub Actions processa a planilha e atualiza
   `data/latest.json` automaticamente.
6. Confira o resultado na página publicada.

## Estrutura

```
index.html                     -> a página do painel
scripts/parse_excel.py         -> lê o Excel e gera o JSON
uploads/                       -> onde você sobe o .xlsm
data/latest.json               -> dados usados pela página (gerado automaticamente)
.github/workflows/process.yml  -> roda o script sempre que um arquivo é enviado a uploads/
```

## Configuração inicial (primeira vez)

1. Crie o repositório no GitHub e suba estes arquivos.
2. **Atenção:** a pasta `.github/workflows/` começa com ponto — se você usar
   drag-and-drop pelo Windows, ela pode não aparecer/ser enviada. Se isso
   acontecer, crie o arquivo manualmente pela interface do GitHub
   ("Add file" → "Create new file" → digite o caminho completo
   `.github/workflows/process.yml`).
3. Em **Settings → Pages**, selecione a branch principal (`main`) e pasta raiz (`/`).
4. Em **Settings → Actions → General → Workflow permissions**, marque
   **"Read and write permissions"** (necessário para o Actions conseguir
   salvar o `data/latest.json` de volta no repositório).
5. Suba o arquivo `.xlsm` inicial em `uploads/` para gerar o primeiro `data/latest.json`.
