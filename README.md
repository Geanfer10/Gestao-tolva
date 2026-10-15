# Gestão de Tolva — Goiás Verde Alimentos

Painel de controle de troca de filtros das tolvas, publicado via GitHub Pages:
https://geanfer10.github.io/Gestao-tolva/

## Como funciona

- A página (`index.html`) lê `data/latest.json` e calcula **Próxima troca**,
  **Dias restantes** e **Status** direto no navegador, usando a data de hoje —
  igual à fórmula da planilha Excel (`=PRÓXIMA-HOJE()`).
- Por isso o painel muda **sozinho todos os dias**. A planilha só precisa chegar
  ao GitHub quando algo muda nela (troca de filtro, OS, estoque).
- Quem leva a planilha até o GitHub é o **envio automático** instalado no
  computador do PCM (pasta `envio-automatico/`): a cada 5 minutos ele confere se
  a planilha foi salva e, se mudou, envia para `uploads/GESTAO_TOLVA.xlsm`.
  O GitHub Actions reprocessa e o painel atualiza em 1 a 2 minutos.

## Dia a dia

1. Abra a planilha `GESTÃO TOLVA.xlsm` no Excel, faça as alterações e **salve**.
2. Pronto. Em até uns 7 minutos o painel mostra a versão nova.

Precisa que vá na hora? Dê dois cliques em `enviar-agora.bat`.
Histórico dos envios: `%LOCALAPPDATA%\TolvaSync\envios.log`.

O computador precisa estar ligado e com o seu usuário conectado. Se ele estiver
desligado, o envio acontece quando você entrar de novo.

## Instalar o envio automático (uma vez por computador)

### 1. Criar o token do GitHub

O token é a "senha" que deixa o programa enviar a planilha. Ele só dá acesso a
este repositório.

1. No GitHub, clique na sua foto → **Settings** → **Developer settings** (no fim
   do menu à esquerda) → **Personal access tokens** → **Fine-grained tokens** →
   **Generate new token**.
2. **Token name:** `Envio planilha Tolva`.
   **Expiration:** a maior opção disponível (por exemplo, 1 ano).
3. **Repository access:** *Only select repositories* → escolha `Gestao-tolva`.
4. **Permissions → Repository permissions → Contents:** *Read and write*.
5. Clique em **Generate token** e **copie** o texto que aparece (começa com
   `github_pat_`). Ele só é mostrado uma vez.

### 2. Instalar

1. Baixe a pasta `envio-automatico/` para o computador (qualquer lugar).
2. Dê dois cliques em **`instalar.bat`**.
3. Escolha a planilha `GESTÃO TOLVA.xlsm` na janela que abrir.
4. Cole o token (o texto não aparece enquanto você cola) e aperte **Enter**.
5. O instalador agenda a verificação a cada 5 minutos e já faz um envio de teste.

O token fica guardado criptografado com a sua conta do Windows: só o seu usuário,
nesse computador, consegue usá-lo. Ele não é salvo no GitHub nem na planilha.

### Quando algo der errado

- **Aviso amarelo no canto da tela** "O GitHub recusou o token": o token venceu.
  Gere outro (passo 1) e rode o `instalar.bat` de novo.
- **"Não encontrei a planilha"**: a planilha foi movida ou renomeada. Rode o
  `instalar.bat` de novo e escolha o arquivo no lugar novo.
- **Remover do computador:** dois cliques em `desinstalar.bat`.

Se a empresa bloquear tarefas agendadas ou scripts, dá para continuar subindo a
planilha à mão em `uploads/` (com o nome `GESTAO_TOLVA.xlsm`, substituindo o
arquivo que já está lá).

## Estrutura

```
index.html                     -> a página do painel
scripts/parse_excel.py         -> lê a planilha e gera o JSON
uploads/GESTAO_TOLVA.xlsm      -> a planilha (enviada pelo envio automático)
data/latest.json               -> dados usados pela página (gerado automaticamente)
data/historico.json            -> trocas registradas, para o gráfico por mês
envio-automatico/              -> programa que envia a planilha do computador
.github/workflows/process.yml  -> reprocessa sempre que a planilha chega em uploads/
```
