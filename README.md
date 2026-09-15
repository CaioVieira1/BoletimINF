# Gerador de boletins — guia de instalação e uso

Arquivos deste pacote:

| Arquivo | O que é |
|---|---|
| `app.py` | o aplicativo Streamlit (é o único arquivo de código) |
| `requirements.txt` | a lista de bibliotecas que precisam ser instaladas |
| `Boletim_TURMA_01/02/03.pptx` | exemplos já gerados com a sua planilha e o seu modelo |

---

## 1. Instalar o Python

1. Baixe em <https://www.python.org/downloads/> (versão 3.10 ou mais nova).
2. Na primeira tela do instalador, **marque "Add python.exe to PATH"** antes de clicar em Install.
3. Para conferir, abra o Prompt de Comando (tecle `Win`, digite `cmd`) e rode:

```bat
python --version
```

## 2. Montar a pasta do projeto

Crie uma pasta, por exemplo `C:\boletins`, e coloque dentro dela o `app.py` e o `requirements.txt`.
No Prompt de Comando, entre nessa pasta:

```bat
cd C:\boletins
```

## 3. Criar um ambiente isolado e instalar as bibliotecas

```bat
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

Depois de ativar o ambiente, a linha do prompt começa com `(.venv)`.
No Mac ou Linux, o comando de ativação é `source .venv/bin/activate`.

Isso só precisa ser feito uma vez. Nas próximas vezes, basta `cd C:\boletins` e `.venv\Scripts\activate`.

## 4. Rodar

```bat
streamlit run app.py
```

O navegador abre sozinho em `http://localhost:8501`. Para desligar, volte ao Prompt e tecle `Ctrl + C`.

## 5. Usar

1. Digite a senha (**030826**) e clique em Entrar.
2. Envie o modelo `.pptx` e a planilha `.xlsx`.
3. Confira a prévia: ela mostra quantos alunos (slides) cada turma vai gerar e quantas mensagens existem em cada aba Grade.
4. Clique em **Gerar PowerPoints**.
5. Baixe tudo de uma vez no `.zip` ou um arquivo por turma.

---

## Como o app lê a planilha

- Processa **toda aba cujo nome começa com "TURMA"** (TURMA 01, TURMA 02, TURMA 04...). Não precisa mexer no código quando você criar uma turma nova.
- **A linha 1 é o cabeçalho e é ignorada.**
- Uma linha só vira slide se a **coluna A (Student)** estiver preenchida — linhas em branco no fim da planilha são descartadas sozinhas.

| Marcador no PowerPoint | Coluna |
|---|---|
| `{{NOME}}` | A — Student |
| `{{LIS}}` | B — Listening |
| `{{SP}}` | C — Speaking |
| `{{WR}}` | D — Writing |
| `{{PREP}}` | E — Class preparation |
| `{{CONS}}` | F — Consolidation exercises |
| `{{ABSENCES}}` | H — Abscences |
| `{{BOOK}}` | I — Book |
| `{{MIDTERM TEST}}` | J — Midterm test |
| `{{CLASS}}` | K — Class |
| `{{DATE}}` | L — Date |
| `{{MESSAGE}}` | sorteado conforme a coluna G — Overall |

Se um dia você reordenar as colunas de novo, é só ajustar o bloco `MAPA_COLUNAS`, no topo do `app.py` — é uma linha por marcador.

Os valores saem **exatamente como aparecem no Excel**: `4` na coluna E vira `4/7`, `6` na coluna F vira `6/6`, e as datas seguem o formato da célula (`18/09/2026` na coluna Date, `02/11` no Midterm test, porque essa célula está formatada como `dd/mm`). Se quiser o Midterm test como `8/10`, é só corrigir a formatação dessa célula na planilha — o app acompanha.

### A mensagem

- Coluna G = `A` → sorteia uma frase da coluna A da aba **Grade A**; `B` → **Grade B**; e assim por diante.
- O sorteio é **sem repetição**: dentro da mesma geração, o app só repete uma frase depois de usar todas as 10 daquele grupo. Como cada turma tem poucos alunos, na prática ninguém recebe a mesma frase do colega.
- Em **Opções** existe a caixa "Fixar o sorteio". Marcada, a mesma planilha gera sempre as mesmas frases (útil se você precisar regerar um arquivo idêntico). Desmarcada, cada clique sorteia de novo.

### Se a planilha vier sem os valores das fórmulas

As colunas G, I, J, K e L são fórmulas. Quando o arquivo é salvo pelo Excel, o valor calculado vai junto e o app lê normalmente. Se por algum motivo vier vazio, o app se vira: repete o último valor preenchido acima (Book, Midterm, Class, Date) e recalcula o conceito Overall a partir de Listening/Speaking/Writing (A=10, B=7,5, C=5, D=2,5; ≥85% = A, ≥70% = B, ≥50% = C, abaixo = D). Quando isso acontece, aparece um aviso amarelo na tela.

---

## Trocar a senha

No topo do `app.py`:

```python
SENHA_PADRAO = "030826"
```

Se publicar o app na internet, é melhor não deixar a senha no código — veja a seção seguinte.

## Publicar online (opcional, gratuito)

Assim dá para usar do celular ou de outro computador, sem instalar nada.

1. Crie uma conta no GitHub e um repositório (pode ser privado) com o `app.py` e o `requirements.txt`.
2. Entre em <https://share.streamlit.io> com essa conta do GitHub e clique em **New app**.
3. Aponte para o repositório e para o arquivo `app.py`.
4. Em **Advanced settings → Secrets**, cole:

   ```toml
   APP_PASSWORD = "030826"
   ```

   O app usa esse valor no lugar da senha do código.
5. Clique em Deploy. Em um ou dois minutos você recebe um link `.streamlit.app`.

## Problemas comuns

| Sintoma | Solução |
|---|---|
| `'streamlit' não é reconhecido` | o ambiente não está ativo: rode `.venv\Scripts\activate` antes |
| `'python' não é reconhecido` | reinstale o Python marcando "Add python.exe to PATH" |
| Aviso "Marcadores previstos que não estão no modelo" | você apagou (ou escreveu diferente) algum `{{...}}` no PowerPoint |
| O marcador aparece no slide como `{{NOME}}` | o texto foi digitado com nome diferente do previsto; confira a grafia |
| "Nenhuma aba começando com TURMA" | renomeie a aba para começar com a palavra TURMA |
| Mudou o `app.py` e nada mudou na tela | clique no menu ⋮ do Streamlit → Rerun, ou recarregue a página |
