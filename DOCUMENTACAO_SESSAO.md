# Documentação de Sessão — Sistema de Recrutamento PRF
**Data:** 21/04/2026  
**Arquivos envolvidos:** `Login.html` (640 linhas), `Código.gs.js` (2.406 linhas)

---

## 1. Correção de Bug Crítico no Login 2FA

### Problema
Após inserir o código 2FA na Etapa 2, o sistema retornava para a tela inicial de login (Etapa 1 — e-mail) em vez de avançar para o formulário ou para a seleção de destino.

### Causa Raiz
Três bugs independentes no `withSuccessHandler` da função `verificarCodigo()`:

| # | Local | Bug |
|---|-------|-----|
| 1 | `verificarCodigo()` — success handler | `res` poderia ser `null` ou `undefined`; acessar `res.formUrl` lançava `TypeError` silencioso |
| 2 | `verificarCodigo()` — linha `else` | `window.top.location.href = res.formUrl` com `formUrl` indefinido recarregava a página de login |
| 3 | Botão `btnProsseguir` | Em caso de erro silencioso (TypeError no handler), o botão ficava travado em "Verificando..." |

### Correções Aplicadas em `Login.html`

**Guard contra `res` nulo** — antes de acessar propriedades:
```js
if (!res || typeof res !== 'object') {
  _resetBtnP();
  elCodigo.classList.add('invalid');
  elErro.textContent = 'Código incorreto. Verifique e tente novamente.';
  elErro.classList.add('visivel');
  return;
}
```

**`else if` com verificação de URL** — substituiu o `else` cego:
```js
// Antes (bug):
} else {
  window.top.location.href = res.formUrl; // poderia ser undefined
}

// Depois (correto):
} else if (res.formUrl) {
  window.top.location.href = res.formUrl;
} else {
  _resetBtnP();
  mostrarToast('Erro inesperado na resposta do servidor. Tente novamente.', 'erro');
}
```

**Função `_resetBtnP()` extraída** — garante reset em todos os caminhos de erro:
```js
function _resetBtnP() {
  btnP.disabled  = false;
  btnP.innerHTML = btnLabel;
}
```

---

## 2. Nova Funcionalidade — Código 2FA Existente na Etapa 1

### Comportamento Implementado
Ao digitar um e-mail `@prf.gov.br` válido, a Etapa 1 passa a exibir **dois caminhos simultâneos**:

```
[E-mail preenchido]
 
┌──────────────────────────────────┐
│  Já tenho um código              │
│  [_______ 000000 _______]        │
│  [       Prosseguir            ] │
└──────────────────────────────────┘
           ── ou ──
[   Solicitar novo código 2FA    ]
```

**Fluxo do caminho "Já tenho um código":**
1. Chama `verificarCodigo2FA(email, codigo)` no servidor
2. **Código válido** → segue o fluxo normal (destino ou formulário)
3. **Código inválido** → gera e envia novo código automaticamente → abre a Etapa 2 com a mesma mensagem padrão (transparente para o usuário)

### Mudanças no HTML (`Login.html`)

**CSS adicionado:**
```css
/* Divisor visual entre os dois caminhos */
.divider-ou {
  display:flex; align-items:center; gap:10px;
  margin:14px 0; color:#b0b8c4; font-size:11px;
}
.divider-ou::before, .divider-ou::after { content:''; flex:1; height:1px; background:#dde3ea; }

/* Caixa com borda para o campo de código existente */
.box-codigo-existente {
  border:1.5px solid #dde3ea; border-radius:8px;
  padding:14px 14px 10px; margin-bottom:14px; background:#f8fafc;
}
```

**HTML adicionado** (dentro de `#etapaEmail`, acima do botão):
```html
<div id="secaoCodigoExistente" style="display:none;">
  <div class="box-codigo-existente">
    <label class="field-label" for="codigoExistente">Já tenho um código <span class="obrig">*</span></label>
    <input type="text" id="codigoExistente" placeholder="000000" maxlength="6"
           inputmode="numeric" class="code-input"
           oninput="this.value=this.value.replace(/\D/,'')"
           onkeydown="if(event.key==='Enter') tentarCodigoExistente()">
    <div class="msg-erro" id="erroCodigoExistente"></div>
    <button class="btn btn-primary" id="btnUsarCodigo" onclick="tentarCodigoExistente()">
      Prosseguir
    </button>
  </div>
  <div class="divider-ou"><span>ou</span></div>
</div>
```

### Mudanças no JavaScript (`Login.html`)

**`validarEmail()` atualizada** — mostra/esconde `#secaoCodigoExistente` conforme e-mail válido:
```js
secao.style.display = email.endsWith('@prf.gov.br') ? '' : 'none';
```
Também limpa o campo e erros quando o e-mail deixa de ser válido.

**Helpers extraídos** (compartilhados pelos dois caminhos de verificação):

| Função | Responsabilidade |
|--------|-----------------|
| `_mostrarEtapaCodigo(email)` | Transição para Etapa 2 após envio bem-sucedido de código |
| `_processarRespostaVerificacao(res, onErro)` | Processa o objeto retornado por `verificarCodigo2FA` — avança para destino, redireciona ou chama `onErro` |

**`tentarCodigoExistente()` — nova função:**
```
1. Valida 6 dígitos
2. Chama verificarCodigo2FA(email, codigo) no servidor
3a. Sucesso com URLs → _processarRespostaVerificacao()
3b. Sucesso sem URLs (código inválido) → enviarNovoAposInvalido()
3c. Falha (throw do servidor) → enviarNovoAposInvalido()

enviarNovoAposInvalido():
  → Chama enviarEmail2FA(email)
  → Em sucesso → _mostrarEtapaCodigo(email)
```

**`voltarEmail()` atualizada** — limpa campo de código existente e restaura visibilidade da seção corretamente.

**`enviarCodigo2FA()` refatorada** — usa `_mostrarEtapaCodigo()` internamente; rótulo do botão atualizado para "Solicitar novo código 2FA".

---

## 3. Análise de Bugs no Carregamento de Respostas (`Código.gs.js`)

### Contexto
A aba "respostas" da planilha possui **39 colunas** (índices 0–38):

```
 0  Data/Hora            13  Doutorados           26  Extroversão
 1  E-mail               14  Currículo SouGov     27  Amabilidade
 2  Nome                 15  Cônjuge Nome         28  Conscienciosidade
 3  Matrícula            16  Cônjuge Matrícula    29  Estab. Emocional
 4  Cargo                17  Cônjuge E-mail       30  Abertura
 5  Unidade Oportunidade 18  Tipo de União        31  Assinatura  ← JSON base64
 6  Conhecimento Unidade 19  Data da União        32  Status
 7  DDD                  20  Endereço Cônjuge 1   33  ID Confirmação
 8  Telefone             21  Endereço Cônjuge 2   34  Status Cônjuge
 9  Cônjuge              22  Lotação Cônjuge 1    35  Data Confirmação Cônjuge
10  Graduações           23  Lotação Cônjuge 2    36  PDF Respostas
11  Pós-Graduações       24  URL Comprov. União   37  PDF Termo
12  Mestrados            25  URL Comprob. Coab.   38  PDF Concordância Cônjuge
```

Todos os mapeamentos de índice no código (`verificarInscricaoExistente`, `salvarProgressoInscricao`, `processarInscricao`, `obterCandidato`, `listarCandidatos`, `_processarRespostaConjuge`) estão **corretos e alinhados** com `HEADERS_RESPOSTAS`.

### Bug 1 — `verificarInscricaoExistente` não retorna o `status` da inscrição

**Local:** `Código.gs.js:225`  
**Impacto:** Alto — re-submissões indevidas possíveis

```js
// Lê apenas 27 colunas (índices 0-26), parando em Extroversão
const data = sheet.getRange(1, 1, lastRow, 27).getValues();
```

O campo `Status` está no índice 32, fora do range lido. O formulário **não sabe** se o candidato já finalizou ("Inscrito") ou está em rascunho ("Em preenchimento"), o que pode permitir reenvios acidentais.

**Correção sugerida:**
```js
// Ler até índice 32 (Status), ainda evitando Assinatura (31) e demais
const data = sheet.getRange(1, 1, lastRow, 34).getValues();
// E adicionar ao objeto retornado:
status: String(row[32] || ''),
```

O comentário na linha também está incorreto: diz "colunas 1–26" mas o código lê 27 colunas (0–26). A justificativa de evitar a Assinatura (coluna 32) é válida, mas o range lido é mais restrito do que o necessário.

---

### Bug 2 — `listarCandidatos` e `inicializarPainel` leem todas as colunas incluindo Assinatura

**Local:** `Código.gs.js:1864`, `1929`, `1966`  
**Impacto:** Alto — risco de timeout (limite GAS: 6 min) e lentidão no Painel

```js
const allRows = sheet.getDataRange().getValues().slice(1);
```

`getDataRange()` lê **todas as 39 colunas**, incluindo a coluna `Assinatura` (índice 31) que contém um JSON base64 do desenho da assinatura digital — podendo ter **dezenas de KB por candidato**.

`verificarInscricaoExistente` já tinha essa otimização (lê só 27 colunas), mas as funções do Painel não aplicam o mesmo cuidado.

**Correção sugerida para `listarCandidatos`** (usa no máximo índice 34):
```js
const lastRow = sheet.getLastRow();
const allRows = sheet.getRange(1, 1, lastRow, 35).getValues().slice(1);
```

**Correção sugerida para `obterCandidato`** (precisa de todas exceto Assinatura):
```js
// Duas leituras para pular a coluna 32 (Assinatura, 1-indexed)
const p1 = sheet.getRange(1, 1,  lastRow, 31).getValues().slice(1); // cols 0-30
const p2 = sheet.getRange(1, 33, lastRow, 7).getValues().slice(1);  // cols 32-38
const allRows = p1.map(function(r, i) { return r.concat([''], p2[i] || []); });
```

---

### Bug 3 — `salvarProgressoInscricao` não atualiza o cabeçalho quando a planilha já existe

**Local:** `Código.gs.js:420`  
**Impacto:** Médio — cabeçalho da planilha pode ficar desatualizado

`processarInscricao` (envio final) tem o bloco que atualiza o cabeçalho:
```js
if (sheetResp.getLastColumn() < HEADERS_RESPOSTAS.length) {
  cabecalho.setValues([HEADERS_RESPOSTAS]);
}
```

`salvarProgressoInscricao` (salvamentos parciais a cada avanço de tela) **não tem esse bloco**, apenas cria o cabeçalho na primeira criação da aba. Se `HEADERS_RESPOSTAS` cresceu após a criação da planilha, as novas colunas ficam sem rótulo no Sheets.

**Correção sugerida:**
```js
let sheetResp = ss.getSheetByName('respostas');
if (!sheetResp) {
  sheetResp = ss.insertSheet('respostas');
  sheetResp.appendRow(HEADERS_RESPOSTAS);
  sheetResp.setFrozenRows(1);
} else if (sheetResp.getLastColumn() < HEADERS_RESPOSTAS.length) {
  sheetResp.getRange(1, 1, 1, HEADERS_RESPOSTAS.length).setValues([HEADERS_RESPOSTAS]);
}
```

---

### Bug 4 — Campo `conjuge` retorna `'nao'` para rascunhos sem resposta

**Local:** `Código.gs.js:263`  
**Impacto:** Baixo — pré-seleciona "Não" mesmo sem decisão do candidato

```js
// Retorna 'nao' quando o campo está vazio (rascunho incompleto)
conjuge: String(row[9] || '').toLowerCase() === 'sim' ? 'sim' : 'nao',
```

Se o candidato salvou um rascunho antes de responder a pergunta sobre cônjuge, o campo fica vazio no Sheets. Ao recarregar o formulário, o sistema pré-seleciona "Não" e **esconde o bloco de campos de cônjuge**, que o candidato ainda precisaria ver.

**Correção sugerida:**
```js
conjuge: (function(v) {
  v = String(v || '').toLowerCase();
  return v === 'sim' ? 'sim' : (v === 'não' || v === 'nao' ? 'nao' : null);
})(row[9]),
```
O `preencherFormulario` no `Index.html` já trata `null` corretamente (não seleciona nada).

---

## 4. Resumo das Pendências (`Código.gs.js`)

As correções abaixo **não foram aplicadas** nesta sessão — apenas identificadas. Requerem edição no `Código.gs.js` e reimplantação do script:

| Prioridade | Bug | Arquivo:Linha | Correção |
|------------|-----|---------------|----------|
| 🔴 Alta | `listarCandidatos`/`inicializarPainel` leem coluna Assinatura | `:1864, :1929` | Limitar `getRange` às colunas necessárias |
| 🔴 Alta | `verificarInscricaoExistente` não retorna `status` | `:225` | Ler até coluna 34; adicionar `status` ao retorno |
| 🟡 Média | `salvarProgressoInscricao` não atualiza cabeçalho | `:420` | Adicionar bloco `else if` de atualização |
| 🟢 Baixa | Campo `conjuge` padrão `'nao'` em rascunhos | `:263` | Retornar `null` quando campo vazio |

---

## 5. Estado Atual dos Arquivos

| Arquivo | Linhas | Status |
|---------|--------|--------|
| `Login.html` | 640 | ✅ Alterado e salvo (bugs corrigidos + nova funcionalidade) |
| `Código.gs.js` | 2.406 | ⚠️ Não alterado — 4 bugs identificados, pendentes de correção |
| `Index.html` | — | Não modificado |
| `Painel.html` | — | Não modificado |
