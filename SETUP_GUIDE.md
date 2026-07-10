# Guia de Configuração — Sistema de Controle de Admissão
## Premier Logistics

Este guia detalha o passo a passo para colocar o sistema de controle de admissão no ar usando ferramentas 100% gratuitas: **Google Sheets**, **Google Apps Script** e **GitHub Pages**.

---

### Passo 1: Criar a Planilha no Google Sheets

1. Acesse o [Google Sheets](https://sheets.google.com) com a conta Gmail da empresa.
2. Crie uma nova planilha em branco.
3. Dê um nome à planilha (ex: `Premier - Controle de Admissão`).
4. Copie o **ID da Planilha** a partir da URL do seu navegador.
   * *Exemplo de URL:* `https://docs.google.com/spreadsheets/d/1A2B3C4D5E6F7G8H9I0J/edit`
   * *O ID é:* `1A2B3C4D5E6F7G8H9I0J` (guarde este ID).

---

### Passo 2: Configurar o Google Apps Script (Backend)

1. Acesse o [Google Apps Script Editor](https://script.google.com).
2. Clique em **Novo Projeto** (New Project).
3. Dê um nome ao projeto (ex: `Premier - Admissao API`).
4. Apague todo o código padrão e cole o conteúdo completo do arquivo `apps-script.js`.
5. No início do script, configure a constante `CONFIG` com o ID obtido no Passo 1:
   ```javascript
   const CONFIG = {
     SPREADSHEET_ID:  'COLE_O_ID_DA_PLANILHA_AQUI', // <- Cole aqui!
     FORM_URL:        'https://SEU-USUARIO.github.io/ficha-rh/index.html', // <- Atualize após o Passo 4
     ADMIN_SECRET:    'Premier2025AdminSecret!', // Mantenha ou mude (deve bater com o admin.html)
     EMAIL_FROM_NAME: 'RH — Premier Logistics',
     EMAIL_SUBJECT:   'Premier Logistics — Ficha de Cadastro de Admissão'
   };
   ```
6. Clique no ícone de disquete (Salvar) ou pressione `Ctrl + S`.

---

### Passo 3: Publicar o Apps Script como Web App

1. No canto superior direito da tela do Apps Script, clique em **Implantar** (Deploy) > **Nova implantação** (New deployment).
2. Clique no ícone de engrenagem ao lado de "Selecionar tipo" e escolha **App da Web** (Web app).
3. Preencha as configurações exatamente assim:
   * **Descrição**: `Versão Inicial`
   * **Executar como**: `Eu (seu-email@gmail.com)` (Execute as: Me)
   * **Quem tem acesso**: `Qualquer um` (Who has access: Anyone) — *Isso é necessário para os candidatos enviarem o formulário.*
4. Clique em **Implantar** (Deploy).
5. O Google solicitará que você conceda permissões de acesso aos dados (Sheets e Gmail).
   * Clique em **Autorizar acesso** (Authorize access).
   * Escolha sua conta Google.
   * Clique em **Avançado** (Advanced) > **Ir para Premier - Admissao API (não seguro)**.
   * Clique em **Permitir** (Allow).
6. Copie a **URL do App da Web** gerada (ex: `https://script.google.com/macros/s/AKfycb.../exec`).

---

### Passo 4: Atualizar os arquivos HTML

#### No arquivo `index.html`:
1. Abra o arquivo `index.html`.
2. Na linha 70 (no topo da tag `<script type="text/babel">`), substitua `'COLE_A_URL_DO_APPS_SCRIPT_AQUI'` pela URL do Web App gerada no Passo 3:
   ```javascript
   const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycb.../exec';
   ```

#### No arquivo `admin.html`:
1. Abra o arquivo `admin.html`.
2. Próximo à linha 760 (na tag `<script>`), substitua a URL do Apps Script e o endereço base do formulário:
   ```javascript
   const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycb.../exec';
   const FORM_BASE_URL   = 'https://SEU-USUARIO.github.io/ficha-rh/index.html'; // Altere para seu endereço do GitHub Pages
   ```

---

### Passo 5: Hospedar no GitHub Pages

Para disponibilizar o painel do RH e o formulário do candidato de qualquer lugar:

1. Crie uma conta gratuita no [GitHub](https://github.com).
2. Crie um repositório (pode ser **Privado** para manter o código seguro). Dê o nome de `ficha-rh`.
3. Suba os seguintes arquivos para a raiz do repositório:
   * `index.html`
   * `admin.html`
   * Eventuais logos/imagens de suporte (`sulamerica.png`, etc.)
4. Nas configurações do repositório (**Settings**):
   * Acesse a aba **Pages** no menu esquerdo.
   * Em "Build and deployment", selecione a branch `main` (ou `master`) e a pasta `/ (root)`.
   * Clique em **Save**.
5. Em poucos minutos, o site estará ativo:
   * **Formulário do Candidato**: `https://seu-usuario.github.io/ficha-rh/index.html`
   * **Painel do RH**: `https://seu-usuario.github.io/ficha-rh/admin.html`

> [!IMPORTANT]
> Lembre-se de atualizar o `FORM_URL` no Passo 2 (no Apps Script) e o `FORM_BASE_URL` no Passo 4 (no `admin.html`) com o link oficial final do GitHub Pages. Sempre que mudar o código do Apps Script, faça um **Novo deploy** para aplicar as atualizações.

---

### Senhas de Acesso Padrão
* **Senha do Painel do RH (admin.html)**: `Premier@2025` (editável no código do `admin.html`)
* **Secret de API (apps-script.js & admin.html)**: `Premier2025AdminSecret!` (comunicação segura entre painel e planilha)
