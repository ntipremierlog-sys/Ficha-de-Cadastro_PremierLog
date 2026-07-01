# Ficha de Cadastro de Admissão — Premier Logistics

Este é um projeto web interativo e responsivo para o preenchimento da **Ficha de Cadastro de Admissão** da **Premier Logistics Gestão Empresarial Ltda**.

O sistema foi desenvolvido como uma página única (SPA) estática, o que permite o seu funcionamento completo diretamente no navegador sem a necessidade de servidores ou instalações complexas.

## 🚀 Como Executar o Projeto

1. Faça o download ou clone este repositório.
2. Dê um duplo clique no arquivo [`index.html`](file:///c:/Users/financeiro3/OneDrive%20-%20Premier%20Logistics/%C3%81rea%20de%20Trabalho/NTI/Ficha%20-%20RH/index.html) para abri-lo em qualquer navegador web moderno (Google Chrome, Microsoft Edge, Mozilla Firefox, Safari, etc.).
3. Preencha os dados e gere o PDF assinado.

---

## ✨ Funcionalidades Principais

* **Identidade Visual Premium:** Estilo moderno com as cores corporativas da Premier Logistics (Azul Escuro, Laranja e Ouro) e tipografia elegante (Garamond).
* **Validação Robusta de Formulário:** 
  - Verificação de preenchimento obrigatório para dados sensíveis, incluindo **Endereço Residencial**, **WhatsApp**, **E-mail**, **Grau de Instrução** e **Tamanhos de Fardamento (Bota, Camisa, Calça)**.
  - Máscaras automáticas e validação de formato para **CPF**, **RG**, **WhatsApp** e **E-mail**.
* **Assinatura Digital & Eletrônica:**
  - **Opção 1 (Digital/Texto):** Digitação do nome completo com visualização instantânea de um exemplo de assinatura manuscrita estilizada em Garamond.
  - **Opção 2 (Manual/Desenho):** Assinatura manual utilizando o mouse (computador) ou a tela sensível ao toque (mobile/tablet).
* **Geração de PDF Otimizada:**
  - Exportação direta dos dados preenchidos no formato oficial da ficha admissionária.
  - Inclusão e posicionamento corretos dos logotipos da **Premier Logistics** e da **SulAmérica Saúde** ao lado dos respectivos blocos no documento final.
  - Quebras de página e layout ajustados para evitar sobreposição de textos legais.

---

## 🛠️ Tecnologias Utilizadas

* **React (via CDN):** Para a estruturação e controle de estado dos componentes de forma dinâmica.
* **Tailwind CSS (via CDN):** Framework CSS para estilização moderna e responsividade (Mobile/Desktop).
* **html2pdf.js:** Biblioteca JavaScript para conversão e renderização do formulário em formato PDF de alta qualidade.
* **Garamond (Google Fonts):** Fonte tipográfica padrão para uma apresentação clássica e profissional do documento.

---

## 📁 Estrutura do Projeto

* `index.html`: Arquivo principal contendo a interface, lógica do React e estilos.
* `Logo_Sul_America_2015.jpg` / `sulamerica.png`: Logotipos da SulAmérica Saúde.
* `Premier - Marca-*.png/jpg`: Versões do logotipo da Premier Logistics.
* `.gitignore`: Arquivo para ignorar arquivos temporários e de sistema no controle de versão.
