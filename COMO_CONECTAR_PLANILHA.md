# 🍫 ChocoNota — Guia Rápido

O **ChocoNota** conta com **Tela de Login e Criptografia SHA-256**, protegido para hospedagem no **GitHub Pages**.

---

## 🔐 Segurança no GitHub

- Nem a senha nem a URL da planilha Google ficam expostas em texto puro no código ([`app.js`](./app.js)).
- A URL da planilha está **criptografada** dentro do código e só é descriptografada em memória após a autenticação com usuário e senha válidos.

### Como trocar sua senha no celular:
1. Após entrar no sistema, toque no ícone **⚙️ (Configurações)** no canto superior direito.
2. Na seção **🔐 Alterar Login e Senha**, digite seu novo usuário e senha e toque em **Atualizar Login e Senha**.
3. Para trancar o aplicativo a qualquer momento, toque no botão **🔒 (Bloquear/Sair)** no topo da tela.
