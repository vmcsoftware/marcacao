// Copie este arquivo para `config.js` e preencha suas chaves do Firebase.
// Este arquivo é apenas um exemplo e não contém segredos reais.

// Estrutura esperada pelo app:
// - `apiKey` e `databaseURL` são obrigatórios para habilitar Firebase Database.
// - Demais campos são recomendados para funcionalidades completas.
//
// Exemplo de configuração:
// window.FIREBASE_CONFIG = {
//   apiKey: "SUA_API_KEY",
//   authDomain: "seu-projeto.firebaseapp.com",
//   databaseURL: "https://seu-projeto-default-rtdb.firebaseio.com",
//   projectId: "seu-projeto",
//   storageBucket: "seu-projeto.appspot.com",
//   messagingSenderId: "000000000000",
//   appId: "1:000000000000:web:abcdef1234567890",
// };

// Dica:
// 1) Crie um arquivo `config.js` na raiz do projeto copiando este conteúdo.
// 2) Substitua os valores pelos da sua instância Firebase.
// 3) Faça commit e push para `main` para acionar o deploy via GitHub Actions.

// Placeholder para evitar erros de referência quando `config.js` não existir.
// O app detecta ausência de chaves e mostra um alerta na interface.
if (!window.FIREBASE_CONFIG) {
  window.FIREBASE_CONFIG = {};
}