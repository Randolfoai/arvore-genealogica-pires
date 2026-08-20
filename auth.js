// Sistema de autenticação para Família Pires
// Senhas armazenadas com hash simples (segurança de sessão, não criptografia)

const AUTH = {
  // Hash simples para demonstração
  hash: (str) => {
    let h = 0;
    for (let i = 0; i < str.length; i++) {
      h = ((h << 5) - h) + str.charCodeAt(i);
      h = h & h;
    }
    return h.toString();
  },

  // Usuários pré-definidos
  users: {
    'user': { hash: null, nome: 'Randolfo', predefined: true },
    'pires': { hash: null, nome: 'Autor', predefined: false },
    'fabia': { hash: null, nome: 'Fábia - Agência', predefined: false }
  },

  // Validar login
  validate: function(username, password) {
    if (!this.users[username]) return false;
    if (this.users[username].hash === null) return false; // Não criada ainda
    return this.users[username].hash === this.hash(password);
  },

  // Criar nova senha (primeira vez)
  create: function(username, password) {
    if (!this.users[username]) return false;
    if (this.users[username].hash !== null) return false; // Já existe
    this.users[username].hash = this.hash(password);
    this.saveToLocalStorage();
    return true;
  },

  // Persistência
  saveToLocalStorage: function() {
    localStorage.setItem('pires_auth', JSON.stringify(this.users));
  },

  loadFromLocalStorage: function() {
    const saved = localStorage.getItem('pires_auth');
    if (saved) {
      this.users = JSON.parse(saved);
    }
  },

  // Sessão do navegador
  setSession: function(username) {
    sessionStorage.setItem('pires_user', username);
  },

  getSession: function() {
    return sessionStorage.getItem('pires_user');
  },

  logout: function() {
    sessionStorage.removeItem('pires_user');
  },

  isAuthenticated: function() {
    return !!this.getSession();
  }
};

// Inicializar
AUTH.loadFromLocalStorage();
