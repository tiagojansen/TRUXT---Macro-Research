// Senha de acesso ao portal. Para trocar: altere o valor de PASSWORD abaixo.
const PASSWORD = "macro2026";
const SESSION_KEY = "macro_auth";

function checkAuth() {
  return sessionStorage.getItem(SESSION_KEY) === "ok";
}

function promptPassword() {
  const overlay = document.getElementById("auth-overlay");
  overlay.style.display = "flex";

  document.getElementById("auth-form").addEventListener("submit", function (e) {
    e.preventDefault();
    const input = document.getElementById("auth-input").value;
    if (input === PASSWORD) {
      sessionStorage.setItem(SESSION_KEY, "ok");
      overlay.style.display = "none";
      document.getElementById("app").style.display = "block";
    } else {
      document.getElementById("auth-error").style.display = "block";
      document.getElementById("auth-input").value = "";
      document.getElementById("auth-input").focus();
    }
  });
}

function initAuth() {
  if (checkAuth()) {
    document.getElementById("app").style.display = "block";
  } else {
    promptPassword();
  }
}
