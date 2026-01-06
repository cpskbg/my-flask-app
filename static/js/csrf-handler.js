/**
 * CSRF Token Handler
 * Gère automatiquement les tokens CSRF pour toutes les requêtes AJAX
 */

// Fonction pour récupérer le token CSRF depuis le meta tag
function getCSRFToken() {
    const metaTag = document.querySelector('meta[name="csrf-token"]');
    if (metaTag) {
        return metaTag.getAttribute('content');
    }
    
    // Fallback: essayer de récupérer depuis un input hidden
    const inputTag = document.querySelector('input[name="csrf_token"]');
    if (inputTag) {
        return inputTag.value;
    }
    
    console.warn('⚠️ Token CSRF non trouvé. Assurez-vous que le meta tag csrf-token existe.');
    return null;
}

// Configuration globale pour fetch() - Ajouter automatiquement le token CSRF
const originalFetch = window.fetch;
window.fetch = function(url, options = {}) {
    // Ne pas modifier les requêtes GET
    if (!options.method || options.method.toUpperCase() === 'GET') {
        return originalFetch(url, options);
    }
    
    // Ajouter le token CSRF aux headers
    const csrfToken = getCSRFToken();
    if (csrfToken) {
        options.headers = options.headers || {};
        
        // Supporter différents formats de headers
        if (options.headers instanceof Headers) {
            options.headers.set('X-CSRFToken', csrfToken);
        } else {
            options.headers['X-CSRFToken'] = csrfToken;
        }
    }
    
    return originalFetch(url, options);
};

// Configuration globale pour XMLHttpRequest
(function() {
    const originalOpen = XMLHttpRequest.prototype.open;
    const originalSend = XMLHttpRequest.prototype.send;
    
    XMLHttpRequest.prototype.open = function(method, url, async, user, password) {
        this._method = method;
        return originalOpen.apply(this, arguments);
    };
    
    XMLHttpRequest.prototype.send = function(data) {
        // Ajouter le token CSRF pour les requêtes non-GET
        if (this._method && this._method.toUpperCase() !== 'GET') {
            const csrfToken = getCSRFToken();
            if (csrfToken) {
                this.setRequestHeader('X-CSRFToken', csrfToken);
            }
        }
        return originalSend.apply(this, arguments);
    };
})();

// Configuration globale pour jQuery AJAX (si jQuery est utilisé)
if (typeof jQuery !== 'undefined') {
    jQuery.ajaxSetup({
        beforeSend: function(xhr, settings) {
            // Ne pas ajouter le token pour les requêtes GET
            if (settings.type !== 'GET') {
                const csrfToken = getCSRFToken();
                if (csrfToken) {
                    xhr.setRequestHeader('X-CSRFToken', csrfToken);
                }
            }
        }
    });
}

// Fonction utilitaire pour créer des requêtes POST avec CSRF
function postWithCSRF(url, data = {}, options = {}) {
    const csrfToken = getCSRFToken();
    
    return fetch(url, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
            'X-CSRFToken': csrfToken,
            ...options.headers
        },
        body: JSON.stringify(data),
        ...options
    });
}

// Fonction utilitaire pour créer des formulaires avec CSRF
function submitFormWithCSRF(formElement) {
    const csrfToken = getCSRFToken();
    
    // Vérifier si le formulaire a déjà un input CSRF
    let csrfInput = formElement.querySelector('input[name="csrf_token"]');
    
    if (!csrfInput) {
        // Créer un input hidden pour le token CSRF
        csrfInput = document.createElement('input');
        csrfInput.type = 'hidden';
        csrfInput.name = 'csrf_token';
        formElement.appendChild(csrfInput);
    }
    
    // Mettre à jour la valeur du token
    csrfInput.value = csrfToken;
    
    // Soumettre le formulaire
    formElement.submit();
}

// Gestionnaire d'erreur CSRF
function handleCSRFError(error) {
    console.error('❌ Erreur CSRF:', error);
    
    // Afficher un message à l'utilisateur
    if (typeof Swal !== 'undefined') {
        Swal.fire({
            title: 'Session Expirée',
            text: 'Votre session a expiré. Veuillez rafraîchir la page.',
            icon: 'warning',
            confirmButtonText: 'Rafraîchir',
            confirmButtonColor: '#f97316'
        }).then((result) => {
            if (result.isConfirmed) {
                window.location.reload();
            }
        });
    } else {
        alert('Session expirée. Veuillez rafraîchir la page.');
        window.location.reload();
    }
}

// Exporter les fonctions pour utilisation globale
window.CSRFHandler = {
    getToken: getCSRFToken,
    postWithCSRF: postWithCSRF,
    submitFormWithCSRF: submitFormWithCSRF,
    handleError: handleCSRFError
};

// Log de confirmation
console.log('✅ CSRF Handler initialisé avec succès');
