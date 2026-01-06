/**
 * Système de Modals pour remplacer les messages Flash
 * Affiche des modals stylisés au lieu des messages flash traditionnels
 */

// Indiquer qu'un système global de flash modals est actif
window.__GLOBAL_FLASH_MODAL_ACTIVE__ = true;

// Fonction pour afficher un modal de succès
function showSuccessModal(message, title = '✅ Succès') {
    const modalHtml = `
        <div class="modal fade" id="flashSuccessModal" tabindex="-1" aria-hidden="true" data-bs-backdrop="static" data-bs-keyboard="false">
            <div class="modal-dialog modal-dialog-centered">
                <div class="modal-content" style="background: linear-gradient(135deg, #28a745 0%, #20c997 100%); border-radius: 15px; box-shadow: 0 10px 40px rgba(40, 167, 69, 0.4); border: none; overflow: hidden;">
                    <!-- En-tête -->
                    <div style="background: rgba(0, 0, 0, 0.15); padding: 15px 25px; border-bottom: 1px solid rgba(255, 255, 255, 0.2); display: flex; justify-content: space-between; align-items: center;">
                        <h5 style="color: #ffffff; margin: 0; font-weight: 700; font-size: 1.3rem; display: flex; align-items: center; gap: 12px;">
                            <i class="fas fa-check-circle" style="font-size: 1.5rem; animation: bounce 0.6s;"></i>
                            <span>${title}</span>
                        </h5>
                        <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal" aria-label="Close"></button>
                    </div>
                    
                    <!-- Corps -->
                    <div style="padding: 30px 25px; text-align: center;">
                        <div style="background: rgba(255, 255, 255, 0.2); padding: 20px; border-radius: 12px; backdrop-filter: blur(10px);">
                            <p style="color: #ffffff; font-size: 1.1rem; margin: 0; font-weight: 500; line-height: 1.6;">
                                ${message}
                            </p>
                        </div>
                    </div>
                    
                    <!-- Pied -->
                    <div style="background: rgba(0, 0, 0, 0.1); padding: 15px 25px; display: flex; justify-content: center;">
                        <button type="button" class="btn btn-light" data-bs-dismiss="modal" style="padding: 10px 30px; font-weight: 600; border-radius: 8px; box-shadow: 0 4px 12px rgba(0, 0, 0, 0.2);">
                            <i class="fas fa-check me-2"></i>OK
                        </button>
                    </div>
                </div>
            </div>
        </div>
    `;
    
    showModal(modalHtml, 'flashSuccessModal');
}

// Fonction pour afficher un modal d'erreur
function showErrorModal(message, title = '❌ Erreur') {
    const modalHtml = `
        <div class="modal fade" id="flashErrorModal" tabindex="-1" aria-hidden="true" data-bs-backdrop="static" data-bs-keyboard="false">
            <div class="modal-dialog modal-dialog-centered">
                <div class="modal-content" style="background: linear-gradient(135deg, #dc3545 0%, #c82333 100%); border-radius: 15px; box-shadow: 0 10px 40px rgba(220, 53, 69, 0.4); border: none; overflow: hidden;">
                    <!-- En-tête -->
                    <div style="background: rgba(0, 0, 0, 0.15); padding: 15px 25px; border-bottom: 1px solid rgba(255, 255, 255, 0.2); display: flex; justify-content: space-between; align-items: center;">
                        <h5 style="color: #ffffff; margin: 0; font-weight: 700; font-size: 1.3rem; display: flex; align-items: center; gap: 12px;">
                            <i class="fas fa-exclamation-circle" style="font-size: 1.5rem; animation: shake 0.5s;"></i>
                            <span>${title}</span>
                        </h5>
                        <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal" aria-label="Close"></button>
                    </div>
                    
                    <!-- Corps -->
                    <div style="padding: 30px 25px; text-align: center;">
                        <div style="background: rgba(255, 255, 255, 0.2); padding: 20px; border-radius: 12px; backdrop-filter: blur(10px);">
                            <p style="color: #ffffff; font-size: 1.1rem; margin: 0; font-weight: 500; line-height: 1.6;">
                                ${message}
                            </p>
                        </div>
                    </div>
                    
                    <!-- Pied -->
                    <div style="background: rgba(0, 0, 0, 0.1); padding: 15px 25px; display: flex; justify-content: center;">
                        <button type="button" class="btn btn-light" data-bs-dismiss="modal" style="padding: 10px 30px; font-weight: 600; border-radius: 8px; box-shadow: 0 4px 12px rgba(0, 0, 0, 0.2);">
                            <i class="fas fa-times me-2"></i>Fermer
                        </button>
                    </div>
                </div>
            </div>
        </div>
    `;
    
    showModal(modalHtml, 'flashErrorModal');
}

// Fonction pour afficher un modal d'avertissement
function showWarningModal(message, title = '⚠️ Attention') {
    const modalHtml = `
        <div class="modal fade" id="flashWarningModal" tabindex="-1" aria-hidden="true" data-bs-backdrop="static" data-bs-keyboard="false">
            <div class="modal-dialog modal-dialog-centered">
                <div class="modal-content" style="background: linear-gradient(135deg, #ffc107 0%, #ff9800 100%); border-radius: 15px; box-shadow: 0 10px 40px rgba(255, 193, 7, 0.4); border: none; overflow: hidden;">
                    <!-- En-tête -->
                    <div style="background: rgba(0, 0, 0, 0.15); padding: 15px 25px; border-bottom: 1px solid rgba(255, 255, 255, 0.2); display: flex; justify-content: space-between; align-items: center;">
                        <h5 style="color: #ffffff; margin: 0; font-weight: 700; font-size: 1.3rem; display: flex; align-items: center; gap: 12px;">
                            <i class="fas fa-exclamation-triangle" style="font-size: 1.5rem; animation: pulse 1s infinite;"></i>
                            <span>${title}</span>
                        </h5>
                        <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal" aria-label="Close"></button>
                    </div>
                    
                    <!-- Corps -->
                    <div style="padding: 30px 25px; text-align: center;">
                        <div style="background: rgba(255, 255, 255, 0.2); padding: 20px; border-radius: 12px; backdrop-filter: blur(10px);">
                            <p style="color: #ffffff; font-size: 1.1rem; margin: 0; font-weight: 500; line-height: 1.6;">
                                ${message}
                            </p>
                        </div>
                    </div>
                    
                    <!-- Pied -->
                    <div style="background: rgba(0, 0, 0, 0.1); padding: 15px 25px; display: flex; justify-content: center;">
                        <button type="button" class="btn btn-light" data-bs-dismiss="modal" style="padding: 10px 30px; font-weight: 600; border-radius: 8px; box-shadow: 0 4px 12px rgba(0, 0, 0, 0.2);">
                            <i class="fas fa-check me-2"></i>Compris
                        </button>
                    </div>
                </div>
            </div>
        </div>
    `;
    
    showModal(modalHtml, 'flashWarningModal');
}

// Fonction pour afficher un modal d'information
function showInfoModal(message, title = 'ℹ️ Information') {
    const modalHtml = `
        <div class="modal fade" id="flashInfoModal" tabindex="-1" aria-hidden="true" data-bs-backdrop="static" data-bs-keyboard="false">
            <div class="modal-dialog modal-dialog-centered">
                <div class="modal-content" style="background: linear-gradient(135deg, #0d6efd 0%, #0a58ca 100%); border-radius: 15px; box-shadow: 0 10px 40px rgba(13, 110, 253, 0.4); border: none; overflow: hidden;">
                    <!-- En-tête -->
                    <div style="background: rgba(0, 0, 0, 0.15); padding: 15px 25px; border-bottom: 1px solid rgba(255, 255, 255, 0.2); display: flex; justify-content: space-between; align-items: center;">
                        <h5 style="color: #ffffff; margin: 0; font-weight: 700; font-size: 1.3rem; display: flex; align-items: center; gap: 12px;">
                            <i class="fas fa-info-circle" style="font-size: 1.5rem;"></i>
                            <span>${title}</span>
                        </h5>
                        <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal" aria-label="Close"></button>
                    </div>
                    
                    <!-- Corps -->
                    <div style="padding: 30px 25px; text-align: center;">
                        <div style="background: rgba(255, 255, 255, 0.2); padding: 20px; border-radius: 12px; backdrop-filter: blur(10px);">
                            <p style="color: #ffffff; font-size: 1.1rem; margin: 0; font-weight: 500; line-height: 1.6;">
                                ${message}
                            </p>
                        </div>
                    </div>
                    
                    <!-- Pied -->
                    <div style="background: rgba(0, 0, 0, 0.1); padding: 15px 25px; display: flex; justify-content: center;">
                        <button type="button" class="btn btn-light" data-bs-dismiss="modal" style="padding: 10px 30px; font-weight: 600; border-radius: 8px; box-shadow: 0 4px 12px rgba(0, 0, 0, 0.2);">
                            <i class="fas fa-check me-2"></i>OK
                        </button>
                    </div>
                </div>
            </div>
        </div>
    `;
    
    showModal(modalHtml, 'flashInfoModal');
}

// Fonction utilitaire pour afficher un modal
function showModal(modalHtml, modalId) {
    // Supprimer l'ancien modal s'il existe
    const existingModal = document.getElementById(modalId);
    if (existingModal) {
        const bsModal = bootstrap.Modal.getInstance(existingModal);
        if (bsModal) {
            bsModal.dispose();
        }
        existingModal.remove();
    }
    
    // Ajouter le nouveau modal au body
    document.body.insertAdjacentHTML('beforeend', modalHtml);
    
    // Afficher le modal
    const modalElement = document.getElementById(modalId);
    const modal = new bootstrap.Modal(modalElement, {
        backdrop: 'static',
        keyboard: false
    });
    modal.show();
    
    // Supprimer le modal du DOM après fermeture
    modalElement.addEventListener('hidden.bs.modal', function () {
        modalElement.remove();
    });
}

// Détection automatique des messages flash et conversion en UN SEUL modal agrégé
document.addEventListener('DOMContentLoaded', function() {
    const nodes = Array.from(document.querySelectorAll('[data-flash-message]'));
    if (!nodes.length) return;

    // Regrouper et dédupliquer par catégorie
    const buckets = { success: new Set(), error: new Set(), danger: new Set(), warning: new Set(), info: new Set() };
    nodes.forEach((el) => {
        const raw = (el.getAttribute('data-flash-message') || '').trim();
        const cat = (el.getAttribute('data-flash-category') || 'info').toLowerCase();
        const norm = raw.replace(/\s+/g, ' ').trim();
        if (cat === 'error' || cat === 'danger') buckets.error.add(norm);
        else if (cat === 'warning') buckets.warning.add(norm);
        else if (cat === 'success') buckets.success.add(norm);
        else buckets.info.add(norm);
        el.remove();
    });

    // Déterminer la catégorie dominante (priorité: error > warning > success > info)
    let dominant = 'info';
    if (buckets.error.size) dominant = 'error';
    else if (buckets.warning.size) dominant = 'warning';
    else if (buckets.success.size) dominant = 'success';

    // Construire un message unique (liste)
    const joinSet = (s) => Array.from(s.values());
    const allMessages = [
        ...joinSet(buckets.error),
        ...joinSet(buckets.warning),
        ...joinSet(buckets.success),
        ...joinSet(buckets.info)
    ];
    if (!allMessages.length) return;

    // Marquer que des messages ont été agrégés pour ce cycle
    window.__FLASH_AGGREGATED__ = true;

    const listHtml = `<ul style="text-align:left; margin:0; padding-left:1.1rem;">${allMessages.map(m => `<li style=\"margin:6px 0;\">${m}</li>`).join('')}</ul>`;

    // Afficher un seul modal selon la catégorie dominante
    if (dominant === 'error') {
        showErrorModal(listHtml, '❌ Erreur');
    } else if (dominant === 'warning') {
        showWarningModal(listHtml, '⚠️ Attention');
    } else if (dominant === 'success') {
        showSuccessModal(listHtml, '✅ Succès');
    } else {
        showInfoModal(listHtml, 'ℹ️ Information');
    }

    // Observer pour bloquer l'apparition de modals locaux success/flash si déjà agrégés
    try {
        const observer = new MutationObserver((mutations) => {
            if (!window.__FLASH_AGGREGATED__) return;
            mutations.forEach((m) => {
                // Nouveaux noeuds ajoutés (ex: custom backdrop ou modal)
                m.addedNodes && m.addedNodes.forEach((n) => {
                    if (!(n instanceof HTMLElement)) return;
                    const id = (n.id || '').toLowerCase();

                    // Ne JAMAIS bloquer les modals globaux créés par ce fichier
                    const isGlobalFlashModal = (
                        id === 'flashsuccessmodal' ||
                        id === 'flasherrormodal' ||
                        id === 'flashwarningmodal' ||
                        id === 'flashinfomodal'
                    );

                    // Annuler uniquement les anciens modals locaux de succès/flash
                    // (ex: successItemModal, successDotationModal, etc.)
                    if (!isGlobalFlashModal && n.classList.contains('modal') && (/success/.test(id) || /flash/.test(id))) {
                        n.classList.remove('show');
                        n.style.display = 'none';
                    }

                    // Normaliser backdrop Bootstrap
                    if (n.classList.contains('modal-backdrop')) {
                        n.style.zIndex = '999998';
                        n.style.backgroundColor = 'rgba(0,0,0,0.45)';
                        n.style.backdropFilter = 'blur(3px)';
                    }
                    // Normaliser les backdrops personnalisés
                    if ((n.id || '').toLowerCase() === 'custommodalbackdrop') {
                        n.style.zIndex = '999998';
                        n.style.position = 'fixed';
                        n.style.top = '0';
                        n.style.left = '0';
                        n.style.width = '100%';
                        n.style.height = '100%';
                        n.style.backgroundColor = 'rgba(0,0,0,0.45)';
                        n.style.backdropFilter = 'blur(3px)';
                    }
                });

                // Changement d'attributs (ex: ajout de class show)
                if (m.type === 'attributes' && m.target instanceof HTMLElement) {
                    const t = m.target;
                    const id2 = (t.id || '').toLowerCase();

                    const isGlobalFlashModalAttr = (
                        id2 === 'flashsuccessmodal' ||
                        id2 === 'flasherrormodal' ||
                        id2 === 'flashwarningmodal' ||
                        id2 === 'flashinfomodal'
                    );

                    if (!isGlobalFlashModalAttr && t.classList.contains('modal') && t.classList.contains('show') && (/success/.test(id2) || /flash/.test(id2))) {
                        t.classList.remove('show');
                        t.style.display = 'none';
                    }
                }
            });
        });
        observer.observe(document.body, { childList: true, subtree: true, attributes: true, attributeFilter: ['class', 'style'] });
    } catch (_) {}
});

// Intercepter l'ouverture des modals de succès locaux pour éviter le "cascade"
document.addEventListener('show.bs.modal', function(ev) {
    try {
        const id = (ev.target && ev.target.id || '').toLowerCase();

        // Ne pas bloquer les modals globaux créés par ce fichier
        const isGlobalFlashModal = (
            id === 'flashsuccessmodal' ||
            id === 'flasherrormodal' ||
            id === 'flashwarningmodal' ||
            id === 'flashinfomodal'
        );

        if (window.__FLASH_AGGREGATED__ && !isGlobalFlashModal && (/success/.test(id) || /flash/.test(id))) {
            ev.preventDefault();
        }
    } catch (_) { /* noop */ }
});

document.addEventListener('DOMContentLoaded', function() {
    try {
        var link = document.querySelector('a.nav-link[title="Stock Bas"]');
        if (link) {
            link.addEventListener('click', function(e) {
                try {
                    e.preventDefault();
                    var href = link.getAttribute('href') || link.href || '/low-stock';
                    var url = new URL(href, window.location.origin);
                    url.searchParams.set('alert', '1');
                    window.location.href = url.toString();
                } catch (_e) {}
            });
        }
    } catch (_e) {}
});

// Animations CSS
const style = document.createElement('style');
style.textContent = `
    @keyframes bounce {
        0%, 100% { transform: translateY(0); }
        50% { transform: translateY(-10px); }
    }
    
    @keyframes shake {
        0%, 100% { transform: translateX(0); }
        25% { transform: translateX(-10px); }
        75% { transform: translateX(10px); }
    }
    
    @keyframes pulse {
        0%, 100% { transform: scale(1); }
        50% { transform: scale(1.1); }
    }
`;
document.head.appendChild(style);

// Exposer les fonctions globalement
window.showSuccessModal = showSuccessModal;
window.showErrorModal = showErrorModal;
window.showWarningModal = showWarningModal;
window.showInfoModal = showInfoModal;
