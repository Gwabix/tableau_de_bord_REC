// Configuration Grist - Version simplifiée
let tablesData = {
    Menus: [],
    ODJ: [],
    Agenda: []
};

// ========================================
// SÉCURITÉ - FONCTIONS UTILITAIRES
// ========================================

/**
 * Échappe le contenu HTML pour prévenir les injections XSS
 * @param {string} text - Le texte à échapper
 * @returns {string} - Le texte échappé
 */
function escapeHtml(text) {
    if (text === null || text === undefined) return '';
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

/**
 * Échappe les attributs HTML
 * @param {string} text - Le texte à échapper
 * @returns {string} - Le texte échappé
 */
function escapeHtmlAttribute(text) {
    if (text === null || text === undefined) return '';
    return String(text)
        .replaceAll('&', '&amp;')
        .replaceAll('<', '&lt;')
        .replaceAll('>', '&gt;')
        .replaceAll('"', '&quot;')
        .replaceAll("'", '&#39;');
}

/**
 * Valide et nettoie les entrées utilisateur
 * @param {string} value - La valeur à valider
 * @param {string} type - Le type de validation
 * @param {number} maxLength - Longueur maximale autorisée
 * @returns {string} - La valeur nettoyée
 */
function validateInput(value, type, maxLength = 500) {
    if (!value || typeof value !== 'string') return '';

    // Limiter la longueur
    value = value.slice(0, maxLength);

    // Nettoyer les caractères de contrôle dangereux
    value = value.replaceAll(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, '');

    return value.trim();
}

/**
 * S'assure qu'une date existe dans la table Agenda
 * Ajoute la date si elle n'existe pas déjà (sécurité OWASP)
 * @param {number|string} dateTimestamp - Le timestamp de la date à vérifier
 * @returns {Promise<boolean>} - true si la date a été ajoutée, false si elle existait déjà
 */
async function ensureAgendaDateExists(dateTimestamp) {
    try {
        // Validation stricte du timestamp
        if (dateTimestamp === null || dateTimestamp === undefined || dateTimestamp === '') {
            console.warn('ensureAgendaDateExists: timestamp invalide (null ou vide)');
            return false;
        }

        // Convertir en nombre si nécessaire
        const timestamp = typeof dateTimestamp === 'string' ? Number.parseFloat(dateTimestamp) : dateTimestamp;

        // Vérifier que c'est un nombre valide
        if (Number.isNaN(timestamp) || !Number.isFinite(timestamp)) {
            console.warn('ensureAgendaDateExists: timestamp invalide (NaN ou Infinity)', dateTimestamp);
            return false;
        }

        // Vérifier que le timestamp est dans une plage raisonnable (entre 2000 et 2100)
        const minTimestamp = 946684800; // 1er janvier 2000
        const maxTimestamp = 4102444800; // 1er janvier 2100
        if (timestamp < minTimestamp || timestamp > maxTimestamp) {
            console.warn('ensureAgendaDateExists: timestamp hors plage acceptable', timestamp);
            return false;
        }

        // Vérifier si la date existe déjà dans l'agenda
        const dateExists = tablesData.Agenda.some(entry => entry.Date === timestamp);

        if (dateExists) {
            // La date existe déjà, pas besoin de l'ajouter
            return false;
        }

        // Ajouter la date dans la table Agenda de manière sécurisée
        console.log('Ajout de la date dans l\'agenda:', new Date(timestamp * 1000).toLocaleDateString('fr-FR'));

        await grist.docApi.applyUserActions([
            ['AddRecord', 'Agenda', null, {
                Date: timestamp
            }]
        ]);

        // Recharger les données de l'agenda pour avoir la liste à jour
        const docApi = grist.docApi;
        const agendaTable = await docApi.fetchTable('Agenda');
        tablesData.Agenda = agendaTable.id.map((id, index) => ({
            id: id,
            Date: agendaTable.Date[index]
        }));

        console.log('Date ajoutée avec succès dans l\'agenda');
        return true;

    } catch (error) {
        console.error('Erreur lors de l\'ajout de la date dans l\'agenda:', error);
        // Ne pas bloquer l'opération principale en cas d'erreur
        // L'erreur est loguée pour le débogage
        return false;
    }
}

let currentDossierCount = 1;
let etatColorMap = {
    "Clôturé": "etat-cloture",
    "Avance très bien": "etat-avance-tres-bien",
    "Avance bien": "etat-avance-bien",
    "RAS": "etat-ras",
    "Des tensions": "etat-tensions",
    "Forte difficulté, blocage": "etat-blocage"
};

// Ordre de tri des états (du pire au meilleur)
let etatSortOrder = [
    "Forte difficulté, blocage",
    "Des tensions",
    "RAS",
    "Avance bien",
    "Avance très bien",
    "Clôturé"
];

// Contexte pour réouverture du formulaire après modification
let modifyContext = {
    type: null,
    value: null,
    secondValue: null
};

// ========================================
// INITIALISATION
// ========================================

function initWidget() {
    grist.ready({
        requiredAccess: 'full'
    });

    grist.onRecord(async function (record) {
        await loadAllTables();
        initializeUI();
        attachEventListeners();
    });
}

async function loadAllTables() {
    try {
        const docApi = grist.docApi;

        const [menusTable, odjTable, agendaTable] = await Promise.all([
            docApi.fetchTable('Menus'),
            docApi.fetchTable('ODJ'),
            docApi.fetchTable('Agenda')
        ]);

        tablesData.Menus = menusTable.id.map((id, index) => ({
            id: id,
            Pres_Dist_: menusTable.Pres_Dist_[index] || '',
            Personnes: menusTable.Personnes[index] || '',
            Etat: menusTable.Etat[index] || ''
        }));

        tablesData.ODJ = odjTable.id.map((id, index) => ({
            id: id,
            Date_de_la_reunion: odjTable.Date_de_la_reunion[index],
            Dossier: odjTable.Dossier[index] || '',
            Porteur_s_: odjTable.Porteur_s_[index] || [],
            Actions_a_mettre_en_uvre_etapes: odjTable.Actions_a_mettre_en_uvre_etapes[index] || '',
            Echeance: odjTable.Echeance[index],
            Etat: odjTable.Etat[index],
            Enregistrement: odjTable.Enregistrement[index]
        }));

        tablesData.Agenda = agendaTable.id.map((id, index) => ({
            id: id,
            Date: agendaTable.Date[index]
        }));

        console.log('Tables chargées:', tablesData);
    } catch (error) {
        console.error('Erreur lors du chargement des tables:', error);
        alert('Erreur lors du chargement des données. Vérifiez les noms des tables.');
    }
}

function initializeUI() {
    populatePorteurs();
    populateEtats();
    populateConsultSelectors();
    populateReunionDateSelect();
    setDefaultDate();
}

// ========================================
// POPULATION DES ÉLÉMENTS D'INTERFACE
// ========================================

function populatePorteurs() {
    const containers = document.querySelectorAll('.dossier-porteurs');
    const personnes = getUniqueValues(tablesData.Menus, 'Personnes').filter(p => p !== 'Autre');

    containers.forEach(container => {
        container.innerHTML = '';
        personnes.forEach(personne => {
            const div = document.createElement('div');
            div.className = 'multi-select-option';
            div.dataset.value = personne;
            div.textContent = personne;
            container.appendChild(div);
        });
    });
}

function populateEtats() {
    const selects = document.querySelectorAll('.dossier-etat');
    const etats = getUniqueValues(tablesData.Menus, 'Etat');

    // Ordre personnalisé pour les états (onglet Saisir uniquement)
    const ordreEtats = [
        "Clôturé",
        "Avance très bien",
        "Avance bien",
        "RAS",
        "Des tensions",
        "Forte difficulté, blocage"
    ];

    // Trier les états selon l'ordre personnalisé
    const etatsTries = etats.sort((a, b) => {
        const indexA = ordreEtats.indexOf(a);
        const indexB = ordreEtats.indexOf(b);

        if (indexA !== -1 && indexB !== -1) {
            return indexA - indexB;
        }
        if (indexA !== -1) return -1;
        if (indexB !== -1) return 1;
        return a.localeCompare(b);
    });

    selects.forEach(select => {
        select.innerHTML = '<option value="">-- Sélectionner --</option>';
        etatsTries.forEach(etat => {
            const option = document.createElement('option');
            option.value = etat;
            option.textContent = etat;
            select.appendChild(option);
        });
    });
}

function populateConsultSelectors() {
    // Dates de réunion (depuis ODJ)
    const dates = getUniqueDates(tablesData.ODJ, 'Date_de_la_reunion');
    const dateSelects = [
        document.getElementById('consult-date-select'),
        document.getElementById('modify-date-select')
    ];

    dateSelects.forEach(select => {
        if (!select) return;
        select.innerHTML = '<option value="">-- Choisir une date --</option>';
        dates.forEach(date => {
            const option = document.createElement('option');
            option.value = date;
            option.textContent = formatDate(date);
            select.appendChild(option);
        });
    });

    // Dates d'échéance (depuis ODJ)
    const echeances = getUniqueDates(tablesData.ODJ, 'Echeance');
    const echeanceSelects = [
        document.getElementById('consult-echeance-select'),
        document.getElementById('modify-echeance-select')
    ];

    echeanceSelects.forEach(select => {
        if (!select) return;
        select.innerHTML = '<option value="">-- Choisir une date --</option>';
        echeances.forEach(date => {
            const option = document.createElement('option');
            option.value = date;
            option.textContent = formatDate(date);
            select.appendChild(option);
        });
    });

    // Porteurs
    const porteurs = getUniquePorteurs();
    const porteurSelects = [
        document.getElementById('consult-porteur-select'),
        document.getElementById('modify-porteur-select')
    ];

    porteurSelects.forEach(select => {
        if (!select) return;
        select.innerHTML = '<option value="">-- Choisir un porteur --</option>';
        porteurs.forEach(porteur => {
            const option = document.createElement('option');
            option.value = porteur;
            option.textContent = porteur;
            select.appendChild(option);
        });
    });

    // Filtres d'état pour consultation par porteur
    const filterContainer = document.getElementById('filter-etat-checkboxes');
    if (filterContainer) {
        const etats = getUniqueValues(tablesData.Menus, 'Etat');

        filterContainer.innerHTML = '';
        etats.forEach(etat => {
            const label = document.createElement('label');
            label.className = 'checkbox-label';

            const checkbox = document.createElement('input');
            checkbox.type = 'checkbox';
            checkbox.name = 'filter-etat';
            checkbox.value = etat;
            checkbox.checked = true;

            const span = document.createElement('span');
            span.textContent = etat;

            label.appendChild(checkbox);
            label.appendChild(span);
            filterContainer.appendChild(label);
        });
    }
}

function getNextMeetingDate() {
    if (!tablesData.Agenda || tablesData.Agenda.length === 0) {
        return null;
    }

    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const todayTimestamp = today.getTime() / 1000;

    const futureDates = tablesData.Agenda
        .map(item => item.Date)
        .filter(date => date >= todayTimestamp)
        .sort((a, b) => a - b);

    return futureDates.length > 0 ? futureDates[0] : null;
}

function getUpcomingMeetingDates() {
    if (!tablesData.Agenda || tablesData.Agenda.length === 0) {
        return [];
    }

    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const todayTimestamp = today.getTime() / 1000;

    return tablesData.Agenda
        .map(item => item.Date)
        .filter(date => date >= todayTimestamp)
        .sort((a, b) => a - b);
}

function populateUpcomingMeetingsSelect() {
    const select = document.getElementById('saisir-reunions-suivantes');
    if (!select) return;

    const upcomingDates = getUpcomingMeetingDates();

    select.innerHTML = '<option value="">-- Sélectionner une autre date --</option>';

    upcomingDates.forEach(date => {
        const option = document.createElement('option');
        option.value = date;
        option.textContent = formatDate(date);
        select.appendChild(option);
    });
}

function setDefaultDate() {
    const dateInput = document.getElementById('saisir-date');
    if (dateInput) {
        const nextMeetingDate = getNextMeetingDate();

        if (nextMeetingDate) {
            // Convertir le timestamp en format YYYY-MM-DD
            const date = new Date(nextMeetingDate * 1000);
            const year = date.getFullYear();
            const month = String(date.getMonth() + 1).padStart(2, '0');
            const day = String(date.getDate()).padStart(2, '0');
            dateInput.value = `${year}-${month}-${day}`;
        } else {
            // Si pas de date dans l'agenda, utiliser aujourd'hui
            const today = new Date().toISOString().split('T')[0];
            dateInput.value = today;
        }
    }

    // Peupler le menu déroulant des réunions suivantes
    populateUpcomingMeetingsSelect();
}

// ========================================
// UTILITAIRES
// ========================================

function getUniqueValues(data, column) {
    const values = data
        .map(row => row[column])
        .filter(val => val && val.toString().trim() !== '');
    return [...new Set(values)].sort();
}

function sortByEtat(dossiers) {
    return dossiers.sort((a, b) => {
        const etatA = getEtatNameById(a.Etat);
        const etatB = getEtatNameById(b.Etat);

        // Vérifier si l'état est vide
        const emptyA = !etatA || etatA.trim() === '';
        const emptyB = !etatB || etatB.trim() === '';

        // Les états vides remontent en haut
        if (emptyA && !emptyB) return -1;
        if (!emptyA && emptyB) return 1;
        if (emptyA && emptyB) return (a.Dossier || '').localeCompare(b.Dossier || '');

        // Pour les états non vides, tri selon l'ordre défini
        const indexA = etatSortOrder.indexOf(etatA);
        const indexB = etatSortOrder.indexOf(etatB);

        // Si les deux états sont dans l'ordre de tri
        if (indexA !== -1 && indexB !== -1) {
            return indexA - indexB;
        }
        // Si seulement A est dans l'ordre de tri
        if (indexA !== -1) return -1;
        // Si seulement B est dans l'ordre de tri
        if (indexB !== -1) return 1;
        // Sinon tri alphabétique
        return etatA.localeCompare(etatB);
    });
}

function getLatestEntriesPerDossier(dossiers) {
    // Grouper par nom de dossier
    const dossierMap = new Map();

    dossiers.forEach(dossier => {
        const nom = dossier.Dossier;
        const existing = dossierMap.get(nom);

        if (!existing) {
            dossierMap.set(nom, dossier);
        } else {
            // Comparer les timestamps d'enregistrement et garder le plus récent
            const existingTimestamp = existing.Enregistrement || 0;
            const currentTimestamp = dossier.Enregistrement || 0;

            if (currentTimestamp > existingTimestamp) {
                dossierMap.set(nom, dossier);
            }
        }
    });

    return Array.from(dossierMap.values());
}

function getUniqueDates(data, column) {
    const dates = data
        .map(row => row[column])
        .filter(val => val);
    return [...new Set(dates)].sort((a, b) => new Date(b) - new Date(a));
}

function getUniquePorteurs() {
    const porteurs = new Set();
    tablesData.ODJ.forEach(row => {
        if (row.Porteur_s_ && Array.isArray(row.Porteur_s_)) {
            row.Porteur_s_.forEach(id => {
                const name = getPersonneNameById(id);
                if (name) porteurs.add(name);
            });
        }
    });
    return [...porteurs].sort();
}

function formatDate(dateString) {
    if (!dateString) return '';
    let date;
    if (typeof dateString === 'number') {
        date = new Date(dateString * 1000);
    } else {
        date = new Date(dateString);
    }
    return date.toLocaleDateString('fr-FR', {
        year: 'numeric',
        month: 'long',
        day: 'numeric'
    });
}

function getPersonneNameById(id) {
    const personne = tablesData.Menus.find(m => m.id === id);
    return personne ? personne.Personnes : '';
}

function getEtatNameById(id) {
    const etat = tablesData.Menus.find(m => m.id === id);
    return etat ? etat.Etat : '';
}

function getPersonneIdByName(name) {
    const personne = tablesData.Menus.find(m => m.Personnes === name);
    return personne ? personne.id : null;
}

function getEtatIdByName(name) {
    const etat = tablesData.Menus.find(m => m.Etat === name);
    return etat ? etat.id : null;
}

// Fonction pour détecter et supprimer les doublons (en ignorant Enregistrement et id)
async function removeDuplicateRecords() {
    const duplicateGroups = new Map();

    // Grouper les enregistrements par leurs attributs (sans Enregistrement et id)
    tablesData.ODJ.forEach(record => {
        // Créer une clé unique basée sur les valeurs (sans id et Enregistrement)
        const key = JSON.stringify({
            Date_de_la_reunion: record.Date_de_la_reunion,
            Dossier: record.Dossier,
            Porteur_s_: record.Porteur_s_,
            Actions_a_mettre_en_uvre_etapes: record.Actions_a_mettre_en_uvre_etapes,
            Echeance: record.Echeance,
            Etat: record.Etat
        });

        if (!duplicateGroups.has(key)) {
            duplicateGroups.set(key, []);
        }
        duplicateGroups.get(key).push(record);
    });

    // Identifier et supprimer les doublons (garder le plus récent)
    const actionsToDelete = [];
    duplicateGroups.forEach(group => {
        if (group.length > 1) {
            // Trier par Enregistrement décroissant (plus récent en premier)
            group.sort((a, b) => (b.Enregistrement || 0) - (a.Enregistrement || 0));

            // Supprimer tous sauf le premier (le plus récent)
            for (let i = 1; i < group.length; i++) {
                actionsToDelete.push(['RemoveRecord', 'ODJ', group[i].id]);
            }
        }
    });

    // Appliquer les suppressions si nécessaire
    if (actionsToDelete.length > 0) {
        await grist.docApi.applyUserActions(actionsToDelete);
        console.log(`${actionsToDelete.length} doublon(s) supprimé(s)`);
    }
}

// ========================================
// GESTION DES ÉVÉNEMENTS
// ========================================

function attachEventListeners() {
    // Onglets
    document.querySelectorAll('.tab-button').forEach(button => {
        button.addEventListener('click', switchTab);
    });

    // Multi-select
    document.querySelectorAll('.multi-select-option').forEach(option => {
        option.addEventListener('click', toggleMultiSelect);
    });

    // Dossiers
    const btnAddDossier = document.getElementById('btn-add-dossier');
    if (btnAddDossier) {
        btnAddDossier.addEventListener('click', addDossier);
    }

    // Validation saisie
    const btnValider = document.getElementById('btn-valider-saisie');
    if (btnValider) {
        btnValider.addEventListener('click', validateSaisie);
    }

    // Consultation
    document.querySelectorAll('input[name="consult-type"]').forEach(radio => {
        radio.addEventListener('change', handleConsultTypeChange);
    });

    const consultDateSelect = document.getElementById('consult-date-select');
    if (consultDateSelect) {
        consultDateSelect.addEventListener('change', consultByDate);
    }

    const consultDossierInput = document.getElementById('consult-dossier-input');
    if (consultDossierInput) {
        attachDossierAutocomplete(consultDossierInput, 'consult');
    }

    const consultPorteurSelect = document.getElementById('consult-porteur-select');
    if (consultPorteurSelect) {
        consultPorteurSelect.addEventListener('change', handlePorteurSelectChange);
    }

    const consultEcheanceSelect = document.getElementById('consult-echeance-select');
    if (consultEcheanceSelect) {
        consultEcheanceSelect.addEventListener('change', consultByEcheance);
    }

    document.querySelectorAll('input[name="sort-type"]').forEach(radio => {
        radio.addEventListener('change', consultByPorteur);
    });

    document.querySelectorAll('input[name="filter-etat"]').forEach(checkbox => {
        checkbox.addEventListener('change', consultByPorteur);
    });

    // Réunion
    const reunionDateSelect = document.getElementById('reunion-date-select');
    if (reunionDateSelect) {
        reunionDateSelect.addEventListener('change', reunionDisplayData);
        // Afficher les données pour la date par défaut
        if (reunionDateSelect.value) {
            reunionDisplayData();
        }
    }

    const btnSaveReunion = document.getElementById('btn-save-reunion');
    if (btnSaveReunion) {
        btnSaveReunion.addEventListener('click', saveReunionModifications);
    }

    const btnPrintReunion = document.getElementById('btn-print-reunion');
    if (btnPrintReunion) {
        btnPrintReunion.addEventListener('click', printReunionResults);
    }

    // Modification
    document.querySelectorAll('input[name="modify-type"]').forEach(radio => {
        radio.addEventListener('change', handleModifyTypeChange);
    });

    const modifyDateSelect = document.getElementById('modify-date-select');
    if (modifyDateSelect) {
        modifyDateSelect.addEventListener('change', modifyByDate);
    }

    let modifyDossierInput = document.getElementById('modify-dossier-input');
    if (modifyDossierInput) {
        attachDossierAutocomplete(modifyDossierInput, 'modify');
    }

    const modifyPorteurSelect = document.getElementById('modify-porteur-select');
    if (modifyPorteurSelect) {
        modifyPorteurSelect.addEventListener('change', handleModifyPorteurSelectChange);
    }

    const modifyPorteurDossierSelect = document.getElementById('modify-porteur-dossier-select');
    if (modifyPorteurDossierSelect) {
        modifyPorteurDossierSelect.addEventListener('change', modifyByPorteurDossier);
    }

    const modifyEcheanceSelect = document.getElementById('modify-echeance-select');
    if (modifyEcheanceSelect) {
        modifyEcheanceSelect.addEventListener('change', modifyByEcheance);
    }

    const btnSaveModif = document.getElementById('btn-save-modifications');
    if (btnSaveModif) {
        btnSaveModif.addEventListener('click', saveModifications);
    }

    const btnCancelModif = document.getElementById('btn-cancel-modifications');
    if (btnCancelModif) {
        btnCancelModif.addEventListener('click', cancelModifications);
    }

    // Boutons de réinitialisation
    const btnClearConsultDossier = document.getElementById('btn-clear-consult-dossier');
    if (btnClearConsultDossier) {
        btnClearConsultDossier.addEventListener('click', clearConsultDossier);
    }

    const btnClearModifyDossier = document.getElementById('btn-clear-modify-dossier');
    if (btnClearModifyDossier) {
        btnClearModifyDossier.addEventListener('click', clearModifyDossier);
    }

    // Bouton Imprimer
    const btnPrintConsult = document.getElementById('btn-print-consult');
    if (btnPrintConsult) {
        btnPrintConsult.addEventListener('click', printConsultResults);
    }

    // Autocomplete
    attachAutocompleteListeners();

    // Menu déroulant des réunions suivantes
    const reunionsSuivantesSelect = document.getElementById('saisir-reunions-suivantes');
    if (reunionsSuivantesSelect) {
        reunionsSuivantesSelect.addEventListener('change', function () {
            if (this.value) {
                const dateInput = document.getElementById('saisir-date');
                if (dateInput) {
                    const timestamp = Number.parseFloat(this.value);
                    // Validation : vérifier que c'est un nombre valide
                    if (!Number.isNaN(timestamp) && Number.isFinite(timestamp)) {
                        const date = new Date(timestamp * 1000);
                        // Vérifier que la date est valide
                        if (!Number.isNaN(date.getTime())) {
                            const year = date.getFullYear();
                            const month = String(date.getMonth() + 1).padStart(2, '0');
                            const day = String(date.getDate()).padStart(2, '0');
                            dateInput.value = `${year}-${month}-${day}`;
                        }
                    }
                }
                // Réinitialiser le menu déroulant
                this.value = '';
            }
        });
    }
}

function switchTab(event) {
    const targetTab = event.target.dataset.tab;

    document.querySelectorAll('.tab-button').forEach(btn => btn.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(content => content.classList.remove('active'));

    event.target.classList.add('active');
    const targetContent = document.getElementById(`tab-${targetTab}`);
    if (targetContent) {
        targetContent.classList.add('active');
    }

    // Vider les sélections lors du changement d'onglet
    clearAllSelections();
}

function clearAllSelections() {
    // Vider les sélecteurs de l'onglet Consulter
    const consultDateSelect = document.getElementById('consult-date-select');
    if (consultDateSelect) consultDateSelect.value = '';

    const consultDossierInput = document.getElementById('consult-dossier-input');
    if (consultDossierInput) {
        consultDossierInput.value = '';
        toggleClearButton('btn-clear-consult-dossier', '');
    }

    const consultPorteurSelect = document.getElementById('consult-porteur-select');
    if (consultPorteurSelect) consultPorteurSelect.value = '';

    const consultEcheanceSelect = document.getElementById('consult-echeance-select');
    if (consultEcheanceSelect) consultEcheanceSelect.value = '';

    // Vider les sélecteurs de l'onglet Modifier
    const modifyDateSelect = document.getElementById('modify-date-select');
    if (modifyDateSelect) modifyDateSelect.value = '';

    const modifyDossierInput = document.getElementById('modify-dossier-input');
    if (modifyDossierInput) {
        modifyDossierInput.value = '';
        toggleClearButton('btn-clear-modify-dossier', '');
    }

    const modifyPorteurSelect = document.getElementById('modify-porteur-select');
    if (modifyPorteurSelect) modifyPorteurSelect.value = '';

    const modifyEcheanceSelect = document.getElementById('modify-echeance-select');
    if (modifyEcheanceSelect) modifyEcheanceSelect.value = '';

    // Cacher les résultats
    const consultResults = document.getElementById('consult-results');
    if (consultResults) consultResults.innerHTML = '';

    const modifyResults = document.getElementById('modify-results');
    if (modifyResults) modifyResults.innerHTML = '';

    const modifyButtons = document.getElementById('modify-buttons');
    if (modifyButtons) modifyButtons.classList.add('hidden');
}

function toggleMultiSelect(event) {
    event.target.classList.toggle('selected');
    // Réorganiser les options pour mettre les sélectionnées en haut
    reorderMultiSelectOptions(event.target.parentElement);
}

function reorderMultiSelectOptions(container) {
    if (!container || !container.classList.contains('dossier-porteurs')) return;

    const options = Array.from(container.querySelectorAll('.multi-select-option'));
    const selected = options.filter(opt => opt.classList.contains('selected'));
    const notSelected = options.filter(opt => !opt.classList.contains('selected'));

    // Vider le conteneur
    container.innerHTML = '';

    // Ajouter d'abord les sélectionnés, puis les non-sélectionnés
    [...selected, ...notSelected].forEach(option => {
        container.appendChild(option);
    });
}

// ========================================
// GESTION DES DOSSIERS
// ========================================

function addDossier() {
    currentDossierCount++;
    const container = document.getElementById('saisir-dossiers');
    const newDossier = document.createElement('div');
    newDossier.className = 'dossier-block';
    newDossier.dataset.dossier = currentDossierCount;

    newDossier.innerHTML = `
        <div class="dossier-header">
            <span class="dossier-number">Dossier ${currentDossierCount}</span>
            <button type="button" class="btn-remove-dossier">Supprimer</button>
        </div>
        <div class="dossier-fields">
            <div class="form-group">
                <label class="form-label">Dossier :</label>
                <div class="autocomplete-container">
                    <input type="text" class="form-input autocomplete-input dossier-intitule" placeholder="Intitulé du dossier">
                    <div class="autocomplete-suggestions"></div>
                </div>
            </div>

            <div class="form-group">
                <label class="form-label">Porteur(s) :</label>
                <div class="multi-select dossier-porteurs">
                </div>
            </div>

            <div class="form-group">
                <label class="form-label">Actions à mettre en œuvre – étapes :</label>
                <div class="autocomplete-container">
                    <textarea class="form-textarea autocomplete-input dossier-actions" placeholder="Décrire les actions et étapes"></textarea>
                    <div class="autocomplete-suggestions"></div>
                </div>
            </div>

            <div class="form-group">
                <label class="form-label">Échéance :</label>
                <input type="date" class="form-input form-date dossier-echeance">
            </div>

            <div class="form-group">
                <label class="form-label">État :</label>
                <select class="form-select dossier-etat">
                    <option value="">-- Sélectionner --</option>
                </select>
            </div>
        </div>
    `;

    container.appendChild(newDossier);

    // Repeupler les porteurs et états
    populatePorteurs();
    populateEtats();

    // Attacher les événements
    newDossier.querySelectorAll('.multi-select-option').forEach(option => {
        option.addEventListener('click', toggleMultiSelect);
    });

    // Attacher l'événement au bouton Supprimer (sécurisé sans onclick inline)
    const removeBtn = newDossier.querySelector('.btn-remove-dossier');
    if (removeBtn) {
        removeBtn.addEventListener('click', function () {
            removeDossier(this);
        });
    }

    const intituleInput = newDossier.querySelector('.dossier-intitule');
    const actionsInput = newDossier.querySelector('.dossier-actions');

    if (intituleInput) attachAutocompleteToInput(intituleInput);
    if (actionsInput) attachAutocompleteToInput(actionsInput);
}

function removeDossier(button) {
    const dossier = button.closest('.dossier-block');
    const container = document.getElementById('saisir-dossiers');

    if (container.children.length > 1) {
        dossier.remove();
        renumberDossiers();
    }
}

function renumberDossiers() {
    const dossiers = document.querySelectorAll('#saisir-dossiers .dossier-block');
    dossiers.forEach((dossier, index) => {
        dossier.dataset.dossier = index + 1;
        const numberSpan = dossier.querySelector('.dossier-number');
        if (numberSpan) {
            numberSpan.textContent = `Dossier ${index + 1}`;
        }
    });
    currentDossierCount = dossiers.length;
}

// ========================================
// AUTOCOMPLETE
// ========================================

function attachAutocompleteListeners() {
    document.querySelectorAll('.autocomplete-input').forEach(input => {
        attachAutocompleteToInput(input);
    });
}

function attachAutocompleteToInput(input) {
    if (!input) return;

    let currentSuggestions = [];
    let highlightedIndex = -1;

    const container = input.closest('.autocomplete-container');
    if (!container) return;

    const suggestionsDiv = container.querySelector('.autocomplete-suggestions');
    if (!suggestionsDiv) return;

    input.addEventListener('input', function () {
        const value = this.value;

        if (value.length < 2) {
            suggestionsDiv.classList.remove('visible');
            return;
        }

        let searchColumn = '';
        if (input.classList.contains('dossier-intitule')) {
            searchColumn = 'Dossier';
        } else if (input.classList.contains('dossier-actions')) {
            searchColumn = 'Actions_a_mettre_en_uvre_etapes';
        } else {
            suggestionsDiv.classList.remove('visible');
            return;
        }

        currentSuggestions = fuzzySearch(value, tablesData.ODJ, searchColumn);
        displaySuggestions(suggestionsDiv, currentSuggestions);
        highlightedIndex = currentSuggestions.length > 0 ? 0 : -1;
        highlightSuggestion(suggestionsDiv, highlightedIndex);
    });

    input.addEventListener('keydown', function (event) {
        if (!suggestionsDiv.classList.contains('visible')) {
            // Ne pas bloquer Entrée pour les textarea (permet les sauts de ligne naturels)
            if (event.key === 'Enter' && input.tagName !== 'TEXTAREA') {
                event.preventDefault();
            }
            return;
        }

        if (event.key === 'ArrowDown') {
            event.preventDefault();
            highlightedIndex = Math.min(highlightedIndex + 1, currentSuggestions.length - 1);
            highlightSuggestion(suggestionsDiv, highlightedIndex);
        } else if (event.key === 'ArrowUp') {
            event.preventDefault();
            highlightedIndex = Math.max(highlightedIndex - 1, 0);
            highlightSuggestion(suggestionsDiv, highlightedIndex);
        } else if (event.key === 'Enter' && highlightedIndex >= 0) {
            event.preventDefault();
            input.value = currentSuggestions[highlightedIndex];
            suggestionsDiv.classList.remove('visible');
        } else if (event.key === 'Escape') {
            suggestionsDiv.classList.remove('visible');
        }
    });

    input.addEventListener('blur', function () {
        setTimeout(() => {
            suggestionsDiv.classList.remove('visible');
        }, 200);
    });

    suggestionsDiv.addEventListener('click', function (event) {
        if (event.target.classList.contains('autocomplete-item')) {
            input.value = event.target.textContent;
            suggestionsDiv.classList.remove('visible');
        }
    });
}

function fuzzySearch(query, data, column) {
    const lowerQuery = query.toLowerCase();
    const matches = [];

    data.forEach(row => {
        const value = row[column];
        if (!value) return;

        const lowerValue = value.toString().toLowerCase();
        if (lowerValue.includes(lowerQuery)) {
            const distance = levenshteinDistance(lowerQuery, lowerValue);
            matches.push({ text: value, distance: distance });
        }
    });

    matches.sort((a, b) => a.distance - b.distance);
    return [...new Set(matches.map(m => m.text))].slice(0, 5);
}

function levenshteinDistance(a, b) {
    const matrix = [];

    for (let i = 0; i <= b.length; i++) {
        matrix[i] = [i];
    }

    for (let j = 0; j <= a.length; j++) {
        matrix[0][j] = j;
    }

    for (let i = 1; i <= b.length; i++) {
        for (let j = 1; j <= a.length; j++) {
            if (b.charAt(i - 1) === a.charAt(j - 1)) {
                matrix[i][j] = matrix[i - 1][j - 1];
            } else {
                matrix[i][j] = Math.min(
                    matrix[i - 1][j - 1] + 1,
                    matrix[i][j - 1] + 1,
                    matrix[i - 1][j] + 1
                );
            }
        }
    }

    return matrix[b.length][a.length];
}

function displaySuggestions(container, suggestions) {
    container.innerHTML = '';

    if (suggestions.length === 0) {
        container.classList.remove('visible');
        return;
    }

    suggestions.forEach(suggestion => {
        const div = document.createElement('div');
        div.className = 'autocomplete-item';
        div.textContent = suggestion;
        container.appendChild(div);
    });

    container.classList.add('visible');
}

function highlightSuggestion(container, index) {
    const items = container.querySelectorAll('.autocomplete-item');
    items.forEach((item, i) => {
        if (i === index) {
            item.classList.add('highlighted');
        } else {
            item.classList.remove('highlighted');
        }
    });
}

// Autocomplete pour la consultation/modification par dossier
function attachDossierAutocomplete(input, mode) {
    if (!input) return;

    let currentSuggestions = [];
    let highlightedIndex = -1;

    const container = input.closest('.autocomplete-container');
    if (!container) return;

    const suggestionsDiv = container.querySelector('.autocomplete-suggestions');
    if (!suggestionsDiv) return;

    // Déterminer l'ID du bouton clear en fonction du mode
    const clearButtonId = mode === 'consult' ? 'btn-clear-consult-dossier' : 'btn-clear-modify-dossier';

    input.addEventListener('focus', function () {
        const value = this.value;

        if (value.length === 0) {
            const allDossiers = [...new Set(tablesData.ODJ.map(odj => odj.Dossier).filter(Boolean))].sort();
            currentSuggestions = allDossiers.slice(0, 10);
            displaySuggestions(suggestionsDiv, currentSuggestions);
            highlightedIndex = 0;
            highlightSuggestion(suggestionsDiv, highlightedIndex);
        } else if (value.length >= 2) {
            currentSuggestions = fuzzySearch(value, tablesData.ODJ, 'Dossier');
            displaySuggestions(suggestionsDiv, currentSuggestions);
            highlightedIndex = 0;
            highlightSuggestion(suggestionsDiv, highlightedIndex);
        }
        // Mettre à jour le bouton clear
        toggleClearButton(clearButtonId, value);
    });

    input.addEventListener('input', function () {
        const value = this.value;

        // Mettre à jour le bouton clear
        toggleClearButton(clearButtonId, value);

        if (value.length === 0) {
            const allDossiers = [...new Set(tablesData.ODJ.map(odj => odj.Dossier).filter(Boolean))].sort();
            currentSuggestions = allDossiers.slice(0, 10);
            displaySuggestions(suggestionsDiv, currentSuggestions);
            highlightedIndex = 0;
            highlightSuggestion(suggestionsDiv, highlightedIndex);
            return;
        }

        if (value.length < 2) {
            suggestionsDiv.classList.remove('visible');
            return;
        }

        currentSuggestions = fuzzySearch(value, tablesData.ODJ, 'Dossier');
        displaySuggestions(suggestionsDiv, currentSuggestions);
        highlightedIndex = currentSuggestions.length > 0 ? 0 : -1;
        highlightSuggestion(suggestionsDiv, highlightedIndex);
    });

    input.addEventListener('keydown', function (event) {
        if (!suggestionsDiv.classList.contains('visible')) return;

        if (event.key === 'ArrowDown') {
            event.preventDefault();
            highlightedIndex = Math.min(highlightedIndex + 1, currentSuggestions.length - 1);
            highlightSuggestion(suggestionsDiv, highlightedIndex);
        } else if (event.key === 'ArrowUp') {
            event.preventDefault();
            highlightedIndex = Math.max(highlightedIndex - 1, 0);
            highlightSuggestion(suggestionsDiv, highlightedIndex);
        } else if (event.key === 'Enter' && highlightedIndex >= 0) {
            event.preventDefault();
            input.value = currentSuggestions[highlightedIndex];
            suggestionsDiv.classList.remove('visible');
            // Mettre à jour le bouton clear après sélection
            toggleClearButton(clearButtonId, input.value);
            if (mode === 'consult') {
                consultByDossier(currentSuggestions[highlightedIndex]);
            } else if (mode === 'modify') {
                modifyByDossier(currentSuggestions[highlightedIndex]);
            }
        } else if (event.key === 'Escape') {
            suggestionsDiv.classList.remove('visible');
        }
    });

    input.addEventListener('blur', function () {
        setTimeout(() => {
            suggestionsDiv.classList.remove('visible');
        }, 200);
    });

    suggestionsDiv.addEventListener('click', function (event) {
        if (event.target.classList.contains('autocomplete-item')) {
            const dossierName = event.target.textContent;
            input.value = dossierName;
            suggestionsDiv.classList.remove('visible');
            // Mettre à jour le bouton clear après sélection
            toggleClearButton(clearButtonId, dossierName);
            if (mode === 'consult') {
                consultByDossier(dossierName);
            } else if (mode === 'modify') {
                modifyByDossier(dossierName);
            }
        }
    });
}

// ========================================
// VALIDATION ET SAUVEGARDE
// ========================================

async function validateSaisie() {
    try {
        const dateReunion = document.getElementById('saisir-date').value;

        if (!dateReunion) {
            alert('Veuillez sélectionner une date de réunion');
            return;
        }

        // Convertir la date en timestamp (valeur attendue par Grist)
        const dateReunionTimestamp = Math.floor(new Date(dateReunion).getTime() / 1000);

        // S'assurer que la date existe dans la table Agenda (sécurité OWASP)
        await ensureAgendaDateExists(dateReunionTimestamp);

        const dossiers = document.querySelectorAll('#saisir-dossiers .dossier-block');

        for (const dossier of dossiers) {
            // Validation et nettoyage des entrées utilisateur
            const intitule = validateInput(
                dossier.querySelector('.dossier-intitule').value,
                'text',
                200
            );
            const actions = validateInput(
                dossier.querySelector('.dossier-actions').value,
                'textarea',
                2000
            );
            const echeance = dossier.querySelector('.dossier-echeance').value;
            const etat = dossier.querySelector('.dossier-etat').value;

            const porteurs = Array.from(dossier.querySelectorAll('.dossier-porteurs .multi-select-option.selected'))
                .map(el => el.dataset.value)
                .map(name => getPersonneIdByName(name))
                .filter(id => id !== null);

            if (intitule) {
                const odjData = {
                    Date_de_la_reunion: dateReunionTimestamp,
                    Dossier: intitule,
                    Porteur_s_: ['L', ...porteurs],
                    Actions_a_mettre_en_uvre_etapes: actions,
                    Echeance: echeance || null,
                    Etat: getEtatIdByName(etat),
                    Enregistrement: Date.now() / 1000
                };

                await grist.docApi.applyUserActions([
                    ['AddRecord', 'ODJ', null, odjData]
                ]);
            }
        }

        alert('Données enregistrées avec succès !');

        await loadAllTables();
        await removeDuplicateRecords();
        await loadAllTables();
        populateConsultSelectors();
        resetSaisieForm();

    } catch (error) {
        console.error('Erreur lors de la sauvegarde:', error);
        alert('Erreur lors de l\'enregistrement des données: ' + error.message);
    }
}

function resetSaisieForm() {
    setDefaultDate();

    const dossiersContainer = document.getElementById('saisir-dossiers');
    if (dossiersContainer) {
        while (dossiersContainer.children.length > 1) {
            dossiersContainer.lastChild.remove();
        }

        const firstDossier = dossiersContainer.firstElementChild;
        if (firstDossier) {
            const intituleInput = firstDossier.querySelector('.dossier-intitule');
            const actionsInput = firstDossier.querySelector('.dossier-actions');
            const echeanceInput = firstDossier.querySelector('.dossier-echeance');
            const etatSelect = firstDossier.querySelector('.dossier-etat');

            if (intituleInput) intituleInput.value = '';
            if (actionsInput) actionsInput.value = '';
            if (echeanceInput) echeanceInput.value = '';
            if (etatSelect) etatSelect.value = '';

            firstDossier.querySelectorAll('.dossier-porteurs .multi-select-option').forEach(el => {
                el.classList.remove('selected');
            });
        }
    }

    currentDossierCount = 1;
}

// ========================================
// CONSULTATION
// ========================================

function handleConsultTypeChange(event) {
    const type = event.target.value;

    const dateSelector = document.getElementById('consult-date-selector');
    const dossierSelector = document.getElementById('consult-dossier-selector');
    const porteurSelector = document.getElementById('consult-porteur-selector');
    const echeanceSelector = document.getElementById('consult-echeance-selector');

    if (dateSelector) dateSelector.classList.toggle('hidden', type !== 'date');
    if (dossierSelector) dossierSelector.classList.toggle('hidden', type !== 'dossier');
    if (porteurSelector) porteurSelector.classList.toggle('hidden', type !== 'porteur');
    if (echeanceSelector) echeanceSelector.classList.toggle('hidden', type !== 'echeance');

    const resultsDiv = document.getElementById('consult-results');
    if (resultsDiv) resultsDiv.innerHTML = '';

    // Vider les sélections lors du changement de mode
    clearAllSelections();
    togglePrintButton();
}

function consultByDate() {
    const dateSelect = document.getElementById('consult-date-select');
    if (!dateSelect) return;

    const date = dateSelect.value;
    if (!date) return;

    const resultsDiv = document.getElementById('consult-results');
    if (!resultsDiv) return;

    const dateValue = typeof date === 'string' ? Number.parseFloat(date) : date;

    let dossiers = tablesData.ODJ.filter(o => o.Date_de_la_reunion == dateValue);

    // Trier par état (du pire au meilleur)
    dossiers = sortByEtat(dossiers);

    // Vider le conteneur
    resultsDiv.innerHTML = '';

    if (dossiers.length === 0) {
        const p = document.createElement('p');
        p.className = 'loading';
        p.textContent = 'Aucune donnée pour cette date';
        resultsDiv.appendChild(p);
        togglePrintButton();
        return;
    }

    // Construction sécurisée avec createElement
    const section = document.createElement('div');
    section.className = 'section';

    const title = document.createElement('h2');
    title.className = 'section-title';
    title.textContent = 'Ordre du jour';
    section.appendChild(title);

    const resultItem = document.createElement('div');
    resultItem.className = 'result-item';

    const header = document.createElement('div');
    header.className = 'result-header';
    header.textContent = `Date : ${formatDate(dateValue)}`;
    resultItem.appendChild(header);
    section.appendChild(resultItem);

    const tableContainer = document.createElement('div');
    tableContainer.className = 'table-container';

    const table = document.createElement('table');
    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');
    ['Dossier', 'Porteur(s)', 'Actions', 'Échéance', 'État'].forEach(headerText => {
        const th = document.createElement('th');
        th.textContent = headerText;
        headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);
    table.appendChild(thead);

    const tbody = document.createElement('tbody');

    dossiers.forEach(dossier => {
        const etatName = getEtatNameById(dossier.Etat);
        const etatClass = etatColorMap[etatName] || '';
        const porteurs = dossier.Porteur_s_ && dossier.Porteur_s_.length > 0
            ? dossier.Porteur_s_.map(id => getPersonneNameById(id)).filter(n => n).join(', ')
            : '';

        const tr = document.createElement('tr');
        if (etatClass) tr.className = etatClass;

        const tdDossier = document.createElement('td');
        tdDossier.textContent = dossier.Dossier || '';
        tr.appendChild(tdDossier);

        const tdPorteurs = document.createElement('td');
        tdPorteurs.textContent = porteurs;
        tr.appendChild(tdPorteurs);

        const tdActions = document.createElement('td');
        tdActions.textContent = dossier.Actions_a_mettre_en_uvre_etapes || '';
        tdActions.style.whiteSpace = 'pre-wrap';
        tr.appendChild(tdActions);

        const tdEcheance = document.createElement('td');
        tdEcheance.textContent = dossier.Echeance ? formatDate(dossier.Echeance) : '';
        tr.appendChild(tdEcheance);

        const tdEtat = document.createElement('td');
        tdEtat.textContent = etatName;
        tr.appendChild(tdEtat);

        tbody.appendChild(tr);
    });

    table.appendChild(tbody);
    tableContainer.appendChild(table);
    section.appendChild(tableContainer);
    resultsDiv.appendChild(section);

    togglePrintButton();
}

function consultByDossier(dossierName) {
    const resultsDiv = document.getElementById('consult-results');
    if (!resultsDiv) return;

    const dossiers = tablesData.ODJ.filter(odj => odj.Dossier === dossierName);

    resultsDiv.innerHTML = '';

    if (dossiers.length === 0) {
        const p = document.createElement('p');
        p.className = 'loading';
        p.textContent = 'Aucun dossier trouvé';
        resultsDiv.appendChild(p);
        togglePrintButton();
        return;
    }

    dossiers.sort((a, b) => {
        const dateA = a.Date_de_la_reunion ? new Date(a.Date_de_la_reunion) : new Date(0);
        const dateB = b.Date_de_la_reunion ? new Date(b.Date_de_la_reunion) : new Date(0);
        return dateB - dateA;
    });

    const section = document.createElement('div');
    section.className = 'section';

    const title = document.createElement('h2');
    title.className = 'section-title';
    title.textContent = 'Historique du dossier';
    section.appendChild(title);

    const tableContainer = document.createElement('div');
    tableContainer.className = 'table-container';

    const table = document.createElement('table');
    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');
    ['Date réunion', 'Porteur(s)', 'Actions', 'Échéance', 'État'].forEach(headerText => {
        const th = document.createElement('th');
        th.textContent = headerText;
        headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);
    table.appendChild(thead);

    const tbody = document.createElement('tbody');

    dossiers.forEach(dossier => {
        const etatName = getEtatNameById(dossier.Etat);
        const etatClass = etatColorMap[etatName] || '';
        const porteurs = dossier.Porteur_s_ && dossier.Porteur_s_.length > 0
            ? dossier.Porteur_s_.map(id => getPersonneNameById(id)).filter(n => n).join(', ')
            : '';

        const tr = document.createElement('tr');
        if (etatClass) tr.className = etatClass;

        const tdDate = document.createElement('td');
        tdDate.textContent = formatDate(dossier.Date_de_la_reunion);
        tr.appendChild(tdDate);

        const tdPorteurs = document.createElement('td');
        tdPorteurs.textContent = porteurs;
        tr.appendChild(tdPorteurs);

        const tdActions = document.createElement('td');
        tdActions.textContent = dossier.Actions_a_mettre_en_uvre_etapes || '';
        tdActions.style.whiteSpace = 'pre-wrap';
        tr.appendChild(tdActions);

        const tdEcheance = document.createElement('td');
        tdEcheance.textContent = dossier.Echeance ? formatDate(dossier.Echeance) : '';
        tr.appendChild(tdEcheance);

        const tdEtat = document.createElement('td');
        tdEtat.textContent = etatName;
        tr.appendChild(tdEtat);

        tbody.appendChild(tr);
    });

    table.appendChild(tbody);
    tableContainer.appendChild(table);
    section.appendChild(tableContainer);
    resultsDiv.appendChild(section);

    togglePrintButton();
}

function handlePorteurSelectChange() {
    const porteurSelect = document.getElementById('consult-porteur-select');
    if (!porteurSelect) return;

    const porteur = porteurSelect.value;
    const filters = document.getElementById('consult-porteur-filters');

    if (porteur && filters) {
        filters.classList.remove('hidden');
        consultByPorteur();
    } else if (filters) {
        filters.classList.add('hidden');
        const resultsDiv = document.getElementById('consult-results');
        if (resultsDiv) resultsDiv.innerHTML = '';
        togglePrintButton();
    }
}

function consultByPorteur() {
    const porteurSelect = document.getElementById('consult-porteur-select');
    if (!porteurSelect) return;

    const porteurName = porteurSelect.value;
    if (!porteurName) return;

    const porteurId = getPersonneIdByName(porteurName);

    const selectedEtats = new Set(
        Array.from(document.querySelectorAll('input[name="filter-etat"]:checked'))
            .map(cb => cb.value)
    );

    let dossiers = tablesData.ODJ.filter(odj => {
        if (!odj.Porteur_s_ || !Array.isArray(odj.Porteur_s_)) return false;
        return odj.Porteur_s_.includes(porteurId);
    });

    dossiers = dossiers.filter(dossier => {
        const etatName = getEtatNameById(dossier.Etat);
        return selectedEtats.has(etatName);
    });

    // Ne garder que le dernier état par dossier
    dossiers = getLatestEntriesPerDossier(dossiers);

    // Trier par état (du pire au meilleur)
    dossiers = sortByEtat(dossiers);

    const resultsDiv = document.getElementById('consult-results');
    if (!resultsDiv) return;

    resultsDiv.innerHTML = '';

    if (dossiers.length === 0) {
        const p = document.createElement('p');
        p.className = 'loading';
        p.textContent = 'Aucun dossier trouvé';
        resultsDiv.appendChild(p);
        togglePrintButton();
        return;
    }

    const section = document.createElement('div');
    section.className = 'section';

    const title = document.createElement('h2');
    title.className = 'section-title';
    title.textContent = 'Dossiers';
    section.appendChild(title);

    const tableContainer = document.createElement('div');
    tableContainer.className = 'table-container';

    const table = document.createElement('table');
    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');
    ['Date réunion', 'Dossier', 'Porteur(s)', 'Actions', 'Échéance', 'État'].forEach(headerText => {
        const th = document.createElement('th');
        th.textContent = headerText;
        headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);
    table.appendChild(thead);

    const tbody = document.createElement('tbody');

    dossiers.forEach(dossier => {
        const etatName = getEtatNameById(dossier.Etat);
        const etatClass = etatColorMap[etatName] || '';
        const porteurs = dossier.Porteur_s_ && dossier.Porteur_s_.length > 0
            ? dossier.Porteur_s_.map(id => getPersonneNameById(id)).filter(n => n).join(', ')
            : '';

        const tr = document.createElement('tr');
        if (etatClass) tr.className = etatClass;

        const tdDate = document.createElement('td');
        tdDate.textContent = formatDate(dossier.Date_de_la_reunion);
        tr.appendChild(tdDate);

        const tdDossier = document.createElement('td');
        tdDossier.textContent = dossier.Dossier || '';
        tr.appendChild(tdDossier);

        const tdPorteurs = document.createElement('td');
        tdPorteurs.textContent = porteurs;
        tr.appendChild(tdPorteurs);

        const tdActions = document.createElement('td');
        tdActions.textContent = dossier.Actions_a_mettre_en_uvre_etapes || '';
        tdActions.style.whiteSpace = 'pre-wrap';
        tr.appendChild(tdActions);

        const tdEcheance = document.createElement('td');
        tdEcheance.textContent = dossier.Echeance ? formatDate(dossier.Echeance) : '';
        tr.appendChild(tdEcheance);

        const tdEtat = document.createElement('td');
        tdEtat.textContent = etatName;
        tr.appendChild(tdEtat);

        tbody.appendChild(tr);
    });

    table.appendChild(tbody);
    tableContainer.appendChild(table);
    section.appendChild(tableContainer);
    resultsDiv.appendChild(section);

    togglePrintButton();
}

function consultByEcheance() {
    const echeanceSelect = document.getElementById('consult-echeance-select');
    if (!echeanceSelect) return;

    const echeance = echeanceSelect.value;
    if (!echeance) return;

    const resultsDiv = document.getElementById('consult-results');
    if (!resultsDiv) return;

    const echeanceValue = typeof echeance === 'string' ? Number.parseFloat(echeance) : echeance;

    let dossiers = tablesData.ODJ.filter(o => o.Echeance == echeanceValue);

    // Ne garder que le dernier état par dossier
    dossiers = getLatestEntriesPerDossier(dossiers);

    // Trier par état (du pire au meilleur)
    dossiers = sortByEtat(dossiers);

    resultsDiv.innerHTML = '';

    if (dossiers.length === 0) {
        const p = document.createElement('p');
        p.className = 'loading';
        p.textContent = 'Aucune donnée pour cette date';
        resultsDiv.appendChild(p);
        togglePrintButton();
        return;
    }

    const section = document.createElement('div');
    section.className = 'section';

    const title = document.createElement('h2');
    title.className = 'section-title';
    title.textContent = 'Dossiers avec échéance';
    section.appendChild(title);

    const resultItem = document.createElement('div');
    resultItem.className = 'result-item';

    const header = document.createElement('div');
    header.className = 'result-header';
    header.textContent = `Échéance : ${formatDate(echeanceValue)}`;
    resultItem.appendChild(header);
    section.appendChild(resultItem);

    const tableContainer = document.createElement('div');
    tableContainer.className = 'table-container';

    const table = document.createElement('table');
    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');
    ['Dossier', 'Porteur(s)', 'Actions', 'Date réunion', 'État'].forEach(headerText => {
        const th = document.createElement('th');
        th.textContent = headerText;
        headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);
    table.appendChild(thead);

    const tbody = document.createElement('tbody');

    dossiers.forEach(dossier => {
        const etatName = getEtatNameById(dossier.Etat);
        const etatClass = etatColorMap[etatName] || '';
        const porteurs = dossier.Porteur_s_ && dossier.Porteur_s_.length > 0
            ? dossier.Porteur_s_.map(id => getPersonneNameById(id)).filter(n => n).join(', ')
            : '';

        const tr = document.createElement('tr');
        if (etatClass) tr.className = etatClass;

        const tdDossier = document.createElement('td');
        tdDossier.textContent = dossier.Dossier || '';
        tr.appendChild(tdDossier);

        const tdPorteurs = document.createElement('td');
        tdPorteurs.textContent = porteurs;
        tr.appendChild(tdPorteurs);

        const tdActions = document.createElement('td');
        tdActions.textContent = dossier.Actions_a_mettre_en_uvre_etapes || '';
        tdActions.style.whiteSpace = 'pre-wrap';
        tr.appendChild(tdActions);

        const tdDate = document.createElement('td');
        tdDate.textContent = formatDate(dossier.Date_de_la_reunion);
        tr.appendChild(tdDate);

        const tdEtat = document.createElement('td');
        tdEtat.textContent = etatName;
        tr.appendChild(tdEtat);

        tbody.appendChild(tr);
    });

    table.appendChild(tbody);
    tableContainer.appendChild(table);
    section.appendChild(tableContainer);
    resultsDiv.appendChild(section);

    togglePrintButton();
}

// ========================================
// BOUTONS DE RÉINITIALISATION
// ========================================

function togglePrintButton() {
    const resultsDiv = document.getElementById('consult-results');
    const printButtonContainer = document.getElementById('consult-print-button-container');

    if (resultsDiv && printButtonContainer) {
        if (resultsDiv.innerHTML.trim() !== '' && !resultsDiv.innerHTML.includes('loading')) {
            printButtonContainer.classList.remove('hidden');
        } else {
            printButtonContainer.classList.add('hidden');
        }
    }
}

function printConsultResults() {
    window.print();
}

function printReunionResults() {
    window.print();
}

function toggleClearButton(buttonId, inputValue) {
    const button = document.getElementById(buttonId);
    if (button) {
        if (inputValue && inputValue.trim() !== '') {
            button.classList.add('visible');
        } else {
            button.classList.remove('visible');
        }
    }
}

function clearConsultDossier() {
    const input = document.getElementById('consult-dossier-input');
    if (input) {
        input.value = '';
        toggleClearButton('btn-clear-consult-dossier', '');
    }

    const resultsDiv = document.getElementById('consult-results');
    if (resultsDiv) {
        resultsDiv.innerHTML = '';
    }
    togglePrintButton();
}

function clearModifyDossier() {
    const input = document.getElementById('modify-dossier-input');
    if (input) {
        input.value = '';
        toggleClearButton('btn-clear-modify-dossier', '');
    }

    const resultsDiv = document.getElementById('modify-results');
    if (resultsDiv) {
        resultsDiv.innerHTML = '';
    }

    const buttons = document.getElementById('modify-buttons');
    if (buttons) {
        buttons.classList.add('hidden');
    }
}

// ========================================
// MODIFICATION
// ========================================

function handleModifyTypeChange(event) {
    const type = event.target.value;

    const dateSelector = document.getElementById('modify-date-selector');
    const dossierSelector = document.getElementById('modify-dossier-selector');
    const porteurSelector = document.getElementById('modify-porteur-selector');
    const echeanceSelector = document.getElementById('modify-echeance-selector');

    if (dateSelector) dateSelector.classList.toggle('hidden', type !== 'date');
    if (dossierSelector) dossierSelector.classList.toggle('hidden', type !== 'dossier');
    if (porteurSelector) porteurSelector.classList.toggle('hidden', type !== 'porteur');
    if (echeanceSelector) echeanceSelector.classList.toggle('hidden', type !== 'echeance');

    const resultsDiv = document.getElementById('modify-results');
    if (resultsDiv) resultsDiv.innerHTML = '';

    const buttons = document.getElementById('modify-buttons');
    if (buttons) buttons.classList.add('hidden');

    // Vider les sélections lors du changement de mode
    clearAllSelections();
}

function modifyByDate() {
    const dateSelect = document.getElementById('modify-date-select');
    if (!dateSelect) return;

    const date = dateSelect.value;
    if (!date) return;

    // Sauvegarder le contexte
    modifyContext.type = 'date';
    modifyContext.value = date;
    modifyContext.secondValue = null;

    const modifyResults = document.getElementById('modify-results');
    if (!modifyResults) return;

    const dateValue = typeof date === 'string' ? Number.parseFloat(date) : date;

    let dossiers = tablesData.ODJ.filter(o => o.Date_de_la_reunion == dateValue);

    // Trier par état (du pire au meilleur)
    dossiers = sortByEtat(dossiers);

    let html = '';

    if (dossiers.length > 0) {
        html += '<div class="section"><h2 class="section-title">Ordre du jour</h2>';
        html += `<div class="result-item">`;
        html += `<div class="result-header">Date : ${escapeHtml(formatDate(dateValue))}</div>`;
        html += '</div>';
        html += '<div class="table-container"><table>';
        html += '<thead><tr><th>Dossier</th><th>Porteur(s)</th><th>Actions</th><th>Échéance</th><th>État</th><th>Changement d\'état</th><th>Date du changement</th></tr></thead>';
        html += '<tbody>';

        dossiers.forEach(dossier => {
            const etatName = getEtatNameById(dossier.Etat);
            const etatClass = etatColorMap[etatName] || '';
            const porteurs = dossier.Porteur_s_ && dossier.Porteur_s_.length > 0
                ? dossier.Porteur_s_.map(id => getPersonneNameById(id)).filter(n => n).join(', ')
                : '';

            html += `<tr class="${escapeHtmlAttribute(etatClass)}" data-dossier-id="${escapeHtmlAttribute(dossier.id)}">`;
            html += `<td contenteditable="false">${escapeHtml(dossier.Dossier || '')}</td>`;
            html += `<td contenteditable="false">${escapeHtml(porteurs)}</td>`;
            html += `<td>${escapeHtml(dossier.Actions_a_mettre_en_uvre_etapes || '').replace(/\n/g, '<br>')}</td>`;
            html += `<td contenteditable="false">${escapeHtml(dossier.Echeance ? formatDate(dossier.Echeance) : '')}</td>`;
            html += `<td contenteditable="false">${escapeHtml(etatName)}</td>`;
            html += `<td><select class="etat-change-select"><option value="">-- Aucun changement --</option></select></td>`;
            html += `<td><input type="date" class="etat-change-date" disabled></td>`;
            html += `</tr>`;
        });

        html += '</tbody></table></div></div>';
    }

    modifyResults.innerHTML = html || '<p class="loading">Aucune donnée pour cette date</p>';

    makeFieldsEditable(modifyResults);

    const buttons = document.getElementById('modify-buttons');
    if (buttons) buttons.classList.remove('hidden');
}

function modifyByDossier(dossierName) {
    // Sauvegarder le contexte
    modifyContext.type = 'dossier';
    modifyContext.value = dossierName;
    modifyContext.secondValue = null;

    const resultsDiv = document.getElementById('modify-results');
    if (!resultsDiv) return;

    const dossiers = tablesData.ODJ.filter(odj => odj.Dossier === dossierName);

    if (dossiers.length === 0) {
        resultsDiv.innerHTML = '<p class="loading">Aucun dossier trouvé</p>';
        return;
    }

    dossiers.sort((a, b) => {
        const dateA = a.Date_de_la_reunion ? new Date(a.Date_de_la_reunion) : new Date(0);
        const dateB = b.Date_de_la_reunion ? new Date(b.Date_de_la_reunion) : new Date(0);
        return dateB - dateA;
    });

    let html = '<div class="section"><h2 class="section-title">Historique du dossier</h2>';
    html += '<div class="table-container"><table>';
    html += '<thead><tr><th>Date réunion</th><th>Porteur(s)</th><th>Actions</th><th>Échéance</th><th>État</th><th>Changement d\'état</th><th>Date du changement</th></tr></thead>';
    html += '<tbody>';

    dossiers.forEach(dossier => {
        const etatName = getEtatNameById(dossier.Etat);
        const etatClass = etatColorMap[etatName] || '';
        const porteurs = dossier.Porteur_s_ && dossier.Porteur_s_.length > 0
            ? dossier.Porteur_s_.map(id => getPersonneNameById(id)).filter(n => n).join(', ')
            : '';

        html += `<tr class="${escapeHtmlAttribute(etatClass)}" data-dossier-id="${escapeHtmlAttribute(dossier.id)}">`;
        html += `<td>${escapeHtml(formatDate(dossier.Date_de_la_reunion))}</td>`;
        html += `<td>${escapeHtml(porteurs)}</td>`;
        html += `<td>${escapeHtml(dossier.Actions_a_mettre_en_uvre_etapes || '').replace(/\n/g, '<br>')}</td>`;
        html += `<td>${escapeHtml(dossier.Echeance ? formatDate(dossier.Echeance) : '')}</td>`;
        html += `<td>${escapeHtml(etatName)}</td>`;
        html += `<td><select class="etat-change-select"><option value="">-- Aucun changement --</option></select></td>`;
        html += `<td><input type="date" class="etat-change-date" disabled></td>`;
        html += `</tr>`;
    });

    html += '</tbody></table></div></div>';
    resultsDiv.innerHTML = html;

    makeFieldsEditable(resultsDiv);

    const buttons = document.getElementById('modify-buttons');
    if (buttons) buttons.classList.remove('hidden');
}

function handleModifyPorteurSelectChange() {
    const porteurSelect = document.getElementById('modify-porteur-select');
    if (!porteurSelect) return;

    const porteurName = porteurSelect.value;
    const dossierSelectorDiv = document.getElementById('modify-porteur-dossier-selector');
    const dossierSelect = document.getElementById('modify-porteur-dossier-select');
    const filtersDiv = document.getElementById('modify-porteur-filters');

    if (!porteurName || !dossierSelectorDiv || !dossierSelect || !filtersDiv) return;

    // Trouver tous les dossiers du porteur
    const porteurId = getPersonneIdByName(porteurName);
    const dossiers = tablesData.ODJ.filter(odj => {
        if (!odj.Porteur_s_ || !Array.isArray(odj.Porteur_s_)) return false;
        return odj.Porteur_s_.includes(porteurId);
    });

    // Obtenir les noms de dossiers uniques
    const dossiersUniques = [...new Set(dossiers.map(d => d.Dossier).filter(Boolean))].sort();

    // Remplir le sélecteur de dossiers
    dossierSelect.innerHTML = '<option value="">-- Afficher tous les dossiers du porteur --</option>';
    dossiersUniques.forEach(dossier => {
        const option = document.createElement('option');
        option.value = dossier;
        option.textContent = dossier;
        dossierSelect.appendChild(option);
    });

    // Afficher le sélecteur de dossiers et les filtres
    dossierSelectorDiv.classList.remove('hidden');
    filtersDiv.classList.remove('hidden');

    // Peupler les cases à cocher d'état pour la modification
    populateModifyPorteurEtatFilters();

    // Attacher les événements pour tri et filtrage
    attachModifyPorteurFilterListeners();

    // Afficher tous les dossiers par défaut
    modifyByPorteurAllDossiers();

    // Cacher les boutons jusqu'à la sélection d'un dossier spécifique
    const buttons = document.getElementById('modify-buttons');
    if (buttons) buttons.classList.add('hidden');
}

function populateModifyPorteurEtatFilters() {
    const etats = getUniqueValues(tablesData.Menus, 'Etat');
    const filterContainer = document.getElementById('modify-filter-etat-checkboxes');

    if (!filterContainer) return;

    filterContainer.innerHTML = '';

    etats.forEach(etat => {
        const label = document.createElement('label');
        label.className = 'checkbox-label';

        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.name = 'modify-filter-etat';
        checkbox.value = etat;
        checkbox.checked = true;

        const span = document.createElement('span');
        span.textContent = etat;

        label.appendChild(checkbox);
        label.appendChild(span);
        filterContainer.appendChild(label);
    });
}

function attachModifyPorteurFilterListeners() {
    document.querySelectorAll('input[name="modify-filter-etat"]').forEach(checkbox => {
        checkbox.addEventListener('change', modifyByPorteurAllDossiers);
    });

    document.getElementById('modify-hide-expired')?.addEventListener('change', function () {
        // Vérifier si un dossier spécifique est sélectionné
        const dossierSelect = document.getElementById('modify-porteur-dossier-select');
        if (dossierSelect && dossierSelect.value !== '') {
            // Un dossier est sélectionné : appeler modifyByPorteurDossier
            modifyByPorteurDossier();
        } else {
            // Aucun dossier sélectionné : afficher tous les dossiers
            modifyByPorteurAllDossiers();
        }
    });

    document.getElementById('modify-porteur-dossier-select')?.addEventListener('change', function () {
        if (this.value === '') {
            // Aucun dossier sélectionné : afficher tous les dossiers
            modifyByPorteurAllDossiers();
        } else {
            // Dossier sélectionné : afficher l'historique de ce dossier
            modifyByPorteurDossier();
        }
    });
}

function modifyByPorteurAllDossiers() {
    const porteurSelect = document.getElementById('modify-porteur-select');
    if (!porteurSelect) return;

    const porteurName = porteurSelect.value;
    if (!porteurName) return;

    const porteurId = getPersonneIdByName(porteurName);
    const hideExpired = document.getElementById('modify-hide-expired')?.checked || false;

    const selectedEtats = new Set(
        Array.from(document.querySelectorAll('input[name="modify-filter-etat"]:checked'))
            .map(cb => cb.value)
    );

    let dossiers = tablesData.ODJ.filter(odj => {
        if (!odj.Porteur_s_ || !Array.isArray(odj.Porteur_s_)) return false;
        return odj.Porteur_s_.includes(porteurId);
    });

    dossiers = dossiers.filter(dossier => {
        const etatName = getEtatNameById(dossier.Etat);
        return selectedEtats.has(etatName);
    });

    // Filtrer les dossiers échus si le toggle est activé
    if (hideExpired) {
        const today = new Date();
        today.setHours(0, 0, 0, 0);

        dossiers = dossiers.filter(dossier => {
            if (!dossier.Echeance) return true; // Garder les dossiers sans échéance

            const echeanceDate = new Date(dossier.Echeance * 1000);
            echeanceDate.setHours(0, 0, 0, 0);
            return echeanceDate >= today;
        });
    }

    // Ne garder que le dernier état par dossier
    dossiers = getLatestEntriesPerDossier(dossiers);

    // Trier par état (du pire au meilleur)
    dossiers = sortByEtat(dossiers);

    const resultsDiv = document.getElementById('modify-results');
    if (!resultsDiv) return;

    if (dossiers.length === 0) {
        resultsDiv.innerHTML = '<p class="loading">Aucun dossier trouvé</p>';
        const buttons = document.getElementById('modify-buttons');
        if (buttons) buttons.classList.add('hidden');
        return;
    }

    // Grouper les dossiers par nom
    const dossiersByName = {};
    dossiers.forEach(dossier => {
        const nom = dossier.Dossier || 'Sans nom';
        if (!dossiersByName[nom]) {
            dossiersByName[nom] = [];
        }
        dossiersByName[nom].push(dossier);
    });

    // Créer un tableau d'objets avec le nom et les données du groupe
    const dossiersGroupes = Object.keys(dossiersByName).map(nom => {
        const group = dossiersByName[nom];

        // Obtenir la date représentative du groupe (la plus récente)
        const dateRepresentative = Math.max(...group.map(d => d.Date_de_la_reunion || 0));

        return {
            nom: nom,
            dossiers: group,
            dateRepresentative: dateRepresentative || 0
        };
    });

    // Trier les groupes par date représentative (du plus récent au plus ancien)
    dossiersGroupes.sort((a, b) => b.dateRepresentative - a.dateRepresentative);

    let html = '<div class="sections-container">';

    dossiersGroupes.forEach(groupe => {
        html += '<div class="section">';
        html += `<h2 class="section-title">${escapeHtml(groupe.nom)}</h2>`;
        html += '<div class="table-container"><table>';
        html += '<thead><tr><th>Date réunion</th><th>Porteur(s)</th><th>Actions</th><th>Échéance</th><th>État</th><th>Changement d\'état</th><th>Date du changement</th></tr></thead>';
        html += '<tbody>';

        groupe.dossiers.forEach(dossier => {
            const etatName = getEtatNameById(dossier.Etat);
            const etatClass = etatColorMap[etatName] || '';
            const porteurs = dossier.Porteur_s_ && dossier.Porteur_s_.length > 0
                ? dossier.Porteur_s_.map(id => getPersonneNameById(id)).filter(n => n).join(', ')
                : '';

            html += `<tr class="${escapeHtmlAttribute(etatClass)}" data-dossier-id="${escapeHtmlAttribute(dossier.id)}">`;
            html += `<td>${escapeHtml(formatDate(dossier.Date_de_la_reunion))}</td>`;
            html += `<td>${escapeHtml(porteurs)}</td>`;
            html += `<td>${escapeHtml(dossier.Actions_a_mettre_en_uvre_etapes || '').replace(/\n/g, '<br>')}</td>`;
            html += `<td>${escapeHtml(dossier.Echeance ? formatDate(dossier.Echeance) : '')}</td>`;
            html += `<td>${escapeHtml(etatName)}</td>`;
            html += `<td><select class="etat-change-select"><option value="">-- Aucun changement --</option></select></td>`;
            html += `<td><input type="date" class="etat-change-date" disabled></td>`;
            html += `</tr>`;
        });

        html += '</tbody></table></div>';
        html += '</div>';
    });

    html += '</div>';
    resultsDiv.innerHTML = html;

    makeFieldsEditable(resultsDiv);

    const buttons = document.getElementById('modify-buttons');
    if (buttons) buttons.classList.remove('hidden');
}

function modifyByPorteurDossier() {
    const porteurSelect = document.getElementById('modify-porteur-select');
    const dossierSelect = document.getElementById('modify-porteur-dossier-select');

    if (!porteurSelect || !dossierSelect) return;

    const porteurName = porteurSelect.value;
    const dossierName = dossierSelect.value;

    if (!porteurName || !dossierName) return;

    // Sauvegarder le contexte
    modifyContext.type = 'porteur';
    modifyContext.value = porteurName;
    modifyContext.secondValue = dossierName;

    // Afficher l'historique du dossier
    const resultsDiv = document.getElementById('modify-results');
    if (!resultsDiv) return;

    const dossiers = tablesData.ODJ.filter(odj => odj.Dossier === dossierName);

    if (dossiers.length === 0) {
        resultsDiv.innerHTML = '<p class="loading">Aucun dossier trouvé</p>';
        return;
    }

    dossiers.sort((a, b) => {
        const dateA = a.Date_de_la_reunion ? new Date(a.Date_de_la_reunion) : new Date(0);
        const dateB = b.Date_de_la_reunion ? new Date(b.Date_de_la_reunion) : new Date(0);
        return dateB - dateA;
    });

    let html = `<div class="section"><h2 class="section-title">Historique du dossier : ${escapeHtml(dossierName)}</h2>`;
    html += '<div class="table-container"><table>';
    html += '<thead><tr><th>Date réunion</th><th>Porteur(s)</th><th>Actions</th><th>Échéance</th><th>État</th><th>Changement d\'état</th><th>Date du changement</th></tr></thead>';
    html += '<tbody>';

    dossiers.forEach(dossier => {
        const etatName = getEtatNameById(dossier.Etat);
        const etatClass = etatColorMap[etatName] || '';
        const porteurs = dossier.Porteur_s_ && dossier.Porteur_s_.length > 0
            ? dossier.Porteur_s_.map(id => getPersonneNameById(id)).filter(n => n).join(', ')
            : '';

        html += `<tr class="${escapeHtmlAttribute(etatClass)}" data-dossier-id="${escapeHtmlAttribute(dossier.id)}">`;
        html += `<td>${escapeHtml(formatDate(dossier.Date_de_la_reunion))}</td>`;
        html += `<td>${escapeHtml(porteurs)}</td>`;
        html += `<td>${escapeHtml(dossier.Actions_a_mettre_en_uvre_etapes || '').replace(/\n/g, '<br>')}</td>`;
        html += `<td>${escapeHtml(dossier.Echeance ? formatDate(dossier.Echeance) : '')}</td>`;
        html += `<td>${escapeHtml(etatName)}</td>`;
        html += `<td><select class="etat-change-select"><option value="">-- Aucun changement --</option></select></td>`;
        html += `<td><input type="date" class="etat-change-date" disabled></td>`;
        html += `</tr>`;
    });

    html += '</tbody></table></div></div>';
    resultsDiv.innerHTML = html;

    makeFieldsEditable(resultsDiv);

    const buttons = document.getElementById('modify-buttons');
    if (buttons) buttons.classList.remove('hidden');
}

function modifyByEcheance() {
    const echeanceSelect = document.getElementById('modify-echeance-select');
    if (!echeanceSelect) return;

    const echeance = echeanceSelect.value;
    if (!echeance) return;

    // Sauvegarder le contexte
    modifyContext.type = 'echeance';
    modifyContext.value = echeance;
    modifyContext.secondValue = null;

    const modifyResults = document.getElementById('modify-results');
    if (!modifyResults) return;

    const echeanceValue = typeof echeance === 'string' ? Number.parseFloat(echeance) : echeance;

    let dossiers = tablesData.ODJ.filter(o => o.Echeance == echeanceValue);

    if (dossiers.length === 0) {
        modifyResults.innerHTML = '<p class="loading">Aucun dossier trouvé</p>';
        return;
    }

    // Ne garder que le dernier état par dossier
    dossiers = getLatestEntriesPerDossier(dossiers);

    // Trier par état (du pire au meilleur)
    dossiers = sortByEtat(dossiers);

    let html = '<div class="section"><h2 class="section-title">Dossiers avec échéance</h2>';
    html += `<div class="result-item">`;
    html += `<div class="result-header">Échéance : ${escapeHtml(formatDate(echeanceValue))}</div>`;
    html += '</div>';
    html += '<div class="table-container"><table>';
    html += '<thead><tr><th>Dossier</th><th>Porteur(s)</th><th>Actions</th><th>Date réunion</th><th>État</th><th>Changement d\'\u00e9tat</th><th>Date du changement</th></tr></thead>';
    html += '<tbody>';

    dossiers.forEach(dossier => {
        const etatName = getEtatNameById(dossier.Etat);
        const etatClass = etatColorMap[etatName] || '';
        const porteurs = dossier.Porteur_s_ && dossier.Porteur_s_.length > 0
            ? dossier.Porteur_s_.map(id => getPersonneNameById(id)).filter(n => n).join(', ')
            : '';

        html += `<tr class="${escapeHtmlAttribute(etatClass)}" data-dossier-id="${escapeHtmlAttribute(dossier.id)}">`;
        html += `<td contenteditable="false">${escapeHtml(dossier.Dossier || '')}</td>`;
        html += `<td contenteditable="false">${escapeHtml(porteurs)}</td>`;
        html += `<td>${escapeHtml(dossier.Actions_a_mettre_en_uvre_etapes || '').replace(/\n/g, '<br>')}</td>`;
        html += `<td contenteditable="false">${escapeHtml(formatDate(dossier.Date_de_la_reunion))}</td>`;
        html += `<td contenteditable="false">${escapeHtml(etatName)}</td>`;
        html += `<td><select class="etat-change-select"><option value="">-- Aucun changement --</option></select></td>`;
        html += `<td><input type="date" class="etat-change-date" disabled></td>`;
        html += `</tr>`;
    });

    html += '</tbody></table></div></div>';
    modifyResults.innerHTML = html;

    makeFieldsEditable(modifyResults);

    const buttons = document.getElementById('modify-buttons');
    if (buttons) buttons.classList.remove('hidden');
}

function makeFieldsEditable(container) {
    // Récupérer la liste des porteurs
    const personnes = getUniqueValues(tablesData.Menus, 'Personnes').filter(p => p !== 'Autre');

    // Rendre les cellules éditables
    container.querySelectorAll('table tbody tr').forEach(row => {
        const cells = row.querySelectorAll('td');
        const dossierId = row.dataset.dossierId;
        const dossierData = tablesData.ODJ.find(d => d.id == dossierId);

        cells.forEach((td, index) => {
            // Index 0: Dossier (texte éditable)
            if (index === 0) {
                td.contentEditable = true;
                td.style.border = '1px solid #d9d9d9';
                td.style.backgroundColor = 'rgba(255, 255, 255, 0.5)';
                td.style.color = '#000';
            }
            // Index 1: Porteur(s) (select multiple)
            else if (index === 1 && dossierData) {
                const currentPorteurs = dossierData.Porteur_s_ || [];
                const select = document.createElement('select');
                select.multiple = true;
                select.style.width = '100%';
                select.style.minHeight = '60px';
                select.style.border = '1px solid #d9d9d9';
                select.style.backgroundColor = 'rgba(255, 255, 255, 0.9)';

                personnes.forEach(personne => {
                    const option = document.createElement('option');
                    option.value = personne;
                    option.textContent = personne;
                    const personneId = getPersonneIdByName(personne);
                    if (currentPorteurs.includes(personneId)) {
                        option.selected = true;
                    }
                    select.appendChild(option);
                });

                // Permettre la sélection/désélection sans maintenir Ctrl
                select.addEventListener('mousedown', function (e) {
                    e.preventDefault();
                    const option = e.target;
                    if (option.tagName === 'OPTION') {
                        option.selected = !option.selected;
                        select.focus();
                        // Réorganiser les options après la sélection
                        setTimeout(() => reorderSelectOptions(select), 10);
                    }
                });

                td.innerHTML = '';
                td.appendChild(select);
                td.style.padding = '4px';

                // Réorganiser les options au chargement initial
                reorderSelectOptions(select);
            }
            // Index 2: Actions (texte éditable avec sauts de ligne)
            else if (index === 2) {
                td.contentEditable = true;
                td.style.border = '1px solid #d9d9d9';
                td.style.backgroundColor = 'rgba(255, 255, 255, 0.5)';
                td.style.color = '#000';

                // Permettre les sauts de ligne avec Entrée (méthode sécurisée)
                td.addEventListener('keydown', function (event) {
                    if (event.key === 'Enter') {
                        event.preventDefault();
                        const selection = window.getSelection();
                        if (selection.rangeCount > 0) {
                            const range = selection.getRangeAt(0);
                            const br = document.createElement('br');
                            range.deleteContents();
                            range.insertNode(br);
                            range.setStartAfter(br);
                            range.setEndAfter(br);
                            selection.removeAllRanges();
                            selection.addRange(range);
                        }
                    }
                });
            }
            // Index 3: Échéance (date picker)
            else if (index === 3 && dossierData) {
                const dateInput = document.createElement('input');
                dateInput.type = 'date';
                dateInput.style.width = '100%';
                dateInput.style.border = '1px solid #d9d9d9';
                dateInput.style.backgroundColor = 'rgba(255, 255, 255, 0.9)';
                dateInput.style.padding = '4px';

                if (dossierData.Echeance) {
                    const date = new Date(dossierData.Echeance * 1000);
                    dateInput.value = date.toISOString().split('T')[0];
                }

                td.innerHTML = '';
                td.appendChild(dateInput);
                td.style.padding = '4px';
            }
        });
    });

    // Peupler les sélecteurs d'état
    const etats = getUniqueValues(tablesData.Menus, 'Etat');
    const ordreEtats = [
        "Clôturé",
        "Avance très bien",
        "Avance bien",
        "RAS",
        "Des tensions",
        "Forte difficulté, blocage",
        "Supprimer le dossier"
    ];

    const etatsTries = etats.sort((a, b) => {
        const indexA = ordreEtats.indexOf(a);
        const indexB = ordreEtats.indexOf(b);
        if (indexA !== -1 && indexB !== -1) return indexA - indexB;
        if (indexA !== -1) return -1;
        if (indexB !== -1) return 1;
        return a.localeCompare(b);
    });

    container.querySelectorAll('.etat-change-select').forEach(select => {
        etatsTries.forEach(etat => {
            const option = document.createElement('option');
            option.value = etat;
            option.textContent = etat;
            select.appendChild(option);
        });

        // Gérer le changement de sélection
        select.addEventListener('change', function () {
            const row = this.closest('tr');
            const dateInput = row.querySelector('.etat-change-date');

            if (this.value === 'Supprimer le dossier') {
                // Désactiver le champ date pour la suppression (pas nécessaire)
                dateInput.disabled = true;
                dateInput.style.setProperty('background-color', '#a4a4a4', 'important');
                dateInput.style.color = '#3b3b3bff';
                dateInput.value = '';
            } else if (this.value) {
                dateInput.disabled = false;
                dateInput.style.setProperty('background-color', '#fff', 'important');
                if (!dateInput.value) {
                    dateInput.value = new Date().toISOString().split('T')[0];
                }

                // Changer la couleur de la ligne immédiatement
                const nouvelEtatClass = etatColorMap[this.value] || '';
                // Retirer toutes les classes d'état existantes
                Object.values(etatColorMap).forEach(cls => row.classList.remove(cls));
                // Ajouter la nouvelle classe d'état
                if (nouvelEtatClass) {
                    row.classList.add(nouvelEtatClass);
                }
            } else {
                dateInput.disabled = true;
                dateInput.style.setProperty('background-color', '#a4a4a4', 'important');
                dateInput.style.color = '#3b3b3bff';
                dateInput.value = '';

                // Restaurer la couleur originale si aucun changement
                const dossierId = row.dataset.dossierId;
                const dossierData = tablesData.ODJ.find(d => d.id == dossierId);
                if (dossierData) {
                    const etatOriginal = getEtatNameById(dossierData.Etat);
                    const etatOriginalClass = etatColorMap[etatOriginal] || '';
                    // Retirer toutes les classes d'état existantes
                    Object.values(etatColorMap).forEach(cls => row.classList.remove(cls));
                    // Ajouter la classe d'état originale
                    if (etatOriginalClass) {
                        row.classList.add(etatOriginalClass);
                    }
                }
            }
        });
    });

    // Initialiser la couleur de fond des champs de date désactivés
    container.querySelectorAll('.etat-change-date').forEach(dateInput => {
        dateInput.style.setProperty('background-color', '#a4a4a4', 'important');
        dateInput.style.color = '#3b3b3bff';
    });
}

function reorderSelectOptions(select) {
    if (!select || select.tagName !== 'SELECT') return;

    const options = Array.from(select.options);
    const selected = options.filter(opt => opt.selected);
    const notSelected = options.filter(opt => !opt.selected);

    // Vider le select
    select.innerHTML = '';

    // Ajouter d'abord les sélectionnés, puis les non-sélectionnés
    [...selected, ...notSelected].forEach(option => {
        select.appendChild(option);
    });
}

async function saveModifications() {
    try {
        const modifyResults = document.getElementById('modify-results');
        if (!modifyResults) return;

        const modifyType = document.querySelector('input[name="modify-type"]:checked');
        if (!modifyType) return;

        const type = modifyType.value;

        if (type === 'date') {
            await saveModificationsByDate();
        } else if (type === 'dossier') {
            await saveModificationsByDossier();
        } else if (type === 'porteur') {
            await saveModificationsByPorteur();
        } else if (type === 'echeance') {
            await saveModificationsByEcheance();
        }

        alert('Modifications enregistrées avec succès !');
        await loadAllTables();
        await removeDuplicateRecords();
        await loadAllTables();
        populateConsultSelectors();

        // Rouvrir le formulaire avec le contexte sauvegardé
        reopenModifyForm();

    } catch (error) {
        console.error('Erreur lors de la sauvegarde:', error);
        alert('Erreur lors de l\'enregistrement des modifications: ' + error.message);
    }
}

function reopenModifyForm() {
    if (!modifyContext.type || !modifyContext.value) return;

    // Restaurer le type de modification sélectionné
    const typeRadio = document.querySelector(`input[name="modify-type"][value="${modifyContext.type}"]`);
    if (typeRadio) {
        typeRadio.checked = true;
        // Déclencher l'événement pour afficher le bon sélecteur
        handleModifyTypeChange({ target: typeRadio });
    }

    // Restaurer la valeur sélectionnée selon le type
    if (modifyContext.type === 'date') {
        const dateSelect = document.getElementById('modify-date-select');
        if (dateSelect) {
            dateSelect.value = modifyContext.value;
            modifyByDate();
        }
    } else if (modifyContext.type === 'dossier') {
        const dossierInput = document.getElementById('modify-dossier-input');
        if (dossierInput) {
            dossierInput.value = modifyContext.value;
            toggleClearButton('btn-clear-modify-dossier', modifyContext.value);
            modifyByDossier(modifyContext.value);
        }
    } else if (modifyContext.type === 'porteur') {
        const porteurSelect = document.getElementById('modify-porteur-select');
        if (porteurSelect) {
            porteurSelect.value = modifyContext.value;
            handleModifyPorteurSelectChange();

            // Restaurer le dossier sélectionné
            setTimeout(() => {
                const dossierSelect = document.getElementById('modify-porteur-dossier-select');
                if (dossierSelect && modifyContext.secondValue) {
                    dossierSelect.value = modifyContext.secondValue;
                    modifyByPorteurDossier();
                }
            }, 100);
        }
    } else if (modifyContext.type === 'echeance') {
        const echeanceSelect = document.getElementById('modify-echeance-select');
        if (echeanceSelect) {
            echeanceSelect.value = modifyContext.value;
            modifyByEcheance();
        }
    }
}

async function saveModificationsByDate() {
    const dateSelect = document.getElementById('modify-date-select');
    if (!dateSelect) return;

    const date = dateSelect.value;
    if (!date) return;

    const modifyResults = document.getElementById('modify-results');

    const table = modifyResults.querySelector('table tbody');
    if (!table) return;

    const rows = table.querySelectorAll('tr');

    // Première passe : collecter tous les dossiers à supprimer
    const dossiersASupprimer = [];
    const rowsData = [];

    for (let index = 0; index < rows.length; index++) {
        const row = rows[index];
        const dossierId = Number.parseInt(row.dataset.dossierId);

        // Trouver le dossier correspondant par son ID
        const dossier = tablesData.ODJ.find(d => d.id === dossierId);
        if (!dossier) continue;

        const cells = row.querySelectorAll('td');
        const etatChangeSelect = row.querySelector('.etat-change-select');
        const etatChangeDateInput = row.querySelector('.etat-change-date');

        // Récupérer les valeurs modifiées
        const nouveauDossier = cells[0].textContent.trim();
        const porteurSelect = cells[1].querySelector('select');
        const nouveauxPorteurs = porteurSelect ?
            Array.from(porteurSelect.selectedOptions).map(opt => getPersonneIdByName(opt.value)).filter(id => id !== null) :
            dossier.Porteur_s_;

        // Récupérer les actions en nettoyant le texte
        let actions = cells[2].textContent.trim();
        if (!actions && cells[2].innerHTML) {
            actions = cells[2].innerHTML.replace(/<br\s*\/?>/gi, '\n').replace(/<[^>]+>/g, '').trim();
        }

        const echeanceInput = cells[3].querySelector('input[type="date"]');
        let nouvelleEcheance = dossier.Echeance;
        if (echeanceInput && echeanceInput.value) {
            nouvelleEcheance = Math.floor(new Date(echeanceInput.value).getTime() / 1000);
        }

        // Vérifier s'il y a un changement d'état
        const nouvelEtat = etatChangeSelect ? etatChangeSelect.value : '';
        const dateChangement = etatChangeDateInput && etatChangeDateInput.value ? Math.floor(new Date(etatChangeDateInput.value).getTime() / 1000) : null;

        // Stocker les données pour traitement ultérieur
        rowsData.push({
            dossier,
            nouveauDossier,
            nouveauxPorteurs,
            actions,
            nouvelleEcheance,
            nouvelEtat,
            dateChangement
        });

        // Identifier les dossiers à supprimer
        if (nouvelEtat === 'Supprimer le dossier') {
            dossiersASupprimer.push({ id: dossier.id, nom: nouveauDossier });
        }
    }

    // Si des dossiers doivent être supprimés, demander une confirmation unique
    if (dossiersASupprimer.length > 0) {
        const listeDossiers = dossiersASupprimer.map(d => `- ${d.nom}`).join('\n');
        const message = `ATTENTION ! Les dossiers suivants vont être supprimés :\n\n${listeDossiers}\n\nÊtes-vous sûr de vouloir confirmer cette action\u00A0?`;

        const confirmation = confirm(message);

        if (confirmation) {
            // Supprimer tous les dossiers confirmés
            for (const dossier of dossiersASupprimer) {
                await grist.docApi.applyUserActions([
                    ['RemoveRecord', 'ODJ', dossier.id]
                ]);
            }
        } else {
            // Annuler toute l'opération si l'utilisateur refuse
            return;
        }
    }

    // Deuxième passe : traiter les autres modifications
    for (const data of rowsData) {
        const { dossier, nouveauDossier, nouveauxPorteurs, actions, nouvelleEcheance, nouvelEtat, dateChangement } = data;

        // Ignorer les dossiers qui ont été supprimés
        if (nouvelEtat === 'Supprimer le dossier') {
            continue;
        } else if (nouvelEtat && dateChangement) {
            // Vérifier si la date saisie est identique à la date de réunion existante (comparaison en timestamps)
            const dateIdentique = dateChangement === dossier.Date_de_la_reunion;

            if (dateIdentique) {
                // Date identique : mettre à jour la ligne existante avec le nouvel état
                const nouveauEtatId = getEtatIdByName(nouvelEtat);

                await grist.docApi.applyUserActions([
                    ['UpdateRecord', 'ODJ', dossier.id, {
                        Dossier: nouveauDossier,
                        Porteur_s_: ['L', ...nouveauxPorteurs],
                        Actions_a_mettre_en_uvre_etapes: actions,
                        Echeance: nouvelleEcheance,
                        Etat: nouveauEtatId
                    }]
                ]);
            } else {
                // Date différente : créer une nouvelle ligne avec le nouvel état
                const nouveauEtatId = getEtatIdByName(nouvelEtat);

                // S'assurer que la nouvelle date existe dans la table Agenda (sécurité OWASP)
                await ensureAgendaDateExists(dateChangement);

                const odjData = {
                    Date_de_la_reunion: dateChangement,
                    Dossier: nouveauDossier,
                    Porteur_s_: ['L', ...nouveauxPorteurs],
                    Actions_a_mettre_en_uvre_etapes: actions,
                    Echeance: nouvelleEcheance,
                    Etat: nouveauEtatId
                };

                await grist.docApi.applyUserActions([
                    ['AddRecord', 'ODJ', null, odjData]
                ]);
            }
        } else if (nouvelEtat && !dateChangement) {
            // Mettre à jour la ligne existante avec le nouvel état (pas de nouvelle ligne)
            const nouveauEtatId = getEtatIdByName(nouvelEtat);

            await grist.docApi.applyUserActions([
                ['UpdateRecord', 'ODJ', dossier.id, {
                    Dossier: nouveauDossier,
                    Porteur_s_: ['L', ...nouveauxPorteurs],
                    Actions_a_mettre_en_uvre_etapes: actions,
                    Echeance: nouvelleEcheance,
                    Etat: nouveauEtatId
                }]
            ]);
        } else {
            // Mise à jour de la ligne existante avec toutes les modifications
            await grist.docApi.applyUserActions([
                ['UpdateRecord', 'ODJ', dossier.id, {
                    Dossier: nouveauDossier,
                    Porteur_s_: ['L', ...nouveauxPorteurs],
                    Actions_a_mettre_en_uvre_etapes: actions,
                    Echeance: nouvelleEcheance
                }]
            ]);
        }
    }
}

async function saveModificationsByDossier() {
    const modifyResults = document.getElementById('modify-results');
    if (!modifyResults) return;

    const table = modifyResults.querySelector('table tbody');
    if (!table) return;

    const rows = table.querySelectorAll('tr');

    let dossierName = '';
    const dossierInput = document.getElementById('modify-dossier-input');
    if (dossierInput) {
        dossierName = dossierInput.value;
    }

    if (!dossierName) return;

    // Première passe : collecter tous les dossiers à supprimer
    const dossiersASupprimer = [];
    const rowsData = [];

    for (let index = 0; index < rows.length; index++) {
        const row = rows[index];
        const dossierId = Number.parseInt(row.dataset.dossierId);

        // Trouver le dossier correspondant par son ID
        const dossier = tablesData.ODJ.find(d => d.id === dossierId);
        if (!dossier) continue;

        const cells = row.querySelectorAll('td');
        const etatChangeSelect = row.querySelector('.etat-change-select');
        const etatChangeDateInput = row.querySelector('.etat-change-date');

        // Récupérer les valeurs modifiées (index : 0=Date réunion, 1=Porteur(s), 2=Actions, 3=Échéance, 4=État)
        const porteurSelect = cells[1].querySelector('select');
        const nouveauxPorteurs = porteurSelect ?
            Array.from(porteurSelect.selectedOptions).map(opt => getPersonneIdByName(opt.value)).filter(id => id !== null) :
            dossier.Porteur_s_;

        // Récupérer les actions en nettoyant le texte
        let actions = cells[2].textContent.trim();
        if (!actions && cells[2].innerHTML) {
            actions = cells[2].innerHTML.replace(/<br\s*\/?>/gi, '\n').replace(/<[^>]+>/g, '').trim();
        }

        const echeanceInput = cells[3].querySelector('input[type="date"]');
        let nouvelleEcheance = dossier.Echeance;
        if (echeanceInput && echeanceInput.value) {
            nouvelleEcheance = Math.floor(new Date(echeanceInput.value).getTime() / 1000);
        }

        const nouveauDossier = dossierName;

        // Vérifier s'il y a un changement d'état
        const nouvelEtat = etatChangeSelect ? etatChangeSelect.value : '';
        const dateChangement = etatChangeDateInput && etatChangeDateInput.value ? Math.floor(new Date(etatChangeDateInput.value).getTime() / 1000) : null;

        // Stocker les données pour traitement ultérieur
        rowsData.push({
            dossier,
            nouveauDossier,
            nouveauxPorteurs,
            actions,
            nouvelleEcheance,
            nouvelEtat,
            dateChangement
        });

        // Identifier les dossiers à supprimer
        if (nouvelEtat === 'Supprimer le dossier') {
            dossiersASupprimer.push({ id: dossier.id, nom: nouveauDossier });
        }
    }

    // Si des dossiers doivent être supprimés, demander une confirmation unique
    if (dossiersASupprimer.length > 0) {
        const listeDossiers = dossiersASupprimer.map(d => `- ${d.nom}`).join('\n');
        const message = `ATTENTION ! Les dossiers suivants vont être supprimés :\n\n${listeDossiers}\n\nÊtes-vous sûr de vouloir confirmer cette action\u00A0?`;

        const confirmation = confirm(message);

        if (confirmation) {
            // Supprimer tous les dossiers confirmés
            for (const dossier of dossiersASupprimer) {
                await grist.docApi.applyUserActions([
                    ['RemoveRecord', 'ODJ', dossier.id]
                ]);
            }
        } else {
            // Annuler toute l'opération si l'utilisateur refuse
            return;
        }
    }

    // Deuxième passe : traiter les autres modifications
    for (const data of rowsData) {
        const { dossier, nouveauDossier, nouveauxPorteurs, actions, nouvelleEcheance, nouvelEtat, dateChangement } = data;

        // Ignorer les dossiers qui ont été supprimés
        if (nouvelEtat === 'Supprimer le dossier') {
            continue;
        } else if (nouvelEtat && dateChangement) {
            // Vérifier si la date saisie est identique à la date de réunion existante (comparaison en timestamps)
            const dateIdentique = dateChangement === dossier.Date_de_la_reunion;

            if (dateIdentique) {
                // Date identique : mettre à jour la ligne existante avec le nouvel état
                const nouveauEtatId = getEtatIdByName(nouvelEtat);

                await grist.docApi.applyUserActions([
                    ['UpdateRecord', 'ODJ', dossier.id, {
                        Dossier: nouveauDossier,
                        Porteur_s_: ['L', ...nouveauxPorteurs],
                        Actions_a_mettre_en_uvre_etapes: actions,
                        Echeance: nouvelleEcheance,
                        Etat: nouveauEtatId
                    }]
                ]);
            } else {
                // Date différente : créer une nouvelle ligne avec le nouvel état
                const nouveauEtatId = getEtatIdByName(nouvelEtat);

                // S'assurer que la nouvelle date existe dans la table Agenda (sécurité OWASP)
                await ensureAgendaDateExists(dateChangement);

                const odjData = {
                    Date_de_la_reunion: dateChangement,
                    Dossier: nouveauDossier,
                    Porteur_s_: ['L', ...nouveauxPorteurs],
                    Actions_a_mettre_en_uvre_etapes: actions,
                    Echeance: nouvelleEcheance,
                    Etat: nouveauEtatId,
                    Enregistrement: Date.now() / 1000
                };

                await grist.docApi.applyUserActions([
                    ['AddRecord', 'ODJ', null, odjData]
                ]);
            }
        } else if (nouvelEtat && !dateChangement) {
            // Mettre à jour la ligne existante avec le nouvel état (pas de nouvelle ligne)
            const nouveauEtatId = getEtatIdByName(nouvelEtat);

            await grist.docApi.applyUserActions([
                ['UpdateRecord', 'ODJ', dossier.id, {
                    Dossier: nouveauDossier,
                    Porteur_s_: ['L', ...nouveauxPorteurs],
                    Actions_a_mettre_en_uvre_etapes: actions,
                    Echeance: nouvelleEcheance,
                    Etat: nouveauEtatId
                }]
            ]);
        } else {
            // Mise à jour de la ligne existante avec toutes les modifications
            await grist.docApi.applyUserActions([
                ['UpdateRecord', 'ODJ', dossier.id, {
                    Dossier: nouveauDossier,
                    Porteur_s_: ['L', ...nouveauxPorteurs],
                    Actions_a_mettre_en_uvre_etapes: actions,
                    Echeance: nouvelleEcheance
                }]
            ]);
        }
    }
}

async function saveModificationsByPorteur() {
    const modifyResults = document.getElementById('modify-results');
    if (!modifyResults) return;

    const table = modifyResults.querySelector('table tbody');
    if (!table) return;

    const rows = table.querySelectorAll('tr');

    // Récupérer le porteur sélectionné
    const porteurSelect = document.getElementById('modify-porteur-select');
    if (!porteurSelect) return;

    const porteurName = porteurSelect.value;
    const porteurId = getPersonneIdByName(porteurName);

    if (!porteurName || !porteurId) return;

    // Récupérer le dossier sélectionné
    const dossierSelect = document.getElementById('modify-porteur-dossier-select');
    if (!dossierSelect) return;

    const dossierName = dossierSelect.value;
    if (!dossierName) return;

    // Première passe : collecter tous les dossiers à supprimer
    const dossiersASupprimer = [];
    const rowsData = [];

    for (let index = 0; index < rows.length; index++) {
        const row = rows[index];
        const dossierId = Number.parseInt(row.dataset.dossierId);

        // Trouver le dossier correspondant par son ID
        const dossier = tablesData.ODJ.find(d => d.id === dossierId);
        if (!dossier) continue;

        const cells = row.querySelectorAll('td');
        const etatChangeSelect = row.querySelector('.etat-change-select');
        const etatChangeDateInput = row.querySelector('.etat-change-date');

        // Récupérer les valeurs modifiées (index : 0=Date réunion, 1=Porteur(s), 2=Actions, 3=Échéance, 4=État)
        const porteurSelectCell = cells[1].querySelector('select');
        const nouveauxPorteurs = porteurSelectCell ?
            Array.from(porteurSelectCell.selectedOptions).map(opt => getPersonneIdByName(opt.value)).filter(id => id !== null) :
            dossier.Porteur_s_;

        // Récupérer les actions en nettoyant le texte
        let actions = cells[2].textContent.trim();
        if (!actions && cells[2].innerHTML) {
            actions = cells[2].innerHTML.replace(/<br\s*\/?>/gi, '\n').replace(/<[^>]+>/g, '').trim();
        }

        const echeanceInput = cells[3].querySelector('input[type="date"]');
        let nouvelleEcheance = dossier.Echeance;
        if (echeanceInput && echeanceInput.value) {
            nouvelleEcheance = Math.floor(new Date(echeanceInput.value).getTime() / 1000);
        }

        const nouveauDossier = dossierName;

        // Vérifier s'il y a un changement d'état
        const nouvelEtat = etatChangeSelect ? etatChangeSelect.value : '';
        const dateChangement = etatChangeDateInput && etatChangeDateInput.value ? Math.floor(new Date(etatChangeDateInput.value).getTime() / 1000) : null;

        // Stocker les données pour traitement ultérieur
        rowsData.push({
            dossier,
            nouveauDossier,
            nouveauxPorteurs,
            actions,
            nouvelleEcheance,
            nouvelEtat,
            dateChangement
        });

        // Identifier les dossiers à supprimer
        if (nouvelEtat === 'Supprimer le dossier') {
            dossiersASupprimer.push({ id: dossier.id, nom: nouveauDossier });
        }
    }

    // Si des dossiers doivent être supprimés, demander une confirmation unique
    if (dossiersASupprimer.length > 0) {
        const listeDossiers = dossiersASupprimer.map(d => `- ${d.nom}`).join('\n');
        const message = `ATTENTION ! Les dossiers suivants vont être supprimés :\n\n${listeDossiers}\n\nÊtes-vous sûr de vouloir confirmer cette action\u00A0?`;

        const confirmation = confirm(message);

        if (confirmation) {
            // Supprimer tous les dossiers confirmés
            for (const dossier of dossiersASupprimer) {
                await grist.docApi.applyUserActions([
                    ['RemoveRecord', 'ODJ', dossier.id]
                ]);
            }
        } else {
            // Annuler toute l'opération si l'utilisateur refuse
            return;
        }
    }

    // Deuxième passe : traiter les autres modifications
    for (const data of rowsData) {
        const { dossier, nouveauDossier, nouveauxPorteurs, actions, nouvelleEcheance, nouvelEtat, dateChangement } = data;

        // Ignorer les dossiers qui ont été supprimés
        if (nouvelEtat === 'Supprimer le dossier') {
            continue;
        } else if (nouvelEtat && dateChangement) {
            // Vérifier si la date saisie est identique à la date de réunion existante (comparaison en timestamps)
            const dateIdentique = dateChangement === dossier.Date_de_la_reunion;

            if (dateIdentique) {
                // Date identique : mettre à jour la ligne existante avec le nouvel état
                const nouveauEtatId = getEtatIdByName(nouvelEtat);

                await grist.docApi.applyUserActions([
                    ['UpdateRecord', 'ODJ', dossier.id, {
                        Dossier: nouveauDossier,
                        Porteur_s_: ['L', ...nouveauxPorteurs],
                        Actions_a_mettre_en_uvre_etapes: actions,
                        Echeance: nouvelleEcheance,
                        Etat: nouveauEtatId
                    }]
                ]);
            } else {
                // Date différente : créer une nouvelle ligne avec le nouvel état
                const nouveauEtatId = getEtatIdByName(nouvelEtat);

                // S'assurer que la nouvelle date existe dans la table Agenda (sécurité OWASP)
                await ensureAgendaDateExists(dateChangement);

                const odjData = {
                    Date_de_la_reunion: dateChangement,
                    Dossier: nouveauDossier,
                    Porteur_s_: ['L', ...nouveauxPorteurs],
                    Actions_a_mettre_en_uvre_etapes: actions,
                    Echeance: nouvelleEcheance,
                    Etat: nouveauEtatId,
                    Enregistrement: Date.now() / 1000
                };

                await grist.docApi.applyUserActions([
                    ['AddRecord', 'ODJ', null, odjData]
                ]);
            }
        } else if (nouvelEtat && !dateChangement) {
            // Mettre à jour la ligne existante avec le nouvel état (pas de nouvelle ligne)
            const nouveauEtatId = getEtatIdByName(nouvelEtat);

            await grist.docApi.applyUserActions([
                ['UpdateRecord', 'ODJ', dossier.id, {
                    Dossier: nouveauDossier,
                    Porteur_s_: ['L', ...nouveauxPorteurs],
                    Actions_a_mettre_en_uvre_etapes: actions,
                    Echeance: nouvelleEcheance,
                    Etat: nouveauEtatId
                }]
            ]);
        } else {
            // Mise à jour de la ligne existante avec toutes les modifications
            await grist.docApi.applyUserActions([
                ['UpdateRecord', 'ODJ', dossier.id, {
                    Dossier: nouveauDossier,
                    Porteur_s_: ['L', ...nouveauxPorteurs],
                    Actions_a_mettre_en_uvre_etapes: actions,
                    Echeance: nouvelleEcheance
                }]
            ]);
        }
    }
}

async function saveModificationsByEcheance() {
    const modifyResults = document.getElementById('modify-results');
    if (!modifyResults) return;

    const table = modifyResults.querySelector('table tbody');
    if (!table) return;

    const rows = table.querySelectorAll('tr');

    const echeanceSelect = document.getElementById('modify-echeance-select');
    if (!echeanceSelect) return;

    const echeance = echeanceSelect.value;
    if (!echeance) return;

    // Première passe : collecter tous les dossiers à supprimer
    const dossiersASupprimer = [];
    const rowsData = [];

    for (let index = 0; index < rows.length; index++) {
        const row = rows[index];
        const dossierId = Number.parseInt(row.dataset.dossierId);

        // Trouver le dossier correspondant par son ID
        const dossier = tablesData.ODJ.find(d => d.id === dossierId);
        if (!dossier) continue;

        const cells = row.querySelectorAll('td');
        const etatChangeSelect = row.querySelector('.etat-change-select');
        const etatChangeDateInput = row.querySelector('.etat-change-date');

        // Récupérer les valeurs modifiées
        const nouveauDossier = cells[0].textContent.trim();
        const porteurSelect = cells[1].querySelector('select');
        const nouveauxPorteurs = porteurSelect ?
            Array.from(porteurSelect.selectedOptions).map(opt => getPersonneIdByName(opt.value)).filter(id => id !== null) :
            dossier.Porteur_s_;

        // Récupérer les actions en nettoyant le texte
        let actions = cells[2].textContent.trim();
        if (!actions && cells[2].innerHTML) {
            actions = cells[2].innerHTML.replace(/<br\s*\/?>/gi, '\n').replace(/<[^>]+>/g, '').trim();
        }

        // Vérifier s'il y a un changement d'état
        const nouvelEtat = etatChangeSelect ? etatChangeSelect.value : '';
        const dateChangement = etatChangeDateInput && etatChangeDateInput.value ? Math.floor(new Date(etatChangeDateInput.value).getTime() / 1000) : null;

        // Stocker les données pour traitement ultérieur
        rowsData.push({
            dossier,
            nouveauDossier,
            nouveauxPorteurs,
            actions,
            nouvelEtat,
            dateChangement
        });

        // Identifier les dossiers à supprimer
        if (nouvelEtat === 'Supprimer le dossier') {
            dossiersASupprimer.push({ id: dossier.id, nom: nouveauDossier });
        }
    }

    // Si des dossiers doivent être supprimés, demander une confirmation unique
    if (dossiersASupprimer.length > 0) {
        const listeDossiers = dossiersASupprimer.map(d => `- ${d.nom}`).join('\n');
        const message = `ATTENTION ! Les dossiers suivants vont être supprimés :\n\n${listeDossiers}\n\nÊtes-vous sûr de vouloir confirmer cette action\u00A0?`;

        const confirmation = confirm(message);

        if (confirmation) {
            // Supprimer tous les dossiers confirmés
            for (const dossier of dossiersASupprimer) {
                await grist.docApi.applyUserActions([
                    ['RemoveRecord', 'ODJ', dossier.id]
                ]);
            }
        } else {
            // Annuler toute l'opération si l'utilisateur refuse
            return;
        }
    }

    // Deuxième passe : traiter les autres modifications
    for (const data of rowsData) {
        const { dossier, nouveauDossier, nouveauxPorteurs, actions, nouvelEtat, dateChangement } = data;

        // Ignorer les dossiers qui ont été supprimés
        if (nouvelEtat === 'Supprimer le dossier') {
            continue;
        } else if (nouvelEtat && dateChangement) {
            // Vérifier si la date saisie est identique à la date de réunion existante (comparaison en timestamps)
            const dateIdentique = dateChangement === dossier.Date_de_la_reunion;

            if (dateIdentique) {
                // Date identique : mettre à jour la ligne existante avec le nouvel état
                const nouveauEtatId = getEtatIdByName(nouvelEtat);

                await grist.docApi.applyUserActions([
                    ['UpdateRecord', 'ODJ', dossier.id, {
                        Dossier: nouveauDossier,
                        Porteur_s_: ['L', ...nouveauxPorteurs],
                        Actions_a_mettre_en_uvre_etapes: actions,
                        Etat: nouveauEtatId
                    }]
                ]);
            } else {
                // Date différente : créer une nouvelle ligne avec le nouvel état
                const nouveauEtatId = getEtatIdByName(nouvelEtat);

                // S'assurer que la nouvelle date existe dans la table Agenda (sécurité OWASP)
                await ensureAgendaDateExists(dateChangement);

                const odjData = {
                    Date_de_la_reunion: dateChangement,
                    Dossier: nouveauDossier,
                    Porteur_s_: ['L', ...nouveauxPorteurs],
                    Actions_a_mettre_en_uvre_etapes: actions,
                    Echeance: echeanceValue,
                    Etat: nouveauEtatId,
                    Enregistrement: Date.now() / 1000
                };

                await grist.docApi.applyUserActions([
                    ['AddRecord', 'ODJ', null, odjData]
                ]);
            }
        } else if (nouvelEtat && !dateChangement) {
            // Mettre à jour la ligne existante avec le nouvel état (pas de nouvelle ligne)
            const nouveauEtatId = getEtatIdByName(nouvelEtat);

            await grist.docApi.applyUserActions([
                ['UpdateRecord', 'ODJ', dossier.id, {
                    Dossier: nouveauDossier,
                    Porteur_s_: ['L', ...nouveauxPorteurs],
                    Actions_a_mettre_en_uvre_etapes: actions,
                    Etat: nouveauEtatId
                }]
            ]);
        } else {
            // Mise à jour de la ligne existante avec toutes les modifications
            await grist.docApi.applyUserActions([
                ['UpdateRecord', 'ODJ', dossier.id, {
                    Dossier: nouveauDossier,
                    Porteur_s_: ['L', ...nouveauxPorteurs],
                    Actions_a_mettre_en_uvre_etapes: actions
                }]
            ]);
        }
    }
}

function cancelModifications() {
    const resultsDiv = document.getElementById('modify-results');
    if (resultsDiv) resultsDiv.innerHTML = '';

    const buttons = document.getElementById('modify-buttons');
    if (buttons) buttons.classList.add('hidden');

    const dateSelect = document.getElementById('modify-date-select');
    if (dateSelect) dateSelect.value = '';
}

// ========================================
// ONGLET RÉUNION
// ========================================

function populateReunionDateSelect() {
    const select = document.getElementById('reunion-date-select');
    if (!select) return;

    const dates = getUniqueDates(tablesData.ODJ, 'Date_de_la_reunion');
    select.innerHTML = '<option value="">-- Choisir une date --</option>';

    // Déterminer la date par défaut (prochaine réunion à compter d'aujourd'hui inclus)
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    let defaultDate = null;

    dates.forEach(dateValue => {
        const date = typeof dateValue === 'number'
            ? new Date(dateValue * 1000)
            : new Date(dateValue);

        const option = document.createElement('option');
        option.value = dateValue;
        option.textContent = formatDate(dateValue);

        // Les dates sont triées en ordre décroissant (plus récente en premier)
        // On continue à itérer pour trouver la date la plus proche >= aujourd'hui
        if (date >= today) {
            defaultDate = dateValue;
        }

        select.appendChild(option);
    });

    // Définir la date par défaut si trouvée
    if (defaultDate !== null) {
        select.value = defaultDate;
    }
}

function reunionDisplayData() {
    const select = document.getElementById('reunion-date-select');
    if (!select) return;

    const dateValue = select.value;
    if (!dateValue) {
        clearReunionDisplay();
        return;
    }

    const numDateValue = typeof dateValue === 'string' ? Number.parseFloat(dateValue) : dateValue;

    // Récupérer les dossiers pour la réunion sélectionnée (Ordre du jour)
    let odjDossiers = tablesData.ODJ.filter(o => o.Date_de_la_reunion == numDateValue);
    // Garder uniquement la dernière version de chaque dossier dans l'ODJ
    odjDossiers = getLatestEntriesPerDossier(odjDossiers);

    // Créer un Set des noms de dossiers présents dans l'ODJ
    const odjDossierNames = new Set(odjDossiers.map(d => d.Dossier));

    // Récupérer les dossiers avec échéance correspondant à la date de réunion
    let dossierEcheance = tablesData.ODJ.filter(o => o.Echeance == numDateValue && o.Date_de_la_reunion != numDateValue);
    // Garder uniquement la dernière version de chaque dossier
    dossierEcheance = getLatestEntriesPerDossier(dossierEcheance);
    // Exclure les dossiers déjà présents dans l'ODJ
    dossierEcheance = dossierEcheance.filter(d => !odjDossierNames.has(d.Dossier));

    // Récupérer les dossiers échus non clôturés
    let dossierExpired = tablesData.ODJ.filter(o => {
        if (o.Echeance === null || o.Echeance === undefined) return false;

        const echeanceNum = typeof o.Echeance === 'number' ? o.Echeance : Number.parseFloat(o.Echeance);
        const reunionNum = typeof numDateValue === 'number' ? numDateValue : Number.parseFloat(numDateValue);

        // Dates antérieures à la réunion
        if (echeanceNum >= reunionNum) return false;

        // Statut n'est pas "Clôturé"
        const etatName = getEtatNameById(o.Etat);
        return etatName !== "Clôturé";
    });
    // Garder uniquement la dernière version de chaque dossier
    dossierExpired = getLatestEntriesPerDossier(dossierExpired);
    // Exclure les dossiers déjà présents dans l'ODJ
    dossierExpired = dossierExpired.filter(d => !odjDossierNames.has(d.Dossier));

    displayODJ(odjDossiers, numDateValue);
    displayDossierEcheance(dossierEcheance, numDateValue);
    displayExpiredDossiers(dossierExpired, numDateValue);

    // Afficher le bouton d'impression
    const printContainer = document.getElementById('reunion-print-button-container');
    if (printContainer) {
        if (odjDossiers.length > 0 || dossierEcheance.length > 0 || dossierExpired.length > 0) {
            printContainer.classList.remove('hidden');
        } else {
            printContainer.classList.add('hidden');
        }
    }
}

function displayODJ(dossiers, dateValue) {
    const container = document.getElementById('reunion-odj-table');
    if (!container) return;

    dossiers = sortByEtat(dossiers);

    let html = '';
    if (dossiers.length > 0) {
        html += '<table>';
        html += '<thead><tr><th>Dossier</th><th>Porteur(s)</th><th>Actions</th><th>Échéance</th><th>État</th><th>Changement d\'état</th><th>Date du changement</th></tr></thead>';
        html += '<tbody>';

        dossiers.forEach(dossier => {
            const etatName = getEtatNameById(dossier.Etat);
            const etatClass = etatColorMap[etatName] || '';
            const porteurs = dossier.Porteur_s_ && dossier.Porteur_s_.length > 0
                ? dossier.Porteur_s_.map(id => getPersonneNameById(id)).filter(n => n).join(', ')
                : '';

            html += `<tr class="${escapeHtmlAttribute(etatClass)}" data-dossier-id="${escapeHtmlAttribute(dossier.id)}">`;
            html += `<td contenteditable="false">${escapeHtml(dossier.Dossier || '')}</td>`;
            html += `<td contenteditable="false">${escapeHtml(porteurs)}</td>`;
            html += `<td>${escapeHtml(dossier.Actions_a_mettre_en_uvre_etapes || '').replace(/\n/g, '<br>')}</td>`;
            html += `<td contenteditable="false">${escapeHtml(dossier.Echeance ? formatDate(dossier.Echeance) : '')}</td>`;
            html += `<td contenteditable="false">${escapeHtml(etatName)}</td>`;
            html += `<td><select class="etat-change-select"><option value="">-- Aucun changement --</option></select></td>`;
            html += `<td><input type="date" class="etat-change-date" disabled></td>`;
            html += `</tr>`;
        });

        html += '</tbody></table>';
        container.innerHTML = html;
        makeFieldsEditableReunion(container);
    } else {
        html = '<p class="loading">Aucun dossier pour cette réunion</p>';
        container.innerHTML = html;
    }
}

function displayDossierEcheance(dossiers, dateValue) {
    const container = document.getElementById('reunion-echeance-table');
    if (!container) return;

    dossiers = getLatestEntriesPerDossier(dossiers);
    dossiers = sortByEtat(dossiers);

    let html = '';
    if (dossiers.length > 0) {
        html += '<table>';
        html += '<thead><tr><th>Dossier</th><th>Porteur(s)</th><th>Actions</th><th>Échéance</th><th>État</th><th>Changement d\'état</th><th>Date du changement</th></tr></thead>';
        html += '<tbody>';

        dossiers.forEach(dossier => {
            const etatName = getEtatNameById(dossier.Etat);
            const etatClass = etatColorMap[etatName] || '';
            const porteurs = dossier.Porteur_s_ && dossier.Porteur_s_.length > 0
                ? dossier.Porteur_s_.map(id => getPersonneNameById(id)).filter(n => n).join(', ')
                : '';

            html += `<tr class="${escapeHtmlAttribute(etatClass)}" data-dossier-id="${escapeHtmlAttribute(dossier.id)}">`;
            html += `<td contenteditable="false">${escapeHtml(dossier.Dossier || '')}</td>`;
            html += `<td contenteditable="false">${escapeHtml(porteurs)}</td>`;
            html += `<td>${escapeHtml(dossier.Actions_a_mettre_en_uvre_etapes || '').replace(/\n/g, '<br>')}</td>`;
            html += `<td contenteditable="false">${escapeHtml(formatDate(dossier.Date_de_la_reunion))}</td>`;
            html += `<td contenteditable="false">${escapeHtml(etatName)}</td>`;
            html += `<td><select class="etat-change-select"><option value="">-- Aucun changement --</option></select></td>`;
            html += `<td><input type="date" class="etat-change-date" disabled></td>`;
            html += `</tr>`;
        });

        html += '</tbody></table>';
        container.innerHTML = html;
        makeFieldsEditableReunion(container);
    } else {
        html = '<p class="loading">Aucun dossier à échéance</p>';
        container.innerHTML = html;
    }
}

function displayExpiredDossiers(dossiers, dateValue) {
    const container = document.getElementById('reunion-expired-table');
    if (!container) return;

    dossiers = getLatestEntriesPerDossier(dossiers);
    dossiers = sortByEtat(dossiers);

    let html = '';
    if (dossiers.length > 0) {
        html += '<table>';
        html += '<thead><tr><th>Dossier</th><th>Porteur(s)</th><th>Actions</th><th>Échéance</th><th>Date réunion</th><th>État</th><th>Changement d\'état</th><th>Date du changement</th></tr></thead>';
        html += '<tbody>';

        dossiers.forEach(dossier => {
            const etatName = getEtatNameById(dossier.Etat);
            const etatClass = etatColorMap[etatName] || '';
            const porteurs = dossier.Porteur_s_ && dossier.Porteur_s_.length > 0
                ? dossier.Porteur_s_.map(id => getPersonneNameById(id)).filter(n => n).join(', ')
                : '';

            html += `<tr class="${escapeHtmlAttribute(etatClass)}" data-dossier-id="${escapeHtmlAttribute(dossier.id)}">`;
            html += `<td contenteditable="false">${escapeHtml(dossier.Dossier || '')}</td>`;
            html += `<td contenteditable="false">${escapeHtml(porteurs)}</td>`;
            html += `<td>${escapeHtml(dossier.Actions_a_mettre_en_uvre_etapes || '').replace(/\n/g, '<br>')}</td>`;
            html += `<td contenteditable="false">${escapeHtml(dossier.Echeance ? formatDate(dossier.Echeance) : '')}</td>`;
            html += `<td contenteditable="false">${escapeHtml(formatDate(dossier.Date_de_la_reunion))}</td>`;
            html += `<td contenteditable="false">${escapeHtml(etatName)}</td>`;
            html += `<td><select class="etat-change-select"><option value="">-- Aucun changement --</option></select></td>`;
            html += `<td><input type="date" class="etat-change-date" disabled></td>`;
            html += `</tr>`;
        });

        html += '</tbody></table>';
        container.innerHTML = html;
        makeFieldsEditableReunion(container);
    } else {
        html = '<p class="loading">Aucun dossier échu non clôturé</p>';
        container.innerHTML = html;
    }
}

function clearReunionDisplay() {
    document.getElementById('reunion-odj-table').innerHTML = '';
    document.getElementById('reunion-echeance-table').innerHTML = '';
    document.getElementById('reunion-expired-table').innerHTML = '';

    const printContainer = document.getElementById('reunion-print-button-container');
    if (printContainer) {
        printContainer.classList.add('hidden');
    }
}

function makeFieldsEditableReunion(container) {
    // Récupérer la liste des porteurs
    const personnes = getUniqueValues(tablesData.Menus, 'Personnes').filter(p => p !== 'Autre');

    // Rendre les cellules éditables
    container.querySelectorAll('table tbody tr').forEach(row => {
        const cells = row.querySelectorAll('td');
        const dossierId = row.dataset.dossierId;
        const dossierData = tablesData.ODJ.find(d => d.id == dossierId);

        cells.forEach((td, index) => {
            // Index 0: Dossier (texte éditable)
            if (index === 0) {
                td.contentEditable = true;
                td.style.border = '1px solid #d9d9d9';
                td.style.backgroundColor = 'rgba(255, 255, 255, 0.5)';
                td.style.color = '#000';
            }
            // Index 1: Porteur(s) (select multiple)
            else if (index === 1 && dossierData) {
                const currentPorteurs = dossierData.Porteur_s_ || [];
                const select = document.createElement('select');
                select.multiple = true;
                select.style.width = '100%';
                select.style.minHeight = '60px';
                select.style.border = '1px solid #d9d9d9';
                select.style.backgroundColor = 'rgba(255, 255, 255, 0.9)';

                personnes.forEach(personne => {
                    const option = document.createElement('option');
                    option.value = personne;
                    option.textContent = personne;
                    const personneId = getPersonneIdByName(personne);
                    if (currentPorteurs.includes(personneId)) {
                        option.selected = true;
                    }
                    select.appendChild(option);
                });

                // Permettre la sélection/désélection sans maintenir Ctrl
                select.addEventListener('mousedown', function (e) {
                    e.preventDefault();
                    const option = e.target;
                    if (option.tagName === 'OPTION') {
                        option.selected = !option.selected;
                        select.focus();
                        // Réorganiser les options après la sélection
                        setTimeout(() => reorderSelectOptions(select), 10);
                    }
                });

                td.innerHTML = '';
                td.appendChild(select);
                td.style.padding = '4px';

                // Réorganiser les options au chargement initial
                reorderSelectOptions(select);
            }
            // Index 2: Actions (texte éditable avec sauts de ligne)
            else if (index === 2) {
                td.contentEditable = true;
                td.style.border = '1px solid #d9d9d9';
                td.style.backgroundColor = 'rgba(255, 255, 255, 0.5)';
                td.style.color = '#000';

                // Permettre les sauts de ligne avec Entrée (méthode sécurisée)
                td.addEventListener('keydown', function (event) {
                    if (event.key === 'Enter') {
                        event.preventDefault();
                        const selection = window.getSelection();
                        if (selection.rangeCount > 0) {
                            const range = selection.getRangeAt(0);
                            const br = document.createElement('br');
                            range.deleteContents();
                            range.insertNode(br);
                            range.setStartAfter(br);
                            range.setEndAfter(br);
                            selection.removeAllRanges();
                            selection.addRange(range);
                        }
                    }
                });
            }
            // Index 3: Échéance (date picker)
            else if (index === 3 && dossierData) {
                const dateInput = document.createElement('input');
                dateInput.type = 'date';
                dateInput.style.width = '100%';
                dateInput.style.border = '1px solid #d9d9d9';
                dateInput.style.backgroundColor = 'rgba(255, 255, 255, 0.9)';
                dateInput.style.padding = '4px';

                if (dossierData.Echeance) {
                    const date = new Date(dossierData.Echeance * 1000);
                    dateInput.value = date.toISOString().split('T')[0];
                }

                td.innerHTML = '';
                td.appendChild(dateInput);
                td.style.padding = '4px';
            }
        });
    });

    // Peupler les sélecteurs d'état
    const etats = getUniqueValues(tablesData.Menus, 'Etat');
    const ordreEtats = [
        "Clôturé",
        "Avance très bien",
        "Avance bien",
        "RAS",
        "Des tensions",
        "Forte difficulté, blocage",
        "Supprimer le dossier"
    ];

    const etatsTries = etats.sort((a, b) => {
        const indexA = ordreEtats.indexOf(a);
        const indexB = ordreEtats.indexOf(b);
        if (indexA !== -1 && indexB !== -1) return indexA - indexB;
        if (indexA !== -1) return -1;
        if (indexB !== -1) return 1;
        return a.localeCompare(b);
    });

    container.querySelectorAll('.etat-change-select').forEach(select => {
        etatsTries.forEach(etat => {
            const option = document.createElement('option');
            option.value = etat;
            option.textContent = etat;
            select.appendChild(option);
        });

        select.addEventListener('change', function () {
            const tr = this.closest('tr');
            const dateInput = tr.querySelector('.etat-change-date');

            if (this.value) {
                dateInput.disabled = false;
                dateInput.style.setProperty('background-color', '#ffffff', 'important');
                dateInput.style.color = '#262633';
                if (!dateInput.value) {
                    dateInput.value = new Date().toISOString().split('T')[0];
                }

                // Changer la couleur de la ligne selon le nouvel état sélectionné
                if (this.value !== 'Supprimer le dossier') {
                    const etatClass = etatColorMap[this.value] || '';
                    tr.className = etatClass;
                }
            } else {
                dateInput.disabled = true;
                dateInput.style.setProperty('background-color', '#a4a4a4', 'important');
                dateInput.value = '';

                // Restaurer la couleur d'origine
                const dossierId = tr.dataset.dossierId;
                const dossierData = tablesData.ODJ.find(d => d.id == dossierId);
                if (dossierData) {
                    const etatName = getEtatNameById(dossierData.Etat);
                    const etatClass = etatColorMap[etatName] || '';
                    tr.className = etatClass;
                }
            }
        });
    });

    // Initialiser la couleur de fond des champs de date désactivés
    container.querySelectorAll('.etat-change-date').forEach(dateInput => {
        dateInput.style.setProperty('background-color', '#a4a4a4', 'important');
        dateInput.style.color = '#3b3b3bff';
    });
}

async function saveReunionModifications() {
    try {
        const tables = ['reunion-odj-table', 'reunion-echeance-table', 'reunion-expired-table'];
        const updateActions = [];
        const newDates = new Set(); // Pour collecter les nouvelles dates à ajouter à l'Agenda

        for (const tableId of tables) {
            const container = document.getElementById(tableId);
            if (!container || !container.querySelector('table')) continue;

            const rows = container.querySelectorAll('table tbody tr');

            for (const row of rows) {
                const dossierId = Number.parseInt(row.dataset.dossierId);
                const dossierData = tablesData.ODJ.find(d => d.id === dossierId);
                if (!dossierData) continue;

                const cells = row.querySelectorAll('td');

                // Récupérer les valeurs modifiées - colonnes communes aux 3 tableaux
                const nouveauDossier = cells[0]?.textContent?.trim() || dossierData.Dossier;

                // Récupérer les porteurs sélectionnés
                const select = cells[1]?.querySelector('select');
                const nouveauxPorteurs = [];
                if (select) {
                    Array.from(select.selectedOptions).forEach(option => {
                        const id = getPersonneIdByName(option.value);
                        if (id) nouveauxPorteurs.push(id);
                    });
                } else {
                    // Si pas de select (ne devrait pas arriver), garder les porteurs existants
                    if (dossierData.Porteur_s_) {
                        nouveauxPorteurs.push(...dossierData.Porteur_s_);
                    }
                }

                let actions = cells[2]?.textContent?.trim() || dossierData.Actions_a_mettre_en_uvre_etapes;
                if (!actions && cells[2] && cells[2].innerHTML) {
                    actions = cells[2].innerHTML.replace(/<br\s*\/?>/gi, '\n').replace(/<[^>]+>/g, '').trim();
                }

                // Récupérer l'échéance selon le type de tableau
                // ODJ: [Dossier, Porteurs, Actions, Échéance, État, Change, Date] -> Échéance à index 3
                // Echeance: [Dossier, Porteurs, Actions, Date réunion, État, Change, Date] -> pas d'échéance éditable
                // Expired: [Dossier, Porteurs, Actions, Échéance, Date réunion, État, Change, Date] -> Échéance à index 3
                let nouvelleEcheance = dossierData.Echeance;

                if (tableId === 'reunion-odj-table' || tableId === 'reunion-expired-table') {
                    // Ces tableaux ont une colonne Échéance à l'index 3
                    const dateInput = cells[3]?.querySelector('input[type="date"]');
                    if (dateInput && dateInput.value) {
                        const dateObj = new Date(dateInput.value);
                        nouvelleEcheance = Math.floor(dateObj.getTime() / 1000);
                    }
                }
                // Pour reunion-echeance-table, on garde l'échéance existante car elle n'est pas affichée/éditable

                // Vérifier le changement d'état (toujours dans les dernières colonnes)
                const etatSelect = row.querySelector('.etat-change-select');
                const nouvelEtat = etatSelect ? etatSelect.value : '';
                const etatDateInput = row.querySelector('.etat-change-date');

                if (nouvelEtat === 'Supprimer le dossier') {
                    // Supprimer le dossier
                    updateActions.push(['RemoveRecord', 'ODJ', dossierId]);
                } else if (nouvelEtat && etatDateInput && etatDateInput.value) {
                    // Ajouter une nouvelle ligne avec le nouvel état et une date spécifique
                    const nouvelEtatId = getEtatIdByName(nouvelEtat);
                    const dateChangement = Math.floor(new Date(etatDateInput.value).getTime() / 1000);

                    // Collecter cette date pour l'ajouter à l'Agenda (sécurité OWASP)
                    newDates.add(dateChangement);

                    updateActions.push(['AddRecord', 'ODJ', null, {
                        Date_de_la_reunion: dateChangement,
                        Dossier: nouveauDossier || dossierData.Dossier,
                        Porteur_s_: ['L', ...nouveauxPorteurs],
                        Actions_a_mettre_en_uvre_etapes: actions,
                        Echeance: nouvelleEcheance,
                        Etat: nouvelEtatId,
                        Enregistrement: Date.now() / 1000
                    }]);
                }

                // Toujours mettre à jour la ligne existante
                updateActions.push(['UpdateRecord', 'ODJ', dossierId, {
                    Dossier: nouveauDossier,
                    Porteur_s_: ['L', ...nouveauxPorteurs],
                    Actions_a_mettre_en_uvre_etapes: actions,
                    Echeance: nouvelleEcheance
                }]);
            }
        }

        // S'assurer que toutes les nouvelles dates existent dans l'Agenda avant d'appliquer les actions
        for (const date of newDates) {
            await ensureAgendaDateExists(date);
        }

        // Appliquer toutes les modifications en une seule action
        if (updateActions.length > 0) {
            await grist.docApi.applyUserActions(updateActions);
        }

        alert('Modifications enregistrées avec succès !');
        await loadAllTables();
        await removeDuplicateRecords();
        await loadAllTables();
        populateReunionDateSelect();
        populateConsultSelectors();

        // Rafraîchir l'affichage avec la date sélectionnée
        const select = document.getElementById('reunion-date-select');
        if (select && select.value) {
            reunionDisplayData();
        }

    } catch (error) {
        console.error('Erreur lors de la sauvegarde:', error);
        alert('Erreur lors de l\'enregistrement des modifications: ' + error.message);
    }
}

// ========================================
// SECTIONS REPLIABLES
// ========================================

// ========================================
// DÉMARRAGE
// ========================================

if (typeof grist !== 'undefined') {
    initWidget();
} else {
    console.error('Grist API non disponible');
    document.body.innerHTML = '<div class="container"><p class="loading">Erreur : Widget doit être utilisé dans Grist</p></div>';
}