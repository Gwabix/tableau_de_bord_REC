// Configuration Grist - Version simplifiée
let tablesData = {
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

// ========================================
// CHOIX DE COLONNES (Porteur_s_ / Etat)
// ========================================
// Depuis l'abandon de la table « Menus », les porteurs et les états sont des
// colonnes Choice/ChoiceList : chaque ligne d'ODJ stocke le texte, pas une
// référence. La liste des choix, leur ordre et leurs couleurs sont lus dans la
// configuration des colonnes (widgetOptions) ; ces constantes ne servent que de
// repli si cette config est illisible.

const DEFAULT_ETAT_ORDER = [
    "Clôturé",
    "Avance très bien",
    "Avance bien",
    "RAS",
    "Des tensions",
    "Forte difficulté, blocage",
    "Supprimer le dossier"
];

const DEFAULT_ETAT_COLORS = {
    "Clôturé": "#4A90E2",
    "Avance très bien": "#479415",
    "Avance bien": "#82b34f",
    "RAS": "#FDD835",
    "Des tensions": "#FF8A80",
    "Forte difficulté, blocage": "#D32F2F"
};

const DEFAULT_ETAT_ROLES = { cloture: "Clôturé", supprimer: "Supprimer le dossier" };

// État module, (re)peuplé par loadColumnChoices() à chaque chargement des tables.
let columnChoices = {
    etats: [...DEFAULT_ETAT_ORDER],
    porteurs: [],
    etatColors: { ...DEFAULT_ETAT_COLORS }
};
let etatRoles = { ...DEFAULT_ETAT_ROLES };

/** Liste des états dans l'ordre d'affichage (sélecteurs). */
function getEtatDropdownOrder() {
    return columnChoices.etats.slice();
}

/** Liste des porteurs proposés à la saisie (choix configurés). */
function getPorteurChoices() {
    return columnChoices.porteurs.slice();
}

/**
 * Ordre de tri des dossiers par état, du pire au meilleur : ordre d'affichage
 * inversé, sans l'état « à supprimer ».
 */
function getEtatSortOrder() {
    return columnChoices.etats.filter(e => e !== etatRoles.supprimer).reverse();
}

/** Noms des porteurs d'un dossier (retire le sentinelle 'L' de la ChoiceList). */
function getDossierPorteurs(dossier) {
    return Array.isArray(dossier.Porteur_s_) ? dossier.Porteur_s_.filter(v => v !== 'L') : [];
}

/** Couleur de texte lisible (noir/blanc) sur un fond hexadécimal donné. */
function contrastText(hex) {
    const match = /^#?([0-9a-f]{6})$/i.exec(hex || '');
    if (!match) return '#262633';
    const n = Number.parseInt(match[1], 16);
    const r = (n >> 16) & 255;
    const g = (n >> 8) & 255;
    const b = n & 255;
    const luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
    return luminance > 0.6 ? '#262633' : '#ffffff';
}

/** Fragment `style="…"` de coloration d'un état (chaîne vide si pas de couleur). */
function etatStyleAttr(etatName) {
    const bg = columnChoices.etatColors[etatName];
    if (!bg || !/^#[0-9a-f]{6}$/i.test(bg)) return '';
    return ` style="background-color:${bg};color:${contrastText(bg)}"`;
}

/** Applique (ou retire) la couleur d'un état sur un élément (ligne de tableau). */
function applyEtatStyle(el, etatName) {
    const bg = columnChoices.etatColors[etatName];
    if (bg && /^#[0-9a-f]{6}$/i.test(bg)) {
        el.style.backgroundColor = bg;
        el.style.color = contrastText(bg);
    } else {
        clearEtatStyle(el);
    }
}

function clearEtatStyle(el) {
    el.style.backgroundColor = '';
    el.style.color = '';
}

function deriveDefaultEtatRoles(etats) {
    const has = name => etats.includes(name);
    return {
        cloture: has(DEFAULT_ETAT_ROLES.cloture) ? DEFAULT_ETAT_ROLES.cloture : (etats[0] || DEFAULT_ETAT_ROLES.cloture),
        supprimer: has(DEFAULT_ETAT_ROLES.supprimer) ? DEFAULT_ETAT_ROLES.supprimer : (etats[etats.length - 1] || DEFAULT_ETAT_ROLES.supprimer)
    };
}

async function readWidgetOption(name) {
    try {
        if (typeof grist.getOption === 'function') return await grist.getOption(name);
        if (grist.widgetApi && typeof grist.widgetApi.getOption === 'function') {
            return await grist.widgetApi.getOption(name);
        }
    } catch (error) {
        console.warn('Lecture option widget impossible:', name, error);
    }
    return undefined;
}

/**
 * Lit la liste des choix (noms, ordre, couleurs) des colonnes Etat / Porteur_s_
 * via les tables de métadonnées Grist, plus les rôles spéciaux des états
 * (option de widget). Repli sur les valeurs par défaut en cas d'échec.
 */
async function loadColumnChoices() {
    try {
        const [metaTables, metaColumns] = await Promise.all([
            grist.docApi.fetchTable('_grist_Tables'),
            grist.docApi.fetchTable('_grist_Tables_column')
        ]);

        const odjRowId = metaTables.id[metaTables.tableId.indexOf('ODJ')];

        const optionsForColumn = colId => {
            for (let i = 0; i < metaColumns.id.length; i++) {
                if (metaColumns.parentId[i] === odjRowId && metaColumns.colId[i] === colId) {
                    try {
                        return JSON.parse(metaColumns.widgetOptions[i] || '{}') || {};
                    } catch {
                        return {};
                    }
                }
            }
            return {};
        };

        const etatOpts = optionsForColumn('Etat');
        const porteurOpts = optionsForColumn('Porteur_s_');

        const etats = Array.isArray(etatOpts.choices) && etatOpts.choices.length
            ? etatOpts.choices.slice()
            : [...DEFAULT_ETAT_ORDER];
        const porteurs = Array.isArray(porteurOpts.choices) ? porteurOpts.choices.slice() : [];

        const etatColors = {};
        etats.forEach(nom => {
            const choiceOpt = etatOpts.choiceOptions && etatOpts.choiceOptions[nom];
            const fill = choiceOpt && choiceOpt.fillColor;
            etatColors[nom] = (fill && /^#[0-9a-f]{6}$/i.test(fill) ? fill : null)
                || DEFAULT_ETAT_COLORS[nom]
                || '';
        });

        columnChoices = { etats, porteurs, etatColors };
    } catch (error) {
        console.warn('Configuration des colonnes illisible, valeurs par défaut utilisées.', error);
        columnChoices = {
            etats: [...DEFAULT_ETAT_ORDER],
            porteurs: [],
            etatColors: { ...DEFAULT_ETAT_COLORS }
        };
    }

    const storedRoles = await readWidgetOption('etatRoles');
    etatRoles = (storedRoles && storedRoles.cloture && storedRoles.supprimer)
        ? { cloture: storedRoles.cloture, supprimer: storedRoles.supprimer }
        : deriveDefaultEtatRoles(columnChoices.etats);
}

// Contexte pour réouverture du formulaire après modification
let modifyContext = {
    type: null,
    value: null,
    secondValue: null
};

let consultPorteurSortState = {
    key: 'Date_de_la_reunion',
    direction: 'desc'
};

// ========================================
// INITIALISATION
// ========================================

let widgetInitialized = false;
let widgetInitializing = false;

function initWidget() {
    grist.ready({
        requiredAccess: 'full'
    });

    // grist.onRecord peut se déclencher à chaque déplacement de curseur dans
    // Grist. On n'initialise (et on ne reconstruit les formulaires) qu'une
    // seule fois : sinon une saisie en cours dans « Saisir » serait écrasée.
    // Les données se rafraîchissent au changement d'onglet et après chaque
    // enregistrement.
    grist.onRecord(async function () {
        if (widgetInitialized || widgetInitializing) return;
        widgetInitializing = true;
        try {
            if (!await loadAllTables()) return; // réessai au prochain onRecord
            initializeUI();
            attachEventListeners();
            widgetInitialized = true;
        } finally {
            widgetInitializing = false;
        }
    });
}

async function loadAllTables() {
    try {
        const docApi = grist.docApi;

        const [odjTable, agendaTable] = await Promise.all([
            docApi.fetchTable('ODJ'),
            docApi.fetchTable('Agenda')
        ]);

        // Liste des porteurs / états (noms, ordre, couleurs) + rôles spéciaux.
        await loadColumnChoices();

        tablesData.ODJ = odjTable.id.map((id, index) => ({
            id: id,
            Date_de_la_reunion: odjTable.Date_de_la_reunion[index],
            Dossier: odjTable.Dossier[index] || '',
            ID_Dossier: odjTable.ID_Dossier[index] || '',
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
        return true;
    } catch (error) {
        console.error('Erreur lors du chargement des tables:', error);
        alert('Erreur lors du chargement des données. Vérifiez les noms des tables.');
        return false;
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
    const personnes = getPorteurChoices();

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
    const etats = getEtatDropdownOrder();

    selects.forEach(select => {
        select.innerHTML = '<option value="">-- Sélectionner --</option>';
        etats.forEach(etat => {
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
    buildEtatFilterCheckboxes(document.getElementById('filter-etat-checkboxes'), 'filter-etat');
}

function getUpcomingMeetingDates() {
    if (!tablesData.Agenda || tablesData.Agenda.length === 0) {
        return [];
    }

    const todayTs = todayCalendarTs();
    return tablesData.Agenda
        .map(item => item.Date)
        .filter(date => date >= todayTs)
        .sort((a, b) => a - b);
}

function getNextMeetingDate() {
    const futureDates = getUpcomingMeetingDates();
    return futureDates.length > 0 ? futureDates[0] : null;
}

function populateUpcomingMeetingsSelect() {
    const select = document.getElementById('saisir-reunions-suivantes');
    if (!select) return;

    // Exclure la date déjà affichée dans le champ « Date de la réunion »
    // (sinon la prochaine réunion apparaîtrait à la fois dans le champ et ici).
    const dateInput = document.getElementById('saisir-date');
    let currentTimestamp = null;
    if (dateInput && dateInput.value) {
        const t = Math.floor(new Date(`${dateInput.value}T00:00:00Z`).getTime() / 1000);
        if (Number.isFinite(t)) currentTimestamp = t;
    }

    const upcomingDates = getUpcomingMeetingDates().filter(date => date !== currentTimestamp);

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
        dateInput.value = nextMeetingDate
            ? tsToInputDate(nextMeetingDate)
            : tsToInputDate(todayCalendarTs());
    }

    // Peupler le menu déroulant des réunions suivantes
    populateUpcomingMeetingsSelect();
}

// ========================================
// UTILITAIRES
// ========================================

/**
 * (Re)construit un groupe de cases à cocher pour filtrer par état, avec une
 * case « Non renseigné » (value '') pour les dossiers sans état saisi.
 * @param {HTMLElement} container
 * @param {string} name - attribut name commun des cases
 */
function buildEtatFilterCheckboxes(container, name) {
    if (!container) return;

    const etats = getEtatDropdownOrder().filter(e => e !== etatRoles.supprimer);
    container.innerHTML = '';

    const addCheckbox = (value, labelText) => {
        const label = document.createElement('label');
        label.className = 'checkbox-label';

        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.name = name;
        checkbox.value = value;
        checkbox.checked = true;

        const span = document.createElement('span');
        span.textContent = labelText;

        label.appendChild(checkbox);
        label.appendChild(span);
        container.appendChild(label);
    };

    etats.forEach(etat => addCheckbox(etat, etat));
    addCheckbox('', 'État non renseigné');
}

function generateDossierID() {
    const now = new Date();
    const yyyy = now.getFullYear();
    const mm = String(now.getMonth() + 1).padStart(2, '0');
    const dd = String(now.getDate()).padStart(2, '0');
    const hh = String(now.getHours()).padStart(2, '0');
    const mn = String(now.getMinutes()).padStart(2, '0');
    const ss = String(now.getSeconds()).padStart(2, '0');
    const ms = String(now.getMilliseconds()).padStart(3, '0');
    return `${yyyy}${mm}${dd}${hh}${mn}${ss}${ms}`;
}

/**
 * Cherche un record du même dossier avec la même Date_de_la_reunion créé aujourd'hui.
 * Permet un "upsert par jour" : une seule nouvelle ligne par dossier par jour.
 */
function findTodayRecord(dossier, targetDateReunion) {
    const now = new Date();
    const todayStart = Math.floor(new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime() / 1000);
    const todayEnd = todayStart + 86399;
    const idKey = dossier.ID_Dossier || null;
    return tablesData.ODJ.find(r => {
        const sameId = idKey ? r.ID_Dossier === idKey : r.Dossier === dossier.Dossier;
        if (!sameId) return false;
        if (r.Date_de_la_reunion !== targetDateReunion) return false;
        const enr = r.Enregistrement || 0;
        return enr >= todayStart && enr <= todayEnd;
    });
}

function sortByEtat(dossiers) {
    const etatSortOrder = getEtatSortOrder();
    return dossiers.sort((a, b) => {
        const etatA = a.Etat || '';
        const etatB = b.Etat || '';

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
        const key = dossier.ID_Dossier || dossier.Dossier;
        const existing = dossierMap.get(key);

        if (!existing) {
            dossierMap.set(key, dossier);
        } else {
            // Comparer les timestamps d'enregistrement et garder le plus récent
            const existingTimestamp = existing.Enregistrement || 0;
            const currentTimestamp = dossier.Enregistrement || 0;

            if (currentTimestamp > existingTimestamp) {
                dossierMap.set(key, dossier);
            }
        }
    });

    return Array.from(dossierMap.values());
}

function getUniqueDates(data, column) {
    const dates = data
        .map(row => row[column])
        .filter(val => val);
    // Timestamps Grist (secondes) : tri numérique décroissant (plus récent d'abord).
    return [...new Set(dates)].sort((a, b) => b - a);
}

/**
 * Noms de dossiers uniques, classés par activité la plus récente d'abord
 * (valeur maximale de la colonne Enregistrement parmi les lignes du dossier).
 * Sert à ordonner les suggestions du champ « Rechercher un dossier ».
 */
function getDossiersByRecentActivity() {
    const lastActivity = new Map();

    tablesData.ODJ.forEach(odj => {
        if (!odj.Dossier) return;
        const ts = odj.Enregistrement || 0;
        const prev = lastActivity.get(odj.Dossier);
        if (prev === undefined || ts > prev) {
            lastActivity.set(odj.Dossier, ts);
        }
    });

    return [...lastActivity.entries()]
        .sort((a, b) => b[1] - a[1])
        .map(([nom]) => nom);
}

/**
 * Porteurs présents dans l'historique ODJ (pas seulement les choix courants) :
 * on doit pouvoir consulter/modifier les dossiers d'un porteur retiré de la liste.
 */
function getUniquePorteurs() {
    const porteurs = new Set();
    tablesData.ODJ.forEach(row => {
        getDossierPorteurs(row).forEach(name => {
            if (name) porteurs.add(name);
        });
    });
    return [...porteurs].sort();
}

// Les colonnes Date de Grist stockent un timestamp « minuit UTC » représentant
// un jour calendaire. On affiche donc toujours en UTC pour que la date reste la
// même quel que soit le fuseau du navigateur (métropole, DOM-TOM…).
function formatDate(dateString) {
    if (!dateString) return '';
    const date = typeof dateString === 'number' ? new Date(dateString * 1000) : new Date(dateString);
    if (Number.isNaN(date.getTime())) return '';
    return date.toLocaleDateString('fr-FR', {
        timeZone: 'UTC',
        year: 'numeric',
        month: 'long',
        day: 'numeric'
    });
}

/** Format compact « JJ/MM/AAAA » pour les cellules de tableau. */
function formatDateShort(value) {
    if (!value && value !== 0) return '';
    const date = typeof value === 'number' ? new Date(value * 1000) : new Date(value);
    if (Number.isNaN(date.getTime())) return '';
    return date.toLocaleDateString('fr-FR', {
        timeZone: 'UTC',
        day: '2-digit',
        month: '2-digit',
        year: 'numeric'
    });
}

/**
 * Timestamp (secondes) de « minuit UTC » du jour calendaire local courant,
 * comparable directement aux dates stockées dans Grist.
 */
function todayCalendarTs() {
    const now = new Date();
    return Date.UTC(now.getFullYear(), now.getMonth(), now.getDate()) / 1000;
}

/** Convertit un timestamp Grist (secondes, minuit UTC) en « AAAA-MM-JJ ». */
function tsToInputDate(ts) {
    if (!ts && ts !== 0) return '';
    const date = new Date(ts * 1000);
    return Number.isNaN(date.getTime()) ? '' : date.toISOString().split('T')[0];
}

// ========================================
// SUPPRESSION DE DOSSIERS
// ========================================

/**
 * Retourne les IDs de tous les enregistrements ODJ appartenant au même dossier.
 * Identité : ID_Dossier s'il est présent, sinon nom exact du dossier (pour les
 * anciens enregistrements créés avant l'introduction de ID_Dossier).
 * @param {object} dossier - Un enregistrement ODJ représentatif du dossier
 * @returns {Array<number>}
 */
function getDossierRecordIds(dossier) {
    const ids = new Set();
    if (dossier && dossier.id !== null && dossier.id !== undefined) {
        ids.add(dossier.id);
    }

    const idKey = dossier && dossier.ID_Dossier ? dossier.ID_Dossier : null;

    tablesData.ODJ.forEach(record => {
        const sameDossier = idKey
            ? record.ID_Dossier === idKey
            : (!record.ID_Dossier && record.Dossier === dossier.Dossier);
        if (sameDossier) ids.add(record.id);
    });

    return [...ids];
}

/**
 * Demande une confirmation unique puis supprime intégralement les dossiers
 * marqués « Supprimer le dossier » (toutes leurs lignes d'historique).
 * @param {Array<{dossier: object, nom: string}>} dossiersASupprimer
 * @returns {Promise<{status: 'none'|'cancelled'|'deleted', deletedIds: Set<number>}>}
 */
async function confirmAndDeleteDossiers(dossiersASupprimer) {
    if (!dossiersASupprimer || dossiersASupprimer.length === 0) {
        return { status: 'none', deletedIds: new Set() };
    }

    const noms = [...new Set(
        dossiersASupprimer.map(d => (d.nom && d.nom.trim()) || '(sans nom)')
    )];
    const liste = noms.map(nom => `- ${nom}`).join('\n');
    const message = `ATTENTION !\n\n`
        + `Les dossiers suivants vont être supprimés définitivement, `
        + `avec l'intégralité de leur historique :\n\n${liste}\n\n`
        + `Confirmer la suppression ?`;

    if (!confirm(message)) {
        return { status: 'cancelled', deletedIds: new Set() };
    }

    const deletedIds = new Set();
    dossiersASupprimer.forEach(({ dossier }) => {
        getDossierRecordIds(dossier).forEach(id => deletedIds.add(id));
    });

    if (deletedIds.size > 0) {
        await grist.docApi.applyUserActions(
            [...deletedIds].map(id => ['RemoveRecord', 'ODJ', id])
        );
    }

    return { status: 'deleted', deletedIds };
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

let eventListenersAttached = false;

function attachEventListeners() {
    // Les éléments statiques ne doivent recevoir leurs écouteurs qu'une seule fois,
    // même si grist.onRecord relance l'initialisation à chaque changement de curseur.
    if (eventListenersAttached) return;
    eventListenersAttached = true;

    // Onglets
    document.querySelectorAll('.tab-button').forEach(button => {
        button.addEventListener('click', switchTab);
    });

    // Multi-select (porteurs, onglet Saisir) : les options sont recréées
    // dynamiquement par populatePorteurs() / addDossier(), on écoute donc sur
    // le conteneur stable par délégation.
    const saisirDossiers = document.getElementById('saisir-dossiers');
    if (saisirDossiers) {
        saisirDossiers.addEventListener('click', function (event) {
            if (event.target.classList.contains('multi-select-option')) {
                toggleMultiSelect(event);
            }
        });
    }

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

    // Filtres d'état (consultation par porteur) : les cases sont recréées
    // dynamiquement par populateConsultSelectors(), on écoute donc sur le
    // conteneur stable par délégation.
    const filterEtatContainer = document.getElementById('filter-etat-checkboxes');
    if (filterEtatContainer) {
        filterEtatContainer.addEventListener('change', function (event) {
            if (event.target.matches('input[name="filter-etat"]')) {
                consultByPorteur();
            }
        });
    }

    // Réunion
    const reunionDateSelect = document.getElementById('reunion-date-select');
    if (reunionDateSelect) {
        reunionDateSelect.addEventListener('change', async () => {
            // Enregistrer une modification en attente avant de recharger la vue
            await flushReunionAutoSave();
            reunionDisplayData();
        });
        // Afficher les données pour la date par défaut
        if (reunionDateSelect.value) {
            reunionDisplayData();
        }
    }

    const tabReunion = document.getElementById('tab-reunion');
    if (tabReunion) {
        tabReunion.addEventListener('focusout', handleReunionAutoSaveEvent);
        tabReunion.addEventListener('change', handleReunionAutoSaveEvent);
        tabReunion.addEventListener('input', handleReunionAutoSaveEvent);
    }

    const tabModifier = document.getElementById('tab-modifier');
    if (tabModifier) {
        tabModifier.addEventListener('focusout', handleModifyAutoSaveEvent);
        tabModifier.addEventListener('change', handleModifyAutoSaveEvent);
        tabModifier.addEventListener('input', handleModifyAutoSaveEvent);
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
        modifyDateSelect.addEventListener('change', withModifyFlush(modifyByDate));
    }

    let modifyDossierInput = document.getElementById('modify-dossier-input');
    if (modifyDossierInput) {
        attachDossierAutocomplete(modifyDossierInput, 'modify');
    }

    const modifyPorteurSelect = document.getElementById('modify-porteur-select');
    if (modifyPorteurSelect) {
        modifyPorteurSelect.addEventListener('change', withModifyFlush(handleModifyPorteurSelectChange));
    }

    const modifyPorteurDossierSelect = document.getElementById('modify-porteur-dossier-select');
    if (modifyPorteurDossierSelect) {
        modifyPorteurDossierSelect.addEventListener('change', handleModifyPorteurDossierSelectChange);
    }

    const modifyHideExpired = document.getElementById('modify-hide-expired');
    if (modifyHideExpired) {
        modifyHideExpired.addEventListener('change', handleModifyPorteurDossierSelectChange);
    }

    // Filtres d'état (modification par porteur) : cases recréées dynamiquement,
    // délégation sur le conteneur stable.
    const modifyFilterEtatContainer = document.getElementById('modify-filter-etat-checkboxes');
    if (modifyFilterEtatContainer) {
        modifyFilterEtatContainer.addEventListener('change', function (event) {
            if (event.target.matches('input[name="modify-filter-etat"]')) {
                handleModifyPorteurDossierSelectChange();
            }
        });
    }

    const modifyEcheanceSelect = document.getElementById('modify-echeance-select');
    if (modifyEcheanceSelect) {
        modifyEcheanceSelect.addEventListener('change', withModifyFlush(modifyByEcheance));
    }

    const btnCloseModif = document.getElementById('btn-close-modifications');
    if (btnCloseModif) {
        btnCloseModif.addEventListener('click', closeModifyForm);
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

    // Champ « Date de la réunion » : tenir à jour la liste des autres dates
    const saisirDateInput = document.getElementById('saisir-date');
    if (saisirDateInput) {
        saisirDateInput.addEventListener('change', populateUpcomingMeetingsSelect);
    }

    // Menu déroulant des réunions suivantes
    const reunionsSuivantesSelect = document.getElementById('saisir-reunions-suivantes');
    if (reunionsSuivantesSelect) {
        reunionsSuivantesSelect.addEventListener('change', function () {
            if (this.value) {
                const dateInput = document.getElementById('saisir-date');
                if (dateInput) {
                    const timestamp = Number.parseFloat(this.value);
                    if (Number.isFinite(timestamp)) {
                        const iso = tsToInputDate(timestamp);
                        if (iso) dateInput.value = iso;
                    }
                }
                // Réinitialiser le menu déroulant et exclure la date choisie
                this.value = '';
                populateUpcomingMeetingsSelect();
            }
        });
    }
}

async function switchTab(event) {
    const targetTab = event.target.dataset.tab;

    // Enregistrer une éventuelle modification en attente avant de quitter l'onglet
    await flushModifyAutoSave();
    await flushReunionAutoSave();

    document.querySelectorAll('.tab-button').forEach(btn => btn.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(content => content.classList.remove('active'));

    event.target.classList.add('active');
    const targetContent = document.getElementById(`tab-${targetTab}`);
    if (targetContent) {
        targetContent.classList.add('active');
    }

    // Vider les sélections lors du changement d'onglet
    clearAllSelections();

    // Recharger les données pour refléter les ajouts / modifications récents
    // (ex. un dossier saisi doit apparaître immédiatement dans « Réunion »).
    await refreshActiveTabData(targetTab);
}

/**
 * Recharge les tables Grist puis rafraîchit les listes et affichages de
 * l'onglet devenu actif.
 * @param {string} targetTab - saisir | reunion | consulter | modifier
 */
async function refreshActiveTabData(targetTab) {
    await loadAllTables();

    // Listes communes (dates de réunion, échéances, porteurs, filtres d'état)
    populateConsultSelectors();

    if (targetTab === 'saisir') {
        // initializeUI() ne s'exécute qu'une fois : on rafraîchit ici les listes
        // dérivées de Menus, mais seulement si la saisie n'a pas commencé (pour
        // ne pas effacer des sélections en cours).
        if (saisirFormIsPristine()) {
            populatePorteurs();
            populateEtats();
        }
        populateUpcomingMeetingsSelect();
    } else if (targetTab === 'reunion') {
        refreshReunionView();
    }
}

/** true si aucun dossier de l'onglet Saisir n'a été renseigné. */
function saisirFormIsPristine() {
    const container = document.getElementById('saisir-dossiers');
    if (!container) return true;

    return [...container.querySelectorAll('.dossier-block')].every(block => {
        const intitule = block.querySelector('.dossier-intitule')?.value.trim();
        const actions = block.querySelector('.dossier-actions')?.value.trim();
        const echeance = block.querySelector('.dossier-echeance')?.value;
        const etat = block.querySelector('.dossier-etat')?.value;
        const porteur = block.querySelector('.dossier-porteurs .multi-select-option.selected');
        return !intitule && !actions && !echeance && !etat && !porteur;
    });
}

/**
 * Repeuple le sélecteur de date de réunion en conservant la date choisie si
 * elle existe toujours, puis réaffiche les tableaux de l'onglet Réunion.
 * reunionDisplayData() vide l'affichage de lui-même si aucune date n'est
 * sélectionnée.
 */
function refreshReunionView() {
    const select = document.getElementById('reunion-date-select');
    const previousValue = select ? select.value : '';

    populateReunionDateSelect();

    if (!select) return;

    if (previousValue && [...select.options].some(opt => opt.value === previousValue)) {
        select.value = previousValue;
    }
    reunionDisplayData();
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

    const byName = (a, b) => (a.dataset.value || '')
        .localeCompare(b.dataset.value || '', 'fr', { sensitivity: 'base' });

    const options = Array.from(container.querySelectorAll('.multi-select-option'));
    // Sélectionnés en haut, non-sélectionnés en dessous ; chaque groupe reste
    // trié par ordre alphabétique (une option désélectionnée retrouve sa place).
    const selected = options.filter(opt => opt.classList.contains('selected')).sort(byName);
    const notSelected = options.filter(opt => !opt.classList.contains('selected')).sort(byName);

    container.innerHTML = '';
    [...selected, ...notSelected].forEach(option => container.appendChild(option));
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

    // Les clics sur .multi-select-option sont gérés par délégation sur
    // #saisir-dossiers (voir attachEventListeners), rien à attacher ici.

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
            currentSuggestions = getDossiersByRecentActivity().slice(0, 10);
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
            currentSuggestions = getDossiersByRecentActivity().slice(0, 10);
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
            const chosen = currentSuggestions[highlightedIndex];
            input.value = chosen;
            suggestionsDiv.classList.remove('visible');
            // Mettre à jour le bouton clear après sélection
            toggleClearButton(clearButtonId, input.value);
            if (mode === 'consult') {
                consultByDossier(chosen);
            } else if (mode === 'modify') {
                flushModifyAutoSave().then(() => modifyByDossier(chosen));
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
                flushModifyAutoSave().then(() => modifyByDossier(dossierName));
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
                .filter(Boolean);

            if (intitule) {
                const odjData = {
                    Date_de_la_reunion: dateReunionTimestamp,
                    Dossier: intitule,
                    ID_Dossier: generateDossierID(),
                    Porteur_s_: ['L', ...porteurs],
                    Actions_a_mettre_en_uvre_etapes: actions,
                    Echeance: echeance || null,
                    Etat: etat || null,
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

    // Ne garder que le dernier état de chaque dossier (le plus récent)
    dossiers = getLatestEntriesPerDossier(dossiers);

    // Trier par état (du pire au meilleur)
    dossiers = sortByEtat(dossiers);

    // Vider le conteneur
    resultsDiv.innerHTML = '';

    if (dossiers.length === 0) {
        const p = document.createElement('p');
        p.className = 'no-results';
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
        const etatName = dossier.Etat || '';
        const porteurs = getDossierPorteurs(dossier).join(', ');

        const tr = document.createElement('tr');
        applyEtatStyle(tr, etatName);
        makeDossierRowClickable(tr, dossier.Dossier);

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
        tdEcheance.textContent = formatDateShort(dossier.Echeance);
        tr.appendChild(tdEcheance);

        const tdEtat = document.createElement('td');
        tdEtat.textContent = etatName;
        tr.appendChild(tdEtat);

        tbody.appendChild(tr);
    });

    table.appendChild(tbody);
    tagConsultTableColumns(table);
    tableContainer.appendChild(table);
    section.appendChild(tableContainer);
    resultsDiv.appendChild(section);

    togglePrintButton();
}

function consultByDossier(dossierName) {
    const resultsDiv = document.getElementById('consult-results');
    if (!resultsDiv) return;

    // Résoudre l'ID_Dossier à partir du nom saisi, puis inclure tous les records partageant cet ID
    const refRecord = tablesData.ODJ.find(odj => odj.Dossier === dossierName && odj.ID_Dossier);
    const dossiers = refRecord
        ? tablesData.ODJ.filter(odj => odj.ID_Dossier === refRecord.ID_Dossier)
        : tablesData.ODJ.filter(odj => odj.Dossier === dossierName);

    resultsDiv.innerHTML = '';

    if (dossiers.length === 0) {
        const p = document.createElement('p');
        p.className = 'no-results';
        p.textContent = 'Aucun dossier trouvé';
        resultsDiv.appendChild(p);
        togglePrintButton();
        return;
    }

    dossiers.sort((a, b) => (b.Enregistrement || 0) - (a.Enregistrement || 0));

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
    ['Date réunion', 'Porteur(s)', 'Actions', 'Échéance', 'État', 'Date d\'enregistrement'].forEach(headerText => {
        const th = document.createElement('th');
        th.textContent = headerText;
        headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);
    table.appendChild(thead);

    const tbody = document.createElement('tbody');

    dossiers.forEach(dossier => {
        const etatName = dossier.Etat || '';
        const porteurs = getDossierPorteurs(dossier).join(', ');

        const tr = document.createElement('tr');
        applyEtatStyle(tr, etatName);

        const tdDate = document.createElement('td');
        tdDate.textContent = formatDateShort(dossier.Date_de_la_reunion);
        tr.appendChild(tdDate);

        const tdPorteurs = document.createElement('td');
        tdPorteurs.textContent = porteurs;
        tr.appendChild(tdPorteurs);

        const tdActions = document.createElement('td');
        tdActions.textContent = dossier.Actions_a_mettre_en_uvre_etapes || '';
        tdActions.style.whiteSpace = 'pre-wrap';
        tr.appendChild(tdActions);

        const tdEcheance = document.createElement('td');
        tdEcheance.textContent = formatDateShort(dossier.Echeance);
        tr.appendChild(tdEcheance);

        const tdEtat = document.createElement('td');
        tdEtat.textContent = etatName;
        tr.appendChild(tdEtat);

        const tdEnregistrement = document.createElement('td');
        tdEnregistrement.textContent = formatDateShort(dossier.Enregistrement);
        tr.appendChild(tdEnregistrement);

        tbody.appendChild(tr);
    });

    table.appendChild(tbody);
    tagConsultTableColumns(table);
    tableContainer.appendChild(table);
    section.appendChild(tableContainer);
    resultsDiv.appendChild(section);

    togglePrintButton();
}

/**
 * Navigue vers la consultation "par dossier" en sélectionnant le bon radio
 * et en affichant l'historique complet du dossier demandé.
 */
function navigateToConsultDossier(dossierName) {
    // Activer le radio "par dossier"
    const radioDossier = document.querySelector('input[name="consult-type"][value="dossier"]');
    if (radioDossier) {
        radioDossier.checked = true;
        radioDossier.dispatchEvent(new Event('change'));
    }

    // Remplir le champ de saisie
    const input = document.getElementById('consult-dossier-input');
    if (input) {
        input.value = dossierName;
        toggleClearButton('btn-clear-consult-dossier', dossierName);
    }

    // Lancer la consultation
    consultByDossier(dossierName);
}

/**
 * Rend une ligne de tableau cliquable : un clic (souris ou clavier) ouvre
 * l'historique complet du dossier (consultation « par dossier »). Utilisé par
 * les consultations par porteur, par date de réunion et par date d'échéance.
 */
function makeDossierRowClickable(tr, dossierName) {
    const label = `Voir l'historique complet du dossier ${dossierName || ''}`.trim();

    tr.classList.add('tr-dossier-link');
    tr.title = 'Voir l\'historique complet de ce dossier';

    // Accessibilité clavier : la ligne devient un contrôle focusable activable
    // par Entrée ou Espace, comme un bouton.
    tr.setAttribute('role', 'button');
    tr.setAttribute('tabindex', '0');
    tr.setAttribute('aria-label', label);

    tr.addEventListener('click', () => navigateToConsultDossier(dossierName));
    tr.addEventListener('keydown', event => {
        if (event.key === 'Enter' || event.key === ' ') {
            event.preventDefault();
            navigateToConsultDossier(dossierName);
        }
    });
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

    const selectedEtats = new Set(
        Array.from(document.querySelectorAll('input[name="filter-etat"]:checked'))
            .map(cb => cb.value)
    );

    let dossiers = tablesData.ODJ.filter(odj => {
        if (!Array.isArray(odj.Porteur_s_)) return false;
        return getDossierPorteurs(odj).includes(porteurName);
    });

    // Garder uniquement le record le plus récent par dossier (avant de filtrer par état)
    dossiers = getLatestEntriesPerDossier(dossiers);

    // Filtrer par état sur le vrai état courant du dossier
    dossiers = dossiers.filter(dossier => {
        const etatName = dossier.Etat || '';
        return selectedEtats.has(etatName);
    });

    dossiers = sortConsultPorteurDossiers(dossiers, consultPorteurSortState);

    const resultsDiv = document.getElementById('consult-results');
    if (!resultsDiv) return;

    resultsDiv.innerHTML = '';

    if (dossiers.length === 0) {
        const p = document.createElement('p');
        p.className = 'no-results';
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
    table.className = 'consult-porteur-table';
    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');
    const headerConfig = [
        { text: 'Date réunion', sortKey: 'Date_de_la_reunion', col: 'date-reunion' },
        { text: 'Dossier', sortKey: 'Dossier', col: 'dossier' },
        { text: 'Actions', sortKey: 'Actions_a_mettre_en_uvre_etapes', col: 'actions' },
        { text: 'Porteur(s)', sortKey: 'Porteur_s_', col: 'porteurs' },
        { text: 'Échéance', sortKey: 'Echeance', col: 'echeance' },
        { text: 'État', sortKey: 'Etat', col: 'etat' },
        { text: 'Date de modification', sortKey: 'Enregistrement', col: 'date-enr' }
    ];

    headerConfig.forEach(({ text, sortKey, col }) => {
        const th = document.createElement('th');
        th.dataset.col = col;
        const isActiveSort = consultPorteurSortState.key === sortKey;

        const button = document.createElement('button');
        button.type = 'button';
        button.className = 'sort-header-button';
        button.setAttribute('aria-label', `Trier par ${text}`);

        const label = document.createElement('span');
        label.textContent = text;
        button.appendChild(label);

        const indicator = document.createElement('span');
        indicator.className = 'sort-indicator';
        indicator.setAttribute('aria-hidden', 'true');

        const upTriangle = document.createElement('span');
        upTriangle.className = 'triangle triangle-up';
        upTriangle.textContent = '▲';
        indicator.appendChild(upTriangle);

        const downTriangle = document.createElement('span');
        downTriangle.className = 'triangle triangle-down';
        downTriangle.textContent = '▼';
        indicator.appendChild(downTriangle);

        button.appendChild(indicator);

        button.addEventListener('click', () => {
            updateConsultPorteurSortState(sortKey);
            consultByPorteur();
        });

        if (isActiveSort) {
            th.classList.add('sorted');
            if (consultPorteurSortState.direction === 'asc') {
                th.classList.add('sorted-asc');
                th.setAttribute('aria-sort', 'ascending');
            } else {
                th.classList.add('sorted-desc');
                th.setAttribute('aria-sort', 'descending');
            }
        } else {
            th.setAttribute('aria-sort', 'none');
        }

        th.appendChild(button);
        headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);
    table.appendChild(thead);

    const tbody = document.createElement('tbody');

    dossiers.forEach(dossier => {
        const etatName = dossier.Etat || '';
        const porteurs = getDossierPorteurs(dossier).join(', ');

        const tr = document.createElement('tr');
        applyEtatStyle(tr, etatName);
        makeDossierRowClickable(tr, dossier.Dossier);

        const tdDate = document.createElement('td');
        tdDate.textContent = formatDateShort(dossier.Date_de_la_reunion);
        tr.appendChild(tdDate);

        const tdDossier = document.createElement('td');
        tdDossier.textContent = dossier.Dossier || '';
        tr.appendChild(tdDossier);

        const tdActions = document.createElement('td');
        tdActions.textContent = dossier.Actions_a_mettre_en_uvre_etapes || '';
        tdActions.style.whiteSpace = 'pre-wrap';
        tr.appendChild(tdActions);

        const tdPorteurs = document.createElement('td');
        tdPorteurs.textContent = porteurs;
        tr.appendChild(tdPorteurs);

        const tdEcheance = document.createElement('td');
        tdEcheance.textContent = formatDateShort(dossier.Echeance);
        tr.appendChild(tdEcheance);

        const tdEtat = document.createElement('td');
        tdEtat.textContent = etatName;
        tr.appendChild(tdEtat);

        const tdEnregistrement = document.createElement('td');
        tdEnregistrement.textContent = formatDateShort(dossier.Enregistrement);
        tr.appendChild(tdEnregistrement);

        tbody.appendChild(tr);
    });

    table.appendChild(tbody);
    tagConsultTableColumns(table);
    tableContainer.appendChild(table);
    section.appendChild(tableContainer);
    resultsDiv.appendChild(section);

    togglePrintButton();
}

function updateConsultPorteurSortState(nextKey) {
    if (consultPorteurSortState.key === nextKey) {
        consultPorteurSortState.direction = consultPorteurSortState.direction === 'asc' ? 'desc' : 'asc';
        return;
    }

    consultPorteurSortState.key = nextKey;
    consultPorteurSortState.direction = 'asc';

    if (nextKey === 'Date_de_la_reunion' || nextKey === 'Echeance' || nextKey === 'Enregistrement') {
        consultPorteurSortState.direction = 'desc';
    }
}

function getConsultPorteurSortValue(dossier, sortKey) {
    switch (sortKey) {
        case 'Date_de_la_reunion':
            return dossier.Date_de_la_reunion || 0;
        case 'Echeance':
            return dossier.Echeance || 0;
        case 'Enregistrement':
            return dossier.Enregistrement || 0;
        case 'Etat': {
            const sortIndex = getEtatSortOrder().indexOf(dossier.Etat || '');
            return sortIndex === -1 ? Number.MAX_SAFE_INTEGER : sortIndex;
        }
        case 'Porteur_s_':
            return getDossierPorteurs(dossier).join(', ').toLowerCase();
        case 'Dossier':
            return (dossier.Dossier || '').toLowerCase();
        case 'Actions_a_mettre_en_uvre_etapes':
            return (dossier.Actions_a_mettre_en_uvre_etapes || '').toLowerCase();
        default:
            return '';
    }
}

function sortConsultPorteurDossiers(dossiers, sortState) {
    const directionFactor = sortState.direction === 'asc' ? 1 : -1;

    return [...dossiers].sort((a, b) => {
        const valueA = getConsultPorteurSortValue(a, sortState.key);
        const valueB = getConsultPorteurSortValue(b, sortState.key);

        if (typeof valueA === 'number' && typeof valueB === 'number') {
            return (valueA - valueB) * directionFactor;
        }

        return String(valueA).localeCompare(String(valueB), 'fr', { sensitivity: 'base' }) * directionFactor;
    });
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
        p.className = 'no-results';
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
        const etatName = dossier.Etat || '';
        const porteurs = getDossierPorteurs(dossier).join(', ');

        const tr = document.createElement('tr');
        applyEtatStyle(tr, etatName);
        makeDossierRowClickable(tr, dossier.Dossier);

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
        tdDate.textContent = formatDateShort(dossier.Date_de_la_reunion);
        tr.appendChild(tdDate);

        const tdEtat = document.createElement('td');
        tdEtat.textContent = etatName;
        tr.appendChild(tdEtat);

        tbody.appendChild(tr);
    });

    table.appendChild(tbody);
    tagConsultTableColumns(table);
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
        if (resultsDiv.innerHTML.trim() !== '' && !resultsDiv.innerHTML.includes('no-results')) {
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

// ----------------------------------------
// Construction des tableaux de dossiers (onglets Modifier et Réunion)
//
// Chaque colonne est identifiée par une clé stable posée sur le <td> via
// data-col, jamais par sa position. makeFieldsEditable() et les fonctions
// saveModifications*/saveReunionModifications s'appuient sur cette clé.
// Clés : dossier | porteurs | actions | echeance | date-reunion | etat
//        | date-enr | change
// ----------------------------------------

const DOSSIER_COLUMN_LABELS = {
    'dossier': 'Dossier',
    'porteurs': 'Porteur(s)',
    'actions': 'Actions',
    'echeance': 'Échéance',
    'date-reunion': 'Date réunion',
    'etat': 'État',
    'date-enr': 'Enregistré le',
    'change': "Changement d'état"
};

function getDossierPorteursText(dossier) {
    return getDossierPorteurs(dossier).join(', ');
}

// En-tête (texte) -> clé de colonne, pour les tableaux de consultation
// construits en createElement (sans data-col d'origine).
const CONSULT_HEADER_TO_COL = {
    'Dossier': 'dossier',
    'Porteur(s)': 'porteurs',
    'Actions': 'actions',
    'Échéance': 'echeance',
    'Date réunion': 'date-reunion',
    'Date d’enregistrement': 'date-enr',
    "Date d'enregistrement": 'date-enr',
    'Date de modification': 'date-enr',
    'État': 'etat'
};

/**
 * Pose data-col sur les <th>/<td> d'un tableau (pour appliquer les largeurs
 * CSS). Utilise data-col déjà présent sur le <th>, sinon le texte de l'en-tête.
 */
function tagConsultTableColumns(table) {
    if (!table) return;
    const headerCells = [...table.querySelectorAll('thead th')];
    const cols = headerCells.map(th => th.dataset.col || CONSULT_HEADER_TO_COL[th.textContent.trim()] || '');

    headerCells.forEach((th, i) => {
        if (cols[i] && !th.dataset.col) th.dataset.col = cols[i];
    });
    table.querySelectorAll('tbody tr').forEach(tr => {
        [...tr.children].forEach((td, i) => {
            if (cols[i]) td.dataset.col = cols[i];
        });
    });
}

/**
 * Génère le HTML d'une cellule selon sa clé de colonne.
 */
function buildDossierCell(dossier, col) {
    switch (col) {
        case 'dossier':
            return `<td data-col="dossier">${escapeHtml(dossier.Dossier || '')}</td>`;
        case 'porteurs':
            return `<td data-col="porteurs">${escapeHtml(getDossierPorteursText(dossier))}</td>`;
        case 'actions':
            return `<td data-col="actions">${escapeHtml(dossier.Actions_a_mettre_en_uvre_etapes || '').replace(/\n/g, '<br>')}</td>`;
        case 'echeance':
            return `<td data-col="echeance">${escapeHtml(formatDateShort(dossier.Echeance))}</td>`;
        case 'date-reunion':
            return `<td data-col="date-reunion">${escapeHtml(formatDateShort(dossier.Date_de_la_reunion))}</td>`;
        case 'etat':
            return `<td data-col="etat">${escapeHtml(dossier.Etat || '')}</td>`;
        case 'date-enr':
            return `<td data-col="date-enr">${escapeHtml(formatDateShort(dossier.Enregistrement))}</td>`;
        case 'change':
            return `<td data-col="change"><select class="etat-change-select"><option value="">-- Aucun changement --</option></select></td>`;
        default:
            return '<td></td>';
    }
}

/**
 * Génère un tableau complet <table>…</table> pour une liste de dossiers.
 * @param {object[]} dossiers
 * @param {string[]} columns - clés de colonnes, dans l'ordre d'affichage
 */
function buildDossierTable(dossiers, columns) {
    const head = '<thead><tr>'
        + columns.map(col => `<th data-col="${escapeHtmlAttribute(col)}">${escapeHtml(DOSSIER_COLUMN_LABELS[col] || '')}</th>`).join('')
        + '</tr></thead>';

    const body = '<tbody>' + dossiers.map(dossier => {
        return `<tr${etatStyleAttr(dossier.Etat || '')} data-dossier-id="${escapeHtmlAttribute(dossier.id)}">`
            + columns.map(col => buildDossierCell(dossier, col)).join('')
            + '</tr>';
    }).join('');

    return `<table class="dossiers-editable">${head}${body}</table>`;
}

/**
 * Lit les valeurs éditées d'une ligne de tableau de dossiers.
 * Le repérage se fait par data-col (jamais par position). Les colonnes absentes
 * retombent sur les valeurs d'origine du dossier.
 * @param {HTMLTableRowElement} row
 * @param {object} dossier - enregistrement ODJ d'origine
 * @returns {{nomCellule: (string|null), porteurs: string[], actions: string,
 *            echeance: *, changementEtat: string}}
 */
function readEditableRow(row, dossier) {
    const dossierCell = row.querySelector('[data-col="dossier"]');
    const nomCellule = dossierCell ? dossierCell.textContent.trim() : null;

    const porteurSelect = row.querySelector('[data-col="porteurs"] select');
    const porteurs = porteurSelect
        ? Array.from(porteurSelect.selectedOptions).map(opt => opt.value).filter(Boolean)
        : getDossierPorteurs(dossier);

    const actionsCell = row.querySelector('[data-col="actions"]');
    const actions = actionsCell
        ? extractCellText(actionsCell).trim()
        : (dossier.Actions_a_mettre_en_uvre_etapes || '');

    const echeanceInput = row.querySelector('[data-col="echeance"] input[type="date"]');
    let echeance = dossier.Echeance;
    if (echeanceInput) {
        // Colonne « Échéance » éditable : on prend la valeur du champ.
        // Champ vidé (bouton d'effacement du date-picker) => échéance retirée.
        echeance = echeanceInput.value
            ? Math.floor(new Date(echeanceInput.value).getTime() / 1000)
            : null;
    }

    // Date de réunion : éditable mais obligatoire — un champ vidé garde l'ancienne valeur.
    const dateReunionInput = row.querySelector('[data-col="date-reunion"] input[type="date"]');
    let dateReunion = dossier.Date_de_la_reunion;
    if (dateReunionInput && dateReunionInput.value) {
        dateReunion = Math.floor(new Date(dateReunionInput.value).getTime() / 1000);
    }

    const etatChangeSelect = row.querySelector('.etat-change-select');
    const changementEtat = etatChangeSelect ? etatChangeSelect.value : '';

    return { nomCellule, porteurs, actions, echeance, dateReunion, changementEtat };
}

/** Affiche le sélecteur correspondant au mode « Modifier » (date/dossier/porteur/échéance). */
function setModifySelectorVisibility(type) {
    const dateSelector = document.getElementById('modify-date-selector');
    const dossierSelector = document.getElementById('modify-dossier-selector');
    const porteurSelector = document.getElementById('modify-porteur-selector');
    const echeanceSelector = document.getElementById('modify-echeance-selector');

    if (dateSelector) dateSelector.classList.toggle('hidden', type !== 'date');
    if (dossierSelector) dossierSelector.classList.toggle('hidden', type !== 'dossier');
    if (porteurSelector) porteurSelector.classList.toggle('hidden', type !== 'porteur');
    if (echeanceSelector) echeanceSelector.classList.toggle('hidden', type !== 'echeance');
}

async function handleModifyTypeChange(event) {
    const type = event.target.value;

    // Enregistrer une éventuelle modification en attente avant de tout réinitialiser
    await flushModifyAutoSave();

    setModifySelectorVisibility(type);

    const resultsDiv = document.getElementById('modify-results');
    if (resultsDiv) resultsDiv.innerHTML = '';

    const buttons = document.getElementById('modify-buttons');
    if (buttons) buttons.classList.add('hidden');

    // Vider les sélections lors du changement de mode
    clearAllSelections();

    modifyContext.type = null;
    modifyContext.value = null;
    modifyContext.secondValue = null;
}

/** Vide la zone de résultats de l'onglet Modifier (plus rien à éditer). */
function hideModifyResults() {
    const resultsDiv = document.getElementById('modify-results');
    if (resultsDiv) resultsDiv.innerHTML = '';
    const buttons = document.getElementById('modify-buttons');
    if (buttons) buttons.classList.add('hidden');
    modifyContext.type = null;
    modifyContext.value = null;
    modifyContext.secondValue = null;
}

function modifyByDate() {
    const dateSelect = document.getElementById('modify-date-select');
    if (!dateSelect) return;

    const date = dateSelect.value;
    if (!date) {
        // Plus aucune date sélectionnable (ex. dernier dossier de la date supprimé)
        hideModifyResults();
        return;
    }

    // Sauvegarder le contexte
    modifyContext.type = 'date';
    modifyContext.value = date;
    modifyContext.secondValue = null;

    const modifyResults = document.getElementById('modify-results');
    if (!modifyResults) return;

    const dateValue = typeof date === 'string' ? Number.parseFloat(date) : date;

    let dossiers = tablesData.ODJ.filter(o => o.Date_de_la_reunion == dateValue);

    // Dédoublonner : un seul record par dossier (le plus récent, par ID_Dossier)
    dossiers = getLatestEntriesPerDossier(dossiers);

    // Trier par état (du pire au meilleur)
    dossiers = sortByEtat(dossiers);

    let html = '';

    if (dossiers.length > 0) {
        html += '<div class="section"><h2 class="section-title">Ordre du jour</h2>';
        html += `<div class="result-item">`;
        html += `<div class="result-header">Date : ${escapeHtml(formatDate(dateValue))}</div>`;
        html += '</div>';
        html += '<div class="table-container">';
        html += buildDossierTable(dossiers, ['dossier', 'porteurs', 'actions', 'echeance', 'etat', 'change']);
        html += '</div></div>';
    }

    modifyResults.innerHTML = html || '<p class="no-results">Aucune donnée pour cette date</p>';

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

    // Résoudre l'ID_Dossier à partir du nom saisi, puis inclure tous les records partageant cet ID
    const refRecord = tablesData.ODJ.find(odj => odj.Dossier === dossierName && odj.ID_Dossier);
    let dossiers = refRecord
        ? tablesData.ODJ.filter(odj => odj.ID_Dossier === refRecord.ID_Dossier)
        : tablesData.ODJ.filter(odj => odj.Dossier === dossierName);

    if (dossiers.length === 0) {
        resultsDiv.innerHTML = '<p class="no-results">Aucun dossier trouvé</p>';
        const buttons = document.getElementById('modify-buttons');
        if (buttons) buttons.classList.add('hidden');
        return;
    }

    // On ne modifie que la dernière version en date ; l'historique complet se
    // consulte dans l'onglet « Consulter ».
    dossiers = getLatestEntriesPerDossier(dossiers);

    let html = '<div class="section"><h2 class="section-title">Dernière version du dossier</h2>';
    html += '<div class="table-container">';
    html += buildDossierTable(dossiers, ['date-reunion', 'porteurs', 'actions', 'echeance', 'etat', 'date-enr', 'change']);
    html += '</div></div>';
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
    const dossiers = tablesData.ODJ.filter(odj => {
        if (!Array.isArray(odj.Porteur_s_)) return false;
        return getDossierPorteurs(odj).includes(porteurName);
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

    // Afficher tous les dossiers par défaut
    modifyByPorteurAllDossiers();
}

/**
 * Rafraîchit l'affichage « Modifier par porteur » selon qu'un dossier précis
 * est sélectionné ou non. Utilisé par le sélecteur de dossier, la case
 * « Masquer les dossiers échus » et les filtres d'état (délégation).
 */
async function handleModifyPorteurDossierSelectChange() {
    // Enregistrer d'éventuelles modifications avant de reconstruire l'affichage
    await flushModifyAutoSave();

    const dossierSelect = document.getElementById('modify-porteur-dossier-select');
    if (dossierSelect && dossierSelect.value !== '') {
        modifyByPorteurDossier();
    } else {
        modifyByPorteurAllDossiers();
    }
}

function populateModifyPorteurEtatFilters() {
    buildEtatFilterCheckboxes(
        document.getElementById('modify-filter-etat-checkboxes'),
        'modify-filter-etat'
    );
}

function modifyByPorteurAllDossiers() {
    const porteurSelect = document.getElementById('modify-porteur-select');
    if (!porteurSelect) return;

    const porteurName = porteurSelect.value;
    if (!porteurName) return;

    // Sauvegarder le contexte (mode tous les dossiers)
    modifyContext.type = 'porteur';
    modifyContext.value = porteurName;
    modifyContext.secondValue = null;

    const hideExpired = document.getElementById('modify-hide-expired')?.checked || false;

    const selectedEtats = new Set(
        Array.from(document.querySelectorAll('input[name="modify-filter-etat"]:checked'))
            .map(cb => cb.value)
    );

    let dossiers = tablesData.ODJ.filter(odj => {
        if (!Array.isArray(odj.Porteur_s_)) return false;
        return getDossierPorteurs(odj).includes(porteurName);
    });

    // Garder uniquement le record le plus récent par dossier (avant de filtrer par état)
    dossiers = getLatestEntriesPerDossier(dossiers);

    // Filtrer par état sur le vrai état courant du dossier
    dossiers = dossiers.filter(dossier => {
        const etatName = dossier.Etat || '';
        return selectedEtats.has(etatName);
    });

    // Filtrer les dossiers échus si le toggle est activé (sur l'échéance courante)
    if (hideExpired) {
        const todayTs = todayCalendarTs();
        dossiers = dossiers.filter(dossier => {
            if (!dossier.Echeance) return true; // Garder les dossiers sans échéance
            return dossier.Echeance >= todayTs;
        });
    }

    // Trier par état (du pire au meilleur)
    dossiers = sortByEtat(dossiers);

    const resultsDiv = document.getElementById('modify-results');
    if (!resultsDiv) return;

    if (dossiers.length === 0) {
        resultsDiv.innerHTML = '<p class="no-results">Aucun dossier trouvé</p>';
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

        // Obtenir la date d'enregistrement représentative du groupe (la plus récente)
        const dateRepresentative = Math.max(...group.map(d => d.Enregistrement || 0));

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
        html += '<div class="table-container">';
        html += buildDossierTable(groupe.dossiers, ['date-reunion', 'porteurs', 'actions', 'echeance', 'etat', 'date-enr', 'change']);
        html += '</div>';
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

    const resultsDiv = document.getElementById('modify-results');
    if (!resultsDiv) return;

    let dossiers = tablesData.ODJ.filter(odj => odj.Dossier === dossierName);

    if (dossiers.length === 0) {
        resultsDiv.innerHTML = '<p class="no-results">Aucun dossier trouvé</p>';
        const buttons = document.getElementById('modify-buttons');
        if (buttons) buttons.classList.add('hidden');
        return;
    }

    // On ne modifie que la dernière version en date (historique dans « Consulter »).
    dossiers = getLatestEntriesPerDossier(dossiers);

    let html = `<div class="section"><h2 class="section-title">Dernière version du dossier : ${escapeHtml(dossierName)}</h2>`;
    html += '<div class="table-container">';
    html += buildDossierTable(dossiers, ['date-reunion', 'porteurs', 'actions', 'echeance', 'etat', 'date-enr', 'change']);
    html += '</div></div>';
    resultsDiv.innerHTML = html;

    makeFieldsEditable(resultsDiv);

    const buttons = document.getElementById('modify-buttons');
    if (buttons) buttons.classList.remove('hidden');
}

function modifyByEcheance() {
    const echeanceSelect = document.getElementById('modify-echeance-select');
    if (!echeanceSelect) return;

    const echeance = echeanceSelect.value;
    if (!echeance) {
        // Plus aucune échéance sélectionnable (dernier dossier concerné supprimé)
        hideModifyResults();
        return;
    }

    // Sauvegarder le contexte
    modifyContext.type = 'echeance';
    modifyContext.value = echeance;
    modifyContext.secondValue = null;

    const modifyResults = document.getElementById('modify-results');
    if (!modifyResults) return;

    const echeanceValue = typeof echeance === 'string' ? Number.parseFloat(echeance) : echeance;

    let dossiers = tablesData.ODJ.filter(o => o.Echeance == echeanceValue);

    if (dossiers.length === 0) {
        modifyResults.innerHTML = '<p class="no-results">Aucun dossier trouvé</p>';
        const buttons = document.getElementById('modify-buttons');
        if (buttons) buttons.classList.add('hidden');
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
    html += '<div class="table-container">';
    html += buildDossierTable(dossiers, ['dossier', 'porteurs', 'actions', 'date-reunion', 'etat', 'change']);
    html += '</div></div>';
    modifyResults.innerHTML = html;

    makeFieldsEditable(modifyResults);

    const buttons = document.getElementById('modify-buttons');
    if (buttons) buttons.classList.remove('hidden');
}

const EDITABLE_CELL_STYLE = {
    border: '1px solid #d9d9d9',
    backgroundColor: 'rgba(255, 255, 255, 0.5)',
    color: '#000'
};

const EDITABLE_CONTROL_STYLE = {
    width: '100%',
    border: '1px solid #d9d9d9',
    backgroundColor: 'rgba(255, 255, 255, 0.9)'
};

function applyStyle(el, style) {
    Object.assign(el.style, style);
}

/**
 * Insère un saut de ligne (<br>) à la position du curseur, sans execCommand.
 * En fin de bloc, un <br> seul n'est pas rendu : on ajoute un <br> témoin
 * après et on place le curseur entre les deux (sinon il faut appuyer deux
 * fois sur Entrée).
 */
function insertLineBreakAtCaret() {
    const selection = window.getSelection();
    if (!selection || selection.rangeCount === 0) return;

    const range = selection.getRangeAt(0);
    range.deleteContents();

    const br = document.createElement('br');
    range.insertNode(br);

    const next = br.nextSibling;
    const atEnd = !next || (next.nodeType === Node.TEXT_NODE && next.textContent === '');
    if (atEnd) {
        br.parentNode.insertBefore(document.createElement('br'), br.nextSibling);
    }

    range.setStartAfter(br);
    range.collapse(true);
    selection.removeAllRanges();
    selection.addRange(range);
}

/**
 * Insère du texte brut à la position du curseur (les \n deviennent des <br>),
 * sans execCommand ni HTML.
 */
function insertPlainTextAtCaret(text) {
    const selection = window.getSelection();
    if (!selection || selection.rangeCount === 0) return;
    const range = selection.getRangeAt(0);
    range.deleteContents();

    const frag = document.createDocumentFragment();
    String(text).split('\n').forEach((line, i) => {
        if (i > 0) frag.appendChild(document.createElement('br'));
        if (line) frag.appendChild(document.createTextNode(line));
    });
    const last = frag.lastChild;
    range.insertNode(frag);

    if (last) {
        range.setStartAfter(last);
        range.setEndAfter(last);
        selection.removeAllRanges();
        selection.addRange(range);
    }
}

/**
 * Extrait le texte d'une cellule éditable en préservant les sauts de ligne :
 * <br> et fins de blocs (<div>, <p>) deviennent des \n. Robuste aux <div>
 * qu'un navigateur ou un collage aurait pu insérer.
 */
function extractCellText(el) {
    let out = '';
    el.childNodes.forEach(node => {
        if (node.nodeType === Node.TEXT_NODE) {
            out += node.textContent;
        } else if (node.nodeName === 'BR') {
            out += '\n';
        } else if (node.nodeType === Node.ELEMENT_NODE) {
            const isBlock = node.nodeName === 'DIV' || node.nodeName === 'P' || node.nodeName === 'LI';
            if (isBlock && out && !out.endsWith('\n')) out += '\n';
            out += extractCellText(node);
            if (isBlock && !out.endsWith('\n')) out += '\n';
        }
    });
    return out;
}

/**
 * Rend éditables les cellules d'un ou plusieurs tableaux de dossiers.
 * Le type de chaque cellule est décrit par son attribut `data-col` posé au
 * moment de la génération (voir buildDossierTable), et non par sa position :
 *   - data-col="dossier"  → texte éditable
 *   - data-col="porteurs" → <select multiple>
 *   - data-col="actions"  → texte éditable multiligne
 *   - data-col="echeance" → <input type="date">
 *   - toute autre valeur  → cellule laissée en lecture seule
 * @param {HTMLElement} container
 */
function makeFieldsEditable(container) {
    const porteurChoices = getPorteurChoices();

    container.querySelectorAll('table tbody tr').forEach(row => {
        const dossierData = tablesData.ODJ.find(d => d.id == row.dataset.dossierId);

        row.querySelectorAll('td[data-col]').forEach(td => {
            const col = td.dataset.col;

            if (col === 'dossier') {
                td.contentEditable = true;
                applyStyle(td, EDITABLE_CELL_STYLE);
                // Champ mono-ligne : pas de saut de ligne.
                td.addEventListener('keydown', function (event) {
                    if (event.key === 'Enter') event.preventDefault();
                });
                td.addEventListener('paste', function (event) {
                    if (!event.clipboardData) return;
                    event.preventDefault();
                    insertPlainTextAtCaret(event.clipboardData.getData('text/plain').replace(/\s+/g, ' '));
                });
            } else if (col === 'actions') {
                td.contentEditable = true;
                applyStyle(td, EDITABLE_CELL_STYLE);
                td.addEventListener('keydown', function (event) {
                    if (event.key === 'Enter') {
                        event.preventDefault();
                        insertLineBreakAtCaret();
                    }
                });
                // Coller en texte brut : évite d'injecter des <div>/<span>/styles
                // qui fausseraient la relecture (fausse « modification »).
                td.addEventListener('paste', function (event) {
                    if (!event.clipboardData) return;
                    event.preventDefault();
                    insertPlainTextAtCaret(event.clipboardData.getData('text/plain'));
                });
            } else if (col === 'porteurs' && dossierData) {
                const currentPorteurs = getDossierPorteurs(dossierData);
                const select = document.createElement('select');
                select.multiple = true;
                applyStyle(select, EDITABLE_CONTROL_STYLE);
                select.style.minHeight = '60px';

                // Choix courants + porteurs déjà sur le dossier même s'ils ont
                // été retirés de la liste (pour ne pas les perdre à l'édition).
                const options = [...new Set([...porteurChoices, ...currentPorteurs])];

                options.forEach(personne => {
                    const option = document.createElement('option');
                    option.value = personne;
                    option.textContent = personne;
                    if (currentPorteurs.includes(personne)) {
                        option.selected = true;
                    }
                    select.appendChild(option);
                });

                // Sélection/désélection sans maintenir Ctrl
                select.addEventListener('mousedown', function (e) {
                    e.preventDefault();
                    if (e.target.tagName === 'OPTION') {
                        e.target.selected = !e.target.selected;
                        select.focus();
                        setTimeout(() => reorderSelectOptions(select), 10);
                        // La sélection programmatique ne déclenche pas d'événement
                        // natif : on le simule pour l'enregistrement automatique
                        // (onglet Réunion, « Modifier par porteur »).
                        select.dispatchEvent(new Event('change', { bubbles: true }));
                    }
                });

                td.innerHTML = '';
                td.appendChild(select);
                td.style.padding = '4px';
                reorderSelectOptions(select);
            } else if ((col === 'echeance' || col === 'date-reunion') && dossierData) {
                const sourceTs = col === 'echeance'
                    ? dossierData.Echeance
                    : dossierData.Date_de_la_reunion;

                const dateInput = document.createElement('input');
                dateInput.type = 'date';
                applyStyle(dateInput, EDITABLE_CONTROL_STYLE);
                dateInput.style.padding = '4px';

                if (sourceTs) {
                    dateInput.value = tsToInputDate(sourceTs);
                }

                td.innerHTML = '';
                td.appendChild(dateInput);
                td.style.padding = '4px';
            }
        });
    });

    populateEtatChangeSelects(container);
}

/**
 * Remplit les <select class="etat-change-select"> et gère la recoloration de
 * la ligne à la volée quand on choisit un nouvel état.
 * @param {HTMLElement} container
 */
function populateEtatChangeSelects(container) {
    const etatsTries = getEtatDropdownOrder();

    container.querySelectorAll('.etat-change-select').forEach(select => {
        etatsTries.forEach(etat => {
            const option = document.createElement('option');
            option.value = etat;
            option.textContent = etat;
            select.appendChild(option);
        });

        select.addEventListener('change', function () {
            const row = this.closest('tr');
            if (!row) return;

            let etatName = this.value;
            if (!etatName) {
                // Aucun changement : restaurer l'état d'origine du dossier
                const dossierData = tablesData.ODJ.find(d => d.id == row.dataset.dossierId);
                etatName = dossierData ? (dossierData.Etat || '') : '';
            }

            applyEtatStyle(row, etatName);
        });
    });
}

function reorderSelectOptions(select) {
    if (!select || select.tagName !== 'SELECT') return;

    const options = Array.from(select.options);
    const selectedValues = new Set(options.filter(opt => opt.selected).map(opt => opt.value));

    // Réordonner en DÉPLAÇANT les noeuds (appendChild d'un enfant déjà présent
    // = déplacement in-place). On évite `select.innerHTML = ''` qui, sur
    // certains navigateurs, réinitialise la sélectedness des <option> à la
    // ré-insertion -> la modification des porteurs était perdue.
    const ordered = [
        ...options.filter(opt => selectedValues.has(opt.value)),
        ...options.filter(opt => !selectedValues.has(opt.value))
    ];
    ordered.forEach(opt => select.appendChild(opt));

    // Réaffirmer la sélection par sécurité.
    ordered.forEach(opt => { opt.selected = selectedValues.has(opt.value); });
}

// ----------------------------------------
// Enregistrement des modifications (onglet Modifier)
//
// Toute modification est écrite comme une NOUVELLE ligne ODJ datée
// d'aujourd'hui (Enregistrement = maintenant), pour suivre l'évolution du
// dossier même hors réunion. Une seule ligne par dossier et par jour : si une
// ligne a déjà été créée aujourd'hui pour ce dossier à cette date de réunion,
// elle est mise à jour (findTodayRecord).
// ----------------------------------------

/** true si la ligne éditée diffère de l'enregistrement source. */
function modifyRowHasChanges(dossier, v) {
    if (v.changementEtat) return true;

    if (v.nouveauDossier !== null && v.nouveauDossier !== (dossier.Dossier || '')) return true;

    if ((v.actions || '').trim() !== (dossier.Actions_a_mettre_en_uvre_etapes || '').trim()) return true;

    if ((v.echeance ?? null) !== (dossier.Echeance ?? null)) return true;

    if ((v.dateReunion ?? null) !== (dossier.Date_de_la_reunion ?? null)) return true;

    const orig = getDossierPorteurs(dossier).slice().sort();
    const next = [...(v.porteurs || [])].sort();
    if (orig.length !== next.length || orig.some((x, i) => x !== next[i])) return true;

    return false;
}

/**
 * Actions Grist pour enregistrer l'état modifié d'un dossier : upsert de la
 * ligne du jour, à la date de réunion (éventuellement modifiée) de la ligne.
 */
function buildModifyUpsertActions(dossier, v) {
    const dateReunion = v.dateReunion ?? dossier.Date_de_la_reunion;
    const existingToday = findTodayRecord(dossier, dateReunion);

    const data = {
        Date_de_la_reunion: dateReunion,
        Dossier: v.nouveauDossier ?? dossier.Dossier,
        ID_Dossier: dossier.ID_Dossier || '',
        Porteur_s_: ['L', ...v.nouveauxPorteurs],
        Actions_a_mettre_en_uvre_etapes: v.actions,
        Echeance: v.nouvelleEcheance ?? null,
        Etat: v.nouvelEtat || dossier.Etat || null,
        Enregistrement: Date.now() / 1000
    };

    return existingToday
        ? [['UpdateRecord', 'ODJ', existingToday.id, data]]
        : [['AddRecord', 'ODJ', null, data]];
}

/** Toutes les <tr> éditables actuellement affichées dans l'onglet Modifier. */
function getModifyResultRows() {
    const modifyResults = document.getElementById('modify-results');
    if (!modifyResults) return [];
    return [...modifyResults.querySelectorAll('table tbody tr')];
}

/**
 * Première passe : lit chaque ligne affichée, calcule les changements et
 * repère les dossiers marqués « Supprimer le dossier ».
 * @param {(dossier: object, edited: object) => (string|null)} resolveNom
 */
function collectModifyRows(resolveNom) {
    const rowsData = [];
    const dossiersASupprimer = [];

    for (const row of getModifyResultRows()) {
        const dossierId = Number.parseInt(row.dataset.dossierId, 10);
        const dossier = tablesData.ODJ.find(d => d.id === dossierId);
        if (!dossier) continue;

        const edited = readEditableRow(row, dossier);
        const nouveauDossier = resolveNom(dossier, edited);
        const v = {
            dossier,
            nouveauDossier,
            nouveauxPorteurs: edited.porteurs,
            actions: edited.actions,
            nouvelleEcheance: edited.echeance,
            dateReunion: edited.dateReunion,
            nouvelEtat: edited.changementEtat
        };
        v.hasChanges = modifyRowHasChanges(dossier, {
            nouveauDossier,
            porteurs: edited.porteurs,
            actions: edited.actions,
            echeance: edited.echeance,
            dateReunion: edited.dateReunion,
            changementEtat: edited.changementEtat
        });

        rowsData.push(v);
        if (edited.changementEtat === etatRoles.supprimer) {
            dossiersASupprimer.push({ dossier, nom: nouveauDossier });
        }
    }

    return { rowsData, dossiersASupprimer };
}

/**
 * Deuxième passe : applique les modifications (une seule action groupée).
 * @returns {Promise<boolean>} true si au moins une écriture a eu lieu
 */
async function processModifyRows(rowsData, deletedIds) {
    const actions = [];
    for (const v of rowsData) {
        if (v.nouvelEtat === etatRoles.supprimer) continue;
        if (deletedIds.has(v.dossier.id)) continue;
        if (!v.hasChanges) continue;
        actions.push(...buildModifyUpsertActions(v.dossier, v));
    }
    if (actions.length === 0) return false;
    await grist.docApi.applyUserActions(actions);
    return true;
}

/**
 * Enchaînement commun à tous les modes : collecte -> suppression -> upsert.
 * @returns {Promise<'saved'|'nochange'|'cancelled'>}
 */
async function runModifySave(resolveNom) {
    if (getModifyResultRows().length === 0) return 'nochange';

    const { rowsData, dossiersASupprimer } = collectModifyRows(resolveNom);

    const { status, deletedIds } = await confirmAndDeleteDossiers(dossiersASupprimer);
    if (status === 'cancelled') return 'cancelled';

    const wrote = await processModifyRows(rowsData, deletedIds);
    return (wrote || status === 'deleted') ? 'saved' : 'nochange';
}

/**
 * Exécute la sauvegarde correspondant au mode « Modifier » courant.
 * @returns {Promise<'saved'|'nochange'|'cancelled'>}
 */
function runModifySaveByType() {
    switch (modifyContext.type) {
        case 'date': return saveModificationsByDate();
        case 'dossier': return saveModificationsByDossier();
        case 'porteur': return saveModificationsByPorteur();
        case 'echeance': return saveModificationsByEcheance();
        default: return Promise.resolve('nochange');
    }
}

function reopenModifyForm() {
    // Copie locale : les fonctions modifyBy* réécrivent modifyContext au fil du re-render.
    const { type, value, secondValue } = modifyContext;
    if (!type || !value) return;

    const typeRadio = document.querySelector(`input[name="modify-type"][value="${type}"]`);
    if (typeRadio) typeRadio.checked = true;
    setModifySelectorVisibility(type);

    if (type === 'date') {
        const dateSelect = document.getElementById('modify-date-select');
        if (dateSelect) {
            dateSelect.value = value;
            modifyByDate();
        }
    } else if (type === 'dossier') {
        const dossierInput = document.getElementById('modify-dossier-input');
        if (dossierInput) {
            dossierInput.value = value;
            toggleClearButton('btn-clear-modify-dossier', value);
            modifyByDossier(value);
        }
    } else if (type === 'porteur') {
        const porteurSelect = document.getElementById('modify-porteur-select');
        if (porteurSelect) {
            porteurSelect.value = value;
            handleModifyPorteurSelectChange();

            // Restaurer le dossier précis éventuellement sélectionné
            setTimeout(() => {
                const dossierSelect = document.getElementById('modify-porteur-dossier-select');
                if (dossierSelect && secondValue) {
                    dossierSelect.value = secondValue;
                    modifyByPorteurDossier();
                }
            }, 100);
        }
    } else if (type === 'echeance') {
        const echeanceSelect = document.getElementById('modify-echeance-select');
        if (echeanceSelect) {
            echeanceSelect.value = value;
            modifyByEcheance();
        }
    }
}

async function saveModificationsByDate() {
    return runModifySave((dossier, edited) => edited.nomCellule);
}

async function saveModificationsByDossier() {
    const dossierInput = document.getElementById('modify-dossier-input');
    const dossierName = dossierInput ? dossierInput.value.trim() : '';
    if (!dossierName) return 'nochange';
    // Le nom n'est pas modifiable dans ce mode : on garde celui du dossier.
    return runModifySave(dossier => dossier.Dossier);
}

async function saveModificationsByPorteur() {
    const porteurSelect = document.getElementById('modify-porteur-select');
    const dossierSelect = document.getElementById('modify-porteur-dossier-select');
    if (!porteurSelect || !porteurSelect.value) return 'nochange';

    const dossierName = dossierSelect ? dossierSelect.value : '';
    const isAllDossiers = !dossierName;
    return runModifySave(dossier => (isAllDossiers ? dossier.Dossier : dossierName));
}

async function saveModificationsByEcheance() {
    return runModifySave((dossier, edited) => edited.nomCellule);
}

async function closeModifyForm() {
    // Les modifications sont enregistrées automatiquement : on s'assure qu'une
    // sauvegarde en attente est bien partie avant de refermer la vue.
    await flushModifyAutoSave();

    const resultsDiv = document.getElementById('modify-results');
    if (resultsDiv) resultsDiv.innerHTML = '';

    const buttons = document.getElementById('modify-buttons');
    if (buttons) buttons.classList.add('hidden');

    modifyContext.type = null;
    modifyContext.value = null;
    modifyContext.secondValue = null;

    clearAllSelections();
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
    const todayTs = todayCalendarTs();
    let defaultDate = null;

    dates.forEach(dateValue => {
        const option = document.createElement('option');
        option.value = dateValue;
        option.textContent = formatDate(dateValue);

        // Les dates sont triées en ordre décroissant (plus récente en premier)
        // On continue à itérer pour trouver la date la plus proche >= aujourd'hui
        if (dateValue >= todayTs) {
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
    // Exclure les dossiers dont le dernier état est « clôturé »
    dossierEcheance = dossierEcheance.filter(d => (d.Etat || '') !== etatRoles.cloture);
    // Exclure les dossiers déjà présents dans l'ODJ
    dossierEcheance = dossierEcheance.filter(d => !odjDossierNames.has(d.Dossier));

    // Récupérer les dossiers échus non clôturés
    // Inclus : échéance passée, OU échéance non définie mais réunion d'origine antérieure
    const reunionNum = typeof numDateValue === 'number' ? numDateValue : Number.parseFloat(numDateValue);
    let dossierExpired = tablesData.ODJ.filter(o => {
        // Pas de date de réunion ou réunion non antérieure → exclure
        const dateReunionNum = typeof o.Date_de_la_reunion === 'number' ? o.Date_de_la_reunion : Number.parseFloat(o.Date_de_la_reunion);
        if (Number.isNaN(dateReunionNum) || dateReunionNum >= reunionNum) return false;

        // Échéance non définie : inclure (dossier d'une réunion passée sans échéance)
        if (o.Echeance === null || o.Echeance === undefined || o.Echeance === '') return true;

        // Échéance définie : inclure seulement si elle est antérieure à la réunion sélectionnée
        const echeanceNum = typeof o.Echeance === 'number' ? o.Echeance : Number.parseFloat(o.Echeance);
        return !Number.isNaN(echeanceNum) && echeanceNum < reunionNum;
    });
    // Garder uniquement la dernière version de chaque dossier (avant de filtrer l'état)
    dossierExpired = getLatestEntriesPerDossier(dossierExpired);
    // Exclure les dossiers dont le dernier état est « clôturé »
    dossierExpired = dossierExpired.filter(d => (d.Etat || '') !== etatRoles.cloture);
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
        html += buildDossierTable(dossiers, ['dossier', 'porteurs', 'actions', 'echeance', 'etat', 'change']);
        container.innerHTML = html;
        makeFieldsEditable(container);
    } else {
        html = '<p class="no-results">Aucun dossier pour cette réunion</p>';
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
        html += buildDossierTable(dossiers, ['dossier', 'porteurs', 'actions', 'date-reunion', 'etat', 'change']);
        container.innerHTML = html;
        makeFieldsEditable(container);
    } else {
        html = '<p class="no-results">Aucun dossier à échéance.</p>';
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
        html += buildDossierTable(dossiers, ['dossier', 'porteurs', 'actions', 'echeance', 'date-reunion', 'etat', 'change']);
        container.innerHTML = html;
        makeFieldsEditable(container);
    } else {
        html = '<p class="no-results">Aucun dossier échu non clôturé.</p>';
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

// ========================================
// ENREGISTREMENT AUTOMATIQUE - ONGLET MODIFIER
// (tous les modes : date / dossier / porteur / échéance)
// ========================================

let modifyAutoSaveTimer = null;
let modifyAutoSaveInFlight = false;

function handleModifyAutoSaveEvent(e) {
    if (!modifyContext.type) return;

    const modifyResults = document.getElementById('modify-results');
    if (!modifyResults || !modifyResults.contains(e.target)) return;

    if (e.type === 'input') {
        // Frappe en cours dans une cellule éditable : repousser (mais ne pas
        // amorcer) un enregistrement pour ne pas reconstruire le tableau
        // pendant que l'utilisateur écrit.
        if (modifyAutoSaveTimer !== null) scheduleModifyAutoSave();
        return;
    }

    // change : listes déroulantes, date, multi-select (pas les contenteditable)
    if (e.type === 'change') {
        if (e.target.contentEditable === 'true') return;
        scheduleModifyAutoSave();
        return;
    }

    // focusout : uniquement les cellules contenteditable, et seulement quand le
    // focus a réellement quitté la zone d'édition (pas un simple passage de
    // cellule à cellule).
    if (e.type === 'focusout') {
        if (e.target.contentEditable !== 'true') return;
        setTimeout(() => {
            if (!modifyResults.contains(document.activeElement)) {
                scheduleModifyAutoSave();
            }
        }, 0);
    }
}

function scheduleModifyAutoSave() {
    clearTimeout(modifyAutoSaveTimer);
    modifyAutoSaveTimer = setTimeout(() => {
        modifyAutoSaveTimer = null;
        void performModifyAutoSave();
    }, 800);
}

/**
 * Force l'exécution immédiate d'un enregistrement en attente (changement
 * d'onglet, de mode, fermeture de la vue).
 */
async function flushModifyAutoSave() {
    if (modifyAutoSaveTimer === null) return;
    clearTimeout(modifyAutoSaveTimer);
    modifyAutoSaveTimer = null;
    await performModifyAutoSave();
}

/**
 * Emballe un gestionnaire d'événement qui va reconstruire l'affichage de
 * l'onglet Modifier : on enregistre d'abord toute modification en attente.
 */
function withModifyFlush(fn) {
    return async function (event) {
        await flushModifyAutoSave();
        return fn.call(this, event);
    };
}

async function performModifyAutoSave() {
    if (modifyAutoSaveInFlight) return;

    const modifyResults = document.getElementById('modify-results');
    if (!modifyResults || !modifyResults.querySelector('table tbody tr')) return;

    modifyAutoSaveInFlight = true;
    showModifySaveStatus('saving');
    try {
        const outcome = await runModifySaveByType();

        if (outcome === 'cancelled') {
            // Suppression annulée : on rouvre le formulaire tel quel
            showModifySaveStatus('idle');
            reopenModifyForm();
            return;
        }
        if (outcome === 'nochange') {
            // Rien à écrire : pas de rechargement, pas de reconstruction du tableau
            showModifySaveStatus('idle');
            return;
        }

        // Pas de removeDuplicateRecords ici : l'upsert par jour (findTodayRecord)
        // empêche déjà les doublons, et une déduplication globale supprimerait
        // à tort une ligne d'historique identique à un état antérieur du dossier.
        await loadAllTables();
        populateConsultSelectors();
        reopenModifyForm();
        showModifySaveStatus('saved');
    } catch (error) {
        console.error('Erreur enregistrement automatique (Modifier) :', error);
        showModifySaveStatus('error');
    } finally {
        modifyAutoSaveInFlight = false;
    }
}

function showModifySaveStatus(state) {
    const el = document.getElementById('modify-autosave-status');
    if (!el) return;

    el.classList.remove('autosave-saving', 'autosave-saved', 'autosave-error');
    while (el.firstChild) el.removeChild(el.firstChild);

    if (state === 'saving') {
        el.classList.add('autosave-saving');
        const spinner = document.createElement('span');
        spinner.className = 'autosave-spinner';
        spinner.textContent = '⟳';
        spinner.setAttribute('aria-hidden', 'true');
        el.appendChild(spinner);
        el.appendChild(document.createTextNode('\u00A0Enregistrement\u2026'));
    } else if (state === 'saved') {
        el.classList.add('autosave-saved');
        el.textContent = '✔ Enregistré';
        setTimeout(() => el.classList.remove('autosave-saved'), 3000);
    } else if (state === 'error') {
        el.classList.add('autosave-error');
        el.textContent = '✖ Erreur d\'enregistrement';
    }
}

// ========================================
// ENREGISTREMENT AUTOMATIQUE - RÉUNION
// ========================================

const REUNION_TABLE_IDS = ['reunion-odj-table', 'reunion-echeance-table', 'reunion-expired-table'];
let reunionAutoSaveTimer = null;
let reunionAutoSaveInFlight = false;

function reunionResultsContains(node) {
    return REUNION_TABLE_IDS.some(id => document.getElementById(id)?.contains(node));
}

function handleReunionAutoSaveEvent(e) {
    if (!reunionResultsContains(e.target)) return;

    if (e.type === 'input') {
        // Frappe en cours : repousser (mais pas amorcer) un enregistrement
        // pour ne pas reconstruire le tableau pendant l'édition.
        if (reunionAutoSaveTimer !== null) scheduleReunionAutoSave();
        return;
    }

    if (e.type === 'change') {
        if (e.target.contentEditable === 'true') return;
        scheduleReunionAutoSave();
        return;
    }

    // focusout : cellules contenteditable, uniquement quand le focus a
    // réellement quitté la zone d'édition (pas un passage de cellule à cellule).
    if (e.type === 'focusout') {
        if (e.target.contentEditable !== 'true') return;
        setTimeout(() => {
            if (!reunionResultsContains(document.activeElement)) {
                scheduleReunionAutoSave();
            }
        }, 0);
    }
}

function scheduleReunionAutoSave() {
    clearTimeout(reunionAutoSaveTimer);
    reunionAutoSaveTimer = setTimeout(() => {
        reunionAutoSaveTimer = null;
        void performReunionAutoSave();
    }, 800);
}

/** Force l'exécution d'un enregistrement Réunion en attente (changement d'onglet). */
async function flushReunionAutoSave() {
    if (reunionAutoSaveTimer === null) return;
    clearTimeout(reunionAutoSaveTimer);
    reunionAutoSaveTimer = null;
    await performReunionAutoSave();
}

async function performReunionAutoSave() {
    if (reunionAutoSaveInFlight) return;

    const hasRows = REUNION_TABLE_IDS.some(id =>
        document.getElementById(id)?.querySelector('table tbody tr')
    );
    if (!hasRows) return;

    reunionAutoSaveInFlight = true;
    showReunionSaveStatus('saving');
    try {
        await saveReunionModifications();
    } finally {
        reunionAutoSaveInFlight = false;
    }
}

function showReunionSaveStatus(state) {
    const el = document.getElementById('reunion-autosave-status');
    if (!el) return;

    el.classList.remove('autosave-saving', 'autosave-saved', 'autosave-error');
    while (el.firstChild) el.removeChild(el.firstChild);

    if (state === 'saving') {
        el.classList.add('autosave-saving');
        const spinner = document.createElement('span');
        spinner.className = 'autosave-spinner';
        spinner.textContent = '⟳';
        spinner.setAttribute('aria-hidden', 'true');
        el.appendChild(spinner);
        el.appendChild(document.createTextNode('\u00A0Enregistrement\u2026'));
    } else if (state === 'saved') {
        el.classList.add('autosave-saved');
        el.textContent = '✔ Enregistré';
        setTimeout(() => {
            el.classList.remove('autosave-saved');
        }, 3000);
    } else if (state === 'error') {
        el.classList.add('autosave-error');
        el.textContent = '✖ Erreur d\'enregistrement';
    }
}

async function saveReunionModifications() {
    try {
        const tables = ['reunion-odj-table', 'reunion-echeance-table', 'reunion-expired-table'];
        const updateActions = [];
        const newDates = new Set(); // Pour collecter les nouvelles dates à ajouter à l'Agenda
        const dossiersASupprimer = []; // Dossiers marqués « Supprimer le dossier »

        // Récupérer la date de réunion sélectionnée dans le sélecteur
        const reunionDateSelect = document.getElementById('reunion-date-select');
        const reunionDateValue = reunionDateSelect ? Number.parseFloat(reunionDateSelect.value) : null;

        for (const tableId of tables) {
            const container = document.getElementById(tableId);
            if (!container || !container.querySelector('table')) continue;

            const rows = container.querySelectorAll('table tbody tr');

            for (const row of rows) {
                const dossierId = Number.parseInt(row.dataset.dossierId);
                const dossierData = tablesData.ODJ.find(d => d.id === dossierId);
                if (!dossierData) continue;

                // Lecture des valeurs éditées, repérage par data-col.
                // La cellule "echeance" n'existe que pour reunion-odj-table et
                // reunion-expired-table ; sinon readEditableRow retombe sur
                // l'échéance d'origine du dossier.
                const edited = readEditableRow(row, dossierData);
                const nouveauDossier = edited.nomCellule ?? dossierData.Dossier;
                const nouveauxPorteurs = edited.porteurs;
                const actions = edited.actions;
                const nouvelleEcheance = edited.echeance;
                const nouvelleDateReunion = edited.dateReunion;
                const nouvelEtat = edited.changementEtat;

                if (nouvelEtat === etatRoles.supprimer) {
                    // Collecté pour suppression groupée après confirmation (tout l'historique)
                    dossiersASupprimer.push({ dossier: dossierData, nom: nouveauDossier || dossierData.Dossier });
                    continue;
                }

                if (nouvelEtat) {
                    // Ajouter une nouvelle ligne avec le nouvel état à la date de la réunion sélectionnée
                    const dateChangement = (reunionDateValue && !Number.isNaN(reunionDateValue))
                        ? reunionDateValue
                        : Math.floor(Date.now() / 1000);

                    // Collecter cette date pour l'ajouter à l'Agenda (sécurité OWASP)
                    newDates.add(dateChangement);

                    const existingToday = findTodayRecord(dossierData, dateChangement);
                    if (existingToday) {
                        updateActions.push(['UpdateRecord', 'ODJ', existingToday.id, {
                            Date_de_la_reunion: dateChangement,
                            Dossier: nouveauDossier || dossierData.Dossier,
                            ID_Dossier: dossierData.ID_Dossier || '',
                            Porteur_s_: ['L', ...nouveauxPorteurs],
                            Actions_a_mettre_en_uvre_etapes: actions,
                            Echeance: nouvelleEcheance,
                            Etat: nouvelEtat,
                            Enregistrement: Date.now() / 1000
                        }]);
                    } else {
                        updateActions.push(['AddRecord', 'ODJ', null, {
                            Date_de_la_reunion: dateChangement,
                            Dossier: nouveauDossier || dossierData.Dossier,
                            ID_Dossier: dossierData.ID_Dossier || '',
                            Porteur_s_: ['L', ...nouveauxPorteurs],
                            Actions_a_mettre_en_uvre_etapes: actions,
                            Echeance: nouvelleEcheance,
                            Etat: nouvelEtat,
                            Enregistrement: Date.now() / 1000
                        }]);
                    }
                }

                // Toujours mettre à jour la ligne existante
                updateActions.push(['UpdateRecord', 'ODJ', dossierId, {
                    Date_de_la_reunion: nouvelleDateReunion,
                    Dossier: nouveauDossier,
                    Porteur_s_: ['L', ...nouveauxPorteurs],
                    Actions_a_mettre_en_uvre_etapes: actions,
                    Echeance: nouvelleEcheance
                }]);
            }
        }

        // Suppression de dossiers : confirmation unique puis suppression de tout l'historique
        const { status: deletionStatus, deletedIds } = await confirmAndDeleteDossiers(dossiersASupprimer);
        if (deletionStatus === 'cancelled') {
            showReunionSaveStatus('idle');
            const selDate = document.getElementById('reunion-date-select');
            if (selDate && selDate.value) reunionDisplayData();
            return;
        }

        // Ne pas mettre à jour des lignes supprimées (marquées ou en cascade)
        const actionsToApply = updateActions.filter(([actionType, , recordId]) =>
            !(actionType === 'UpdateRecord' && deletedIds.has(recordId))
        );

        // S'assurer que toutes les nouvelles dates existent dans l'Agenda avant d'appliquer les actions
        for (const date of newDates) {
            await ensureAgendaDateExists(date);
        }

        // Appliquer toutes les modifications en une seule action
        if (actionsToApply.length > 0) {
            await grist.docApi.applyUserActions(actionsToApply);
        }

        // Pas de removeDuplicateRecords : l'upsert par jour évite déjà les
        // doublons et une déduplication globale effacerait des lignes
        // d'historique légitimes (retour à un état antérieur).
        await loadAllTables();
        populateConsultSelectors();

        // Rafraîchir l'affichage (conserve la date consultée ; les dossiers
        // supprimés disparaissent, plus besoin de recharger le widget).
        refreshReunionView();

        showReunionSaveStatus('saved');

    } catch (error) {
        console.error('Erreur lors de la sauvegarde:', error);
        showReunionSaveStatus('error');
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
    document.body.innerHTML = '<div class="container"><p class="no-results">Erreur : Widget doit être utilisé dans Grist</p></div>';
}