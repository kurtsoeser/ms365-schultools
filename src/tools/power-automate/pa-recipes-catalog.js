(function (global) {
    'use strict';

    /**
     * Katalog der Power-Automate-Rezepte als eigene ausführbare Tools.
     * Freistellungen hat ein eigenes Setup mit ZIP-Paket und steht hier nur als Link.
     */
    const RECIPES = [
        {
            id: 'termine-sync',
            toolId: 'pa-termine-sync',
            title: 'Schultermine → Kalender',
            icon: 'bi-calendar-event',
            kicker: 'Power Automate',
            summary:
                'SharePoint-Liste „Schultermine“ mit Outlook synchronisieren (Schul- und/oder Lehrerkalender).',
            helpId: 'tool-pa-termine-sync',
            storageKey: 'ms365-pa-termine-sync-v1',
            related: [
                { href: 'sharepoint-liste-schultermine.html', label: 'Schultermine-Liste', icon: 'bi-calendar3' },
                { href: 'https://make.powerautomate.com/', label: 'Power Automate', icon: 'bi-box-arrow-up-right', external: true }
            ],
            fields: [
                {
                    id: 'siteUrl',
                    label: 'SharePoint-Website (mit Terminliste)',
                    type: 'url',
                    placeholder: 'https://schule.sharepoint.com/sites/Intranet'
                },
                {
                    id: 'listName',
                    label: 'Listenname',
                    type: 'text',
                    defaultValue: 'Schultermine'
                },
                {
                    id: 'mailboxSchool',
                    label: 'Schulkalender-Mailbox (nur dieser Flow)',
                    type: 'email',
                    placeholder: 'kalender@schule.at'
                },
                {
                    id: 'mailboxTeachers',
                    label: 'Lehrerkalender-Mailbox (optional, nur dieser Flow)',
                    type: 'email',
                    placeholder: 'lehrer-kalender@schule.at'
                }
            ],
            list: null,
            listHint:
                'Liste zuerst mit dem Schultermine-Tool anlegen. Kalender-Postfächer nur hier eintragen – nicht das Freistellungs-Postfach wiederverwenden, außer die Schule will das bewusst so.',
            listActionLabel: null,
            steps: [
                'Trigger: Wenn ein Element erstellt oder geändert wird (SharePoint → Ihre Terminliste).',
                'Bedingung: SyncStatus ist leer oder „pending“.',
                'Aktion: Ereignis erstellen (Outlook) – ganztägig aus AllDay, sonst Beginn/Ende; Mailbox = Schulkalender.',
                'Optional: zweite Verzweigung für Ziel „Lehrer“ (eigene Mailbox).',
                'Zurückschreiben: OutlookEventID, SyncStatus=ok, LastSync; bei Fehler SyncError füllen.'
            ],
            doneLabel: 'Flow in unserer Schule aktiv'
        },
        {
            id: 'antraege',
            toolId: 'pa-antraege',
            title: 'Forms → Anträge',
            icon: 'bi-ui-checks',
            kicker: 'Power Automate',
            summary: 'Microsoft Forms → SharePoint-Liste „Anträge“ → Genehmigung Direktion/Verwaltung.',
            helpId: 'tool-pa-antraege',
            storageKey: 'ms365-pa-antraege-v1',
            related: [
                { href: 'https://forms.office.com/', label: 'Microsoft Forms', icon: 'bi-box-arrow-up-right', external: true },
                { href: 'https://make.powerautomate.com/', label: 'Power Automate', icon: 'bi-box-arrow-up-right', external: true }
            ],
            fields: [
                {
                    id: 'siteUrl',
                    label: 'SharePoint-Website',
                    type: 'url',
                    placeholder: 'https://schule.sharepoint.com/sites/Administration'
                },
                {
                    id: 'listName',
                    label: 'Listenname',
                    type: 'text',
                    defaultValue: 'Anträge'
                },
                {
                    id: 'approverEmail',
                    label: 'Genehmiger (E-Mail)',
                    type: 'email',
                    placeholder: 'verwaltung@schule.at'
                },
                {
                    id: 'formHint',
                    label: 'Forms-Titel (Hinweis)',
                    type: 'text',
                    placeholder: 'z. B. Raumantrag / Exkursion'
                }
            ],
            list: {
                defaultName: 'Anträge',
                description: 'Anträge aus Microsoft Forms (Genehmigung)',
                columns: [
                    {
                        name: 'Status',
                        displayName: 'Status',
                        choice: {
                            allowTextEntry: false,
                            choices: ['Neu', 'In Prüfung', 'Genehmigt', 'Abgelehnt']
                        }
                    },
                    {
                        name: 'Kategorie',
                        displayName: 'Kategorie',
                        choice: {
                            allowTextEntry: true,
                            choices: ['Raum', 'Exkursion', 'Gerät', 'Sonstiges']
                        }
                    },
                    {
                        name: 'Antragsteller',
                        displayName: 'Antragsteller',
                        personOrGroup: { allowMultipleSelection: false, chooseFromType: 'peopleOnly' }
                    },
                    {
                        name: 'Beschreibung',
                        displayName: 'Beschreibung',
                        text: { allowMultipleLines: true, maxLength: 8000 }
                    },
                    {
                        name: 'Beginn',
                        displayName: 'Beginn',
                        dateTime: { displayAs: 'default', format: 'dateTime' }
                    },
                    {
                        name: 'Ende',
                        displayName: 'Ende',
                        dateTime: { displayAs: 'default', format: 'dateTime' }
                    },
                    {
                        name: 'Bemerkungen',
                        displayName: 'Bemerkungen',
                        text: { allowMultipleLines: true, maxLength: 4000 }
                    }
                ]
            },
            listActionLabel: 'Antragsliste anlegen',
            steps: [
                'Microsoft Forms für den Antrag anlegen (Felder passend zur Liste).',
                'Flow: Bei neuer Forms-Antwort → Element in der Antragsliste anlegen (Status = Neu).',
                'Genehmigung (Approve/Reject) an die hinterlegte Adresse.',
                'Bei Genehmigung/Ablehnung: Status setzen; optional Kalender oder Teams-Nachricht.'
            ],
            doneLabel: 'Flow in unserer Schule aktiv'
        },
        {
            id: 'seminar',
            toolId: 'pa-seminar',
            title: 'Seminar / Fortbildung',
            icon: 'bi-mortarboard',
            kicker: 'Power Automate',
            summary: 'Anmeldeliste mit Platzkontingent, Warteliste und Bestätigungsmail.',
            helpId: 'tool-pa-seminar',
            storageKey: 'ms365-pa-seminar-v1',
            related: [
                { href: 'https://make.powerautomate.com/', label: 'Power Automate', icon: 'bi-box-arrow-up-right', external: true }
            ],
            fields: [
                {
                    id: 'siteUrl',
                    label: 'SharePoint-Website',
                    type: 'url',
                    placeholder: 'https://schule.sharepoint.com/sites/Lehrer'
                },
                {
                    id: 'listName',
                    label: 'Listenname',
                    type: 'text',
                    defaultValue: 'SeminarAnmeldungen'
                },
                {
                    id: 'maxSeats',
                    label: 'Standard-Maximalplätze',
                    type: 'number',
                    defaultValue: '20'
                },
                {
                    id: 'notifyEmail',
                    label: 'Organisator (E-Mail)',
                    type: 'email',
                    placeholder: 'fortbildung@schule.at'
                }
            ],
            list: {
                defaultName: 'SeminarAnmeldungen',
                description: 'Anmeldungen zu Seminaren und Fortbildungen',
                columns: [
                    {
                        name: 'Seminar',
                        displayName: 'Seminar',
                        text: { allowMultipleLines: false, maxLength: 255 }
                    },
                    {
                        name: 'Teilnehmer',
                        displayName: 'Teilnehmer',
                        personOrGroup: { allowMultipleSelection: false, chooseFromType: 'peopleOnly' }
                    },
                    {
                        name: 'Status',
                        displayName: 'Status',
                        choice: {
                            allowTextEntry: false,
                            choices: ['Angemeldet', 'Warteliste', 'Bestätigt', 'Abgesagt']
                        }
                    },
                    {
                        name: 'MaxPlaetze',
                        displayName: 'MaxPlätze',
                        number: {}
                    },
                    {
                        name: 'Hinweis',
                        displayName: 'Hinweis',
                        text: { allowMultipleLines: true, maxLength: 2000 }
                    }
                ]
            },
            listActionLabel: 'Anmeldeliste anlegen',
            steps: [
                'Anmeldungen per Forms oder direkt in der Liste erfassen.',
                'Flow bei neuem Element: Anzahl „Angemeldet/Bestätigt“ zählen; bei Überbuchung Status „Warteliste“.',
                'E-Mail an Teilnehmer:in; optional Adaptive Card im Lehrer-Team.',
                'Organisator bei Warteliste oder Absage informieren.'
            ],
            doneLabel: 'Flow in unserer Schule aktiv'
        },
        {
            id: 'gast-erinnerung',
            toolId: 'pa-gast-erinnerung',
            title: 'Gast-Erinnerung',
            icon: 'bi-person-exclamation',
            kicker: 'Power Automate',
            summary: 'Periodisch Gäste ohne Aktivität melden – ergänzt das Gäste-Tool.',
            helpId: 'tool-pa-gast',
            storageKey: 'ms365-pa-gast-v1',
            related: [
                { href: 'gaeste-verwalten.html', label: 'Gäste verwalten', icon: 'bi-shield-check' },
                { href: 'https://make.powerautomate.com/', label: 'Power Automate', icon: 'bi-box-arrow-up-right', external: true }
            ],
            fields: [
                {
                    id: 'itEmail',
                    label: 'IT / Empfänger der Erinnerung',
                    type: 'email',
                    placeholder: 'it@schule.at'
                },
                {
                    id: 'schedule',
                    label: 'Rhythmus',
                    type: 'text',
                    defaultValue: 'monatlich',
                    placeholder: 'z. B. monatlich / quartalsweise'
                },
                {
                    id: 'inactiveDays',
                    label: 'Inaktiv seit (Tage) – Hinweis für Flow',
                    type: 'number',
                    defaultValue: '90'
                }
            ],
            list: null,
            listHint: 'Keine Liste nötig. Gäste prüfen Sie im Gäste-Tool; der Flow meldet periodisch an die IT.',
            listActionLabel: null,
            steps: [
                'Geplanten Flow anlegen (Recurrence laut Rhythmus).',
                'Gäste auflisten (Office 365 Users / Graph) oder CSV aus dem Gäste-Tool als Ausgangspunkt nutzen.',
                'Filter: lange inaktiv / nie angemeldet (Schwellenwert aus den Einstellungen).',
                'E-Mail an IT mit Liste „prüfen / löschen“.'
            ],
            doneLabel: 'Flow in unserer Schule aktiv'
        },
        {
            id: 'diplom-ordner',
            toolId: 'pa-diplom-ordner',
            title: 'Diplom-Ordnerstruktur',
            icon: 'bi-folder-plus',
            kicker: 'Power Automate',
            summary: 'Bei neuem Team mit Alias dipl-* Vorlagenordner und Willkommensnachricht anlegen.',
            helpId: 'tool-pa-diplom',
            storageKey: 'ms365-pa-diplom-v1',
            related: [
                { href: 'diplomarbeiten.html', label: 'Diplomarbeiten-Tool', icon: 'bi-mortarboard' },
                { href: 'https://make.powerautomate.com/', label: 'Power Automate', icon: 'bi-box-arrow-up-right', external: true }
            ],
            fields: [
                {
                    id: 'prefix',
                    label: 'Alias-Präfix',
                    type: 'text',
                    defaultValue: 'dipl-'
                },
                {
                    id: 'folders',
                    label: 'Ordner (kommagetrennt)',
                    type: 'text',
                    defaultValue: 'Exposé, Rohdaten, Abgabe, Feedback'
                },
                {
                    id: 'welcomeText',
                    label: 'Willkommenstext (Kanal General)',
                    type: 'text',
                    defaultValue: 'Willkommen im Diplomarbeitsteam – bitte die Ordnerstruktur nutzen.'
                }
            ],
            list: null,
            listHint: 'Teams legen Sie im Diplomarbeiten-Tool an. Hier die Flow-Parameter für die Ordnerstruktur.',
            listActionLabel: null,
            steps: [
                'Trigger: Wenn ein Team erstellt wird (oder manuell mit Gruppen-ID).',
                'Bedingung: mailNickname beginnt mit dem Präfix (z. B. dipl-).',
                'In der Dokumentenbibliothek Ordner aus der Liste anlegen.',
                'Nachricht im Kanal General mit dem Willkommenstext posten.'
            ],
            doneLabel: 'Flow in unserer Schule aktiv'
        },
        {
            id: 'schilf',
            toolId: 'pa-schilf',
            title: 'Schilf: Kalender & Favoriten',
            icon: 'bi-bookmark-star',
            kicker: 'Checkliste',
            summary: 'Kein Flow – Prozess für Lehrkräfte nach Fortbildung (Kalender, Favoriten, Zielgruppen).',
            helpId: 'tool-pa-schilf',
            storageKey: 'ms365-pa-schilf-v1',
            related: [
                { href: 'sharepoint-intranet-hub.html', label: 'Schul-Intranet', icon: 'bi-house-door' },
                { href: 'sharepoint-liste-schultermine.html', label: 'Schultermine', icon: 'bi-calendar3' }
            ],
            fields: [
                {
                    id: 'calendarName',
                    label: 'Name des Schulkalenders',
                    type: 'text',
                    placeholder: 'z. B. HAK Termine'
                },
                {
                    id: 'intranetUrl',
                    label: 'Intranet-URL',
                    type: 'url',
                    placeholder: 'https://schule.sharepoint.com/sites/Intranet'
                }
            ],
            list: null,
            listHint: 'Reine Checkliste – zum Durchgehen in der Schilf und Abhaken.',
            listActionLabel: null,
            checklist: [
                'Outlook/Teams: Kalender hinzufügen → aus Verzeichnis (Schulkalender).',
                'Favoriten: Intranet, Formulare, Lehrerliste pinnen.',
                'Zielgruppen erklären: öffentliche News vs. Kollegiums-Kanal.',
                'Kurzhandout oder Folie für die nächste Schilf vorbereiten.'
            ],
            steps: [],
            doneLabel: 'In Schilf kommuniziert'
        },
        {
            id: 'schularbeiten-mail',
            toolId: 'pa-schularbeiten-mail',
            title: 'Schularbeiten Status-Mail',
            icon: 'bi-envelope-check',
            kicker: 'Power Automate',
            summary:
                'Bei Statusänderung in der Liste „Schularbeiten“ eine E-Mail an die antragstellende Lehrkraft senden (fixiert/abgelehnt).',
            helpId: 'tool-pa-schularbeiten-mail',
            storageKey: 'ms365-pa-schularbeiten-mail-v1',
            related: [
                { href: 'schularbeiten-planer.html', label: 'Schularbeiten-Planer', icon: 'bi-journal-check' },
                { href: 'sharepoint-liste-schularbeiten.html', label: 'Schularbeiten-Listen', icon: 'bi-list-ul' },
                {
                    href: 'https://make.powerautomate.com/',
                    label: 'Power Automate',
                    icon: 'bi-box-arrow-up-right',
                    external: true
                }
            ],
            fields: [
                {
                    id: 'siteUrl',
                    label: 'SharePoint-Website (mit Schularbeiten-Liste)',
                    type: 'url',
                    placeholder: 'https://schule.sharepoint.com/sites/Intranet'
                },
                {
                    id: 'listName',
                    label: 'Listenname',
                    type: 'text',
                    defaultValue: 'Schularbeiten'
                },
                {
                    id: 'fromMailbox',
                    label: 'Absender-Mailbox (optional, freigegeben)',
                    type: 'email',
                    placeholder: 'direktion@schule.at'
                }
            ],
            list: null,
            listHint:
                'Liste zuerst mit dem Schularbeiten-Listen-Tool anlegen. Empfänger = Feld LehrerEmail oder BeantragtVon (UPN).',
            listActionLabel: null,
            steps: [
                'Trigger: Wenn ein Element erstellt oder geändert wird (SharePoint → Liste „Schularbeiten“).',
                'Bedingung: Status ist „fixiert“ oder „abgelehnt“ (und ggf. vorheriger Status ≠ neuer Status).',
                'Aktion: E-Mail senden (Office 365 Outlook) an LehrerEmail; Fallback BeantragtVon.',
                'Betreff z. B.: „Schularbeit [Status]: [Title] – [KlasseCode]“.',
                'Text: Datum, FachCode, Thema, AblehnungsGrund (wenn abgelehnt), Link zum Planer.',
                'Optional: nur bei geändertem Status (Get changes / Version vergleichen).'
            ],
            doneLabel: 'Status-Mail-Flow aktiv'
        },
        {
            id: 'projektwochen-mail',
            toolId: 'pa-projektwochen-mail',
            title: 'Projektwochen Status-Mail',
            icon: 'bi-envelope-heart',
            kicker: 'Power Automate',
            summary:
                'Bei Freigabe oder Ablehnung in der Liste „PW-Angebote“ eine E-Mail an die antragstellende Lehrkraft senden.',
            helpId: 'tool-pa-projektwochen-mail',
            storageKey: 'ms365-pa-projektwochen-mail-v1',
            related: [
                { href: 'projektwochen.html', label: 'Projektwochen', icon: 'bi-calendar2-week' },
                { href: 'sharepoint-liste-projektwochen.html', label: 'Projektwochen-Listen', icon: 'bi-list-ul' },
                {
                    href: 'https://make.powerautomate.com/',
                    label: 'Power Automate',
                    icon: 'bi-box-arrow-up-right',
                    external: true
                }
            ],
            fields: [
                {
                    id: 'siteUrl',
                    label: 'SharePoint-Website (mit PW-Angebote)',
                    type: 'url',
                    placeholder: 'https://schule.sharepoint.com/sites/Intranet'
                },
                {
                    id: 'listName',
                    label: 'Listenname',
                    type: 'text',
                    defaultValue: 'PW-Angebote'
                },
                {
                    id: 'fromMailbox',
                    label: 'Absender-Mailbox (optional, freigegeben)',
                    type: 'email',
                    placeholder: 'direktion@schule.at'
                }
            ],
            list: null,
            listHint:
                'Liste zuerst mit dem Projektwochen-Listen-Tool anlegen. Empfänger = LehrerEmail oder BeantragtVon.',
            listActionLabel: null,
            steps: [
                'Trigger: Wenn ein Element erstellt oder geändert wird (SharePoint → Liste „PW-Angebote“).',
                'Bedingung: Status ist „freigegeben“ oder „abgelehnt“.',
                'Aktion: E-Mail senden (Office 365 Outlook) an LehrerEmail; Fallback BeantragtVon.',
                'Betreff z. B.: „Projektwochen-Angebot [Status]: [Title] – [Datum]“.',
                'Text: Slot, Ort, Kapazität, Preis, AblehnungsGrund (wenn abgelehnt), optional BookingsBookingUrl.',
                'Optional: nur bei geändertem Status (Get changes / Version vergleichen).'
            ],
            doneLabel: 'Projektwochen Status-Mail-Flow aktiv'
        }
    ];

    const HUB_EXTRA = {
        id: 'freistellung',
        title: 'Freistellungen (KV + Direktion)',
        href: 'freistellung-setup.html',
        icon: 'bi-person-check',
        summary: 'Liste + parametrierbares Flow-Paket für Ziel-Tenants.'
    };

    function getById(id) {
        const key = String(id || '').trim();
        for (let i = 0; i < RECIPES.length; i++) {
            if (RECIPES[i].id === key || RECIPES[i].toolId === key) return RECIPES[i];
        }
        return null;
    }

    global.ms365PaRecipes = {
        list: RECIPES,
        hubExtra: HUB_EXTRA,
        getById: getById
    };
})(window);
