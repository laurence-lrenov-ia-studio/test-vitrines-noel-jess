/**
 * ============================================================
 * VITRINES DE NOËL — BACKEND 2026
 * ============================================================
 *
 * Google Apps Script lié au Google Sheet.
 *
 * FRONT :
 * - page client hébergée sur Netlify
 * - future interface Jess hébergée sur GitHub Pages
 *
 * RÈGLES MÉTIER :
 * - le client peut sélectionner plusieurs week-ends ;
 * - une demande ne bloque pas de place ;
 * - seule une demande CONFIRMÉE consomme une place ;
 * - Jess choisit ensuite le week-end retenu + samedi/dimanche ;
 * - statuts :
 *   NOUVELLE
 *   DEVIS ENVOYÉ
 *   CONFIRMÉE
 *   REFUSÉE
 * - capacité : 2 interventions par week-end ;
 * - demandes closes le 6 novembre 2026 à 23:59.
 */


const VN = {

  SHEETS: {
    PARAMS: 'PARAMETRES',
    PERIODS: 'PERIODES',
    BOOKINGS: 'RESERVATIONS'
  },

  STATUS: {
    NEW: 'NOUVELLE',
    QUOTE: 'DEVIS ENVOYÉ',
    CONFIRMED: 'CONFIRMÉE',
    REFUSED: 'REFUSÉE'
  }

};


/* ============================================================
   MENU
   ============================================================ */

function onOpen() {

  SpreadsheetApp
    .getUi()
    .createMenu('Vitrines de Noël')
    .addItem(
      'Initialiser / vérifier',
      'setupVitrinesNoel'
    )
    .addItem(
      'Afficher la clé admin',
      'showAdminKey'
    )
    .addItem(
      'Autoriser les notifications Gmail',
      'autoriserNotificationsAdmin'
    )
    .addToUi();

}


/* ============================================================
   INITIALISATION
   ============================================================ */

function setupVitrinesNoel() {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();


  if (!ss) {

    throw new Error(
      'Ouvre ce script depuis le Google Sheet Vitrines de Noël.'
    );

  }


  PropertiesService
    .getScriptProperties()
    .setProperty(
      'SPREADSHEET_ID',
      ss.getId()
    );


  setupParams_(
    getOrCreateSheet_(
      ss,
      VN.SHEETS.PARAMS
    )
  );


  setupPeriods_(
    getOrCreateSheet_(
      ss,
      VN.SHEETS.PERIODS
    )
  );


  setupBookings_(
    getOrCreateSheet_(
      ss,
      VN.SHEETS.BOOKINGS
    )
  );


  SpreadsheetApp
    .getUi()
    .alert(
      'Vitrines de Noël',
      'Initialisation terminée.\n\nVérifie NOTIFICATION_EMAIL dans l’onglet PARAMETRES.',
      SpreadsheetApp
        .getUi()
        .ButtonSet
        .OK
    );

}


function getOrCreateSheet_(
  ss,
  name
) {

  return (
    ss.getSheetByName(name) ||
    ss.insertSheet(name)
  );

}


/* ============================================================
   PARAMÈTRES
   ============================================================ */

function setupParams_(sheet) {

  const old =
    readParamsFromSheet_(sheet);


  const adminKey =
    old.ADMIN_KEY ||
    Utilities
      .getUuid()
      .replace(/-/g, '') +
    Utilities
      .getUuid()
      .replace(/-/g, '')
      .slice(0, 12);


  const rows = [

    [
      'CLE',
      'VALEUR',
      'DESCRIPTION'
    ],

    [
      'APP_NAME',
      old.APP_NAME ||
      'Vitrines de Noël',
      'Nom de l’outil'
    ],

    [
      'DATE_FERMETURE',
      old.DATE_FERMETURE ||
      new Date(
        2026,
        10,
        6,
        23,
        59,
        59
      ),
      'Fin des demandes client'
    ],

    [
      'NOTIFICATION_EMAIL',
      old.NOTIFICATION_EMAIL ||
      'A_RENSEIGNER',
      'Email qui reçoit les nouvelles demandes'
    ],

    [
      'CLIENT_CONFIRMATION_EMAIL',
      old.CLIENT_CONFIRMATION_EMAIL ||
      'OUI',
      'OUI / NON'
    ],

    [
      'TIMEZONE',
      old.TIMEZONE ||
      'Europe/Paris',
      'Fuseau horaire'
    ],

    [
      'ADMIN_KEY',
      adminKey,
      'Clé privée de l’interface Jess'
    ],

    [
      'SUCCESS_MESSAGE',
      old.SUCCESS_MESSAGE ||
      'Votre demande est bien enregistrée. On revient vers vous très vite pour confirmer votre intervention et vous transmettre les éléments nécessaires à la suite.',
      'Message après envoi'
    ]

  ];


  sheet.clear();


  sheet
    .getRange(
      1,
      1,
      rows.length,
      3
    )
    .setValues(rows);


  sheet
    .getRange(
      3,
      2
    )
    .setNumberFormat(
      'dd/mm/yyyy hh:mm'
    );


  formatHeader_(
    sheet.getRange(
      1,
      1,
      1,
      3
    )
  );


  sheet.setFrozenRows(1);

  sheet.setColumnWidth(
    1,
    240
  );

  sheet.setColumnWidth(
    2,
    380
  );

  sheet.setColumnWidth(
    3,
    520
  );

}


/* ============================================================
   PÉRIODES
   ============================================================ */

function setupPeriods_(sheet) {

  let existing = [];


  if (
    sheet.getLastRow() > 1 &&
    sheet.getLastColumn() >= 8
  ) {

    existing =
      sheet
        .getRange(
          2,
          1,
          sheet.getLastRow() - 1,
          8
        )
        .getValues()
        .filter(
          r => clean_(r[0])
        );

  }


  sheet.clear();


  sheet
    .getRange(
      1,
      1,
      1,
      8
    )
    .setValues([[
      'ID_PERIODE',
      'LIBELLE',
      'DATE_SAMEDI',
      'DATE_DIMANCHE',
      'CAPACITE',
      'OUVERTE',
      'ORDRE',
      'NOTE'
    ]]);


  formatHeader_(
    sheet.getRange(
      1,
      1,
      1,
      8
    )
  );


  if (existing.length) {

    sheet
      .getRange(
        2,
        1,
        existing.length,
        8
      )
      .setValues(existing);

  }

  else {

    sheet
      .getRange(
        2,
        1,
        3,
        8
      )
      .setValues([

        [
          'WE-1415',
          '14–15 novembre',
          new Date(
            2026,
            10,
            14
          ),
          new Date(
            2026,
            10,
            15
          ),
          2,
          'OUI',
          1,
          ''
        ],

        [
          'WE-2122',
          '21–22 novembre',
          new Date(
            2026,
            10,
            21
          ),
          new Date(
            2026,
            10,
            22
          ),
          2,
          'OUI',
          2,
          ''
        ],

        [
          'WE-2829',
          '28–29 novembre',
          new Date(
            2026,
            10,
            28
          ),
          new Date(
            2026,
            10,
            29
          ),
          2,
          'OUI',
          3,
          ''
        ]

      ]);

  }


  const yesNo =
    SpreadsheetApp
      .newDataValidation()
      .requireValueInList(
        [
          'OUI',
          'NON'
        ],
        true
      )
      .setAllowInvalid(false)
      .build();


  sheet
    .getRange(
      'F2:F100'
    )
    .setDataValidation(
      yesNo
    );


  sheet
    .getRange(
      'C2:D100'
    )
    .setNumberFormat(
      'dd/mm/yyyy'
    );


  sheet.setFrozenRows(1);

}


/* ============================================================
   RÉSERVATIONS
   ============================================================ */

function setupBookings_(sheet) {

  const headers = [[

    'HORODATAGE',
    'ID_DEMANDE',
    'ENTREPRISE',
    'CONTACT',
    'EMAIL',
    'TELEPHONE',
    'PRESTATIONS',

    'IDS_PERIODES_DEMANDEES',
    'PERIODES_DEMANDEES',

    'STATUT',

    'ID_PERIODE_RETENUE',
    'PERIODE_RETENUE',

    'JOUR_POSE',
    'DATE_DEPOSE',

    'COMMENTAIRE',

    'SOURCE'

  ]];


  let existing = [];


  /*
   * Ancienne version avec
   * colonne CLIENT_FIDELE.
   */
  if (
    sheet.getLastRow() > 1 &&
    sheet.getLastColumn() === 17
  ) {

    const oldRows =
      sheet
        .getRange(
          2,
          1,
          sheet.getLastRow() - 1,
          17
        )
        .getValues()
        .filter(
          r => r[0]
        );


    existing =
      oldRows.map(
        r => [

          r[0],
          r[1],
          r[2],
          r[3],
          r[4],
          r[5],
          r[6],
          r[7],
          r[8],
          r[9],
          r[10],
          r[11],
          r[12],
          r[13],
          r[14],
          r[16]

        ]
      );

  }


  else if (
    sheet.getLastRow() > 1 &&
    sheet.getLastColumn() === 16
  ) {

    existing =
      sheet
        .getRange(
          2,
          1,
          sheet.getLastRow() - 1,
          16
        )
        .getValues()
        .filter(
          r => r[0]
        );

  }


  sheet.clear();


  sheet
    .getRange(
      1,
      1,
      1,
      16
    )
    .setValues(headers);


  formatHeader_(
    sheet.getRange(
      1,
      1,
      1,
      16
    )
  );


  if (existing.length) {

    sheet
      .getRange(
        2,
        1,
        existing.length,
        16
      )
      .setValues(existing);

  }


  const statusValidation =
    SpreadsheetApp
      .newDataValidation()
      .requireValueInList(
        Object.values(
          VN.STATUS
        ),
        true
      )
      .setAllowInvalid(false)
      .build();


  const dayValidation =
    SpreadsheetApp
      .newDataValidation()
      .requireValueInList(
        [
          'SAMEDI',
          'DIMANCHE'
        ],
        true
      )
      .setAllowInvalid(true)
      .build();


  sheet
    .getRange(
      'J2:J5000'
    )
    .setDataValidation(
      statusValidation
    );


  sheet
    .getRange(
      'M2:M5000'
    )
    .setDataValidation(
      dayValidation
    );


  sheet
    .getRange(
      'A2:A5000'
    )
    .setNumberFormat(
      'dd/mm/yyyy hh:mm'
    );


  sheet
    .getRange(
      'N2:N5000'
    )
    .setNumberFormat(
      'dd/mm/yyyy'
    );


  sheet.setFrozenRows(1);

}


/* ============================================================
   FORMAT
   ============================================================ */

function formatHeader_(range) {

  range
    .setBackground(
      '#114838'
    )
    .setFontColor(
      '#FFFFFF'
    )
    .setFontWeight(
      'bold'
    );

}


/* ============================================================
   API GET
   ============================================================ */

function doGet(e) {

  try {

    const action =
      clean_(
        e &&
        e.parameter
          ? e.parameter.action
          : ''
      );


    if (
      action === 'health'
    ) {

      return json_({

        ok: true,

        app:
          'Vitrines de Noël',

        timestamp:
          new Date()
            .toISOString()

      });

    }


    if (
      action === 'bootstrap' ||
      !action
    ) {

      return json_({

        ok: true,

        data:
          getClientBootstrap()

      });

    }


    return json_({

      ok: false,

      error:
        'Action GET inconnue.'

    });

  }

  catch (err) {

    return jsonError_(err);

  }

}


/* ============================================================
   API POST
   ============================================================ */

function doPost(e) {

  try {

    const body =
      parsePostBody_(e);


    const action =
      clean_(
        body.action
      );


    const payload =
      body.payload || {};


    const adminKey =
      clean_(
        body.adminKey
      );


    switch (action) {


      case 'submitReservation':

        return json_({

          ok: true,

          data:
            submitReservation(
              payload
            )

        });


      case 'getAdminDashboard':

        return json_({

          ok: true,

          data:
            getAdminDashboard(
              adminKey
            )

        });


      case 'setSelectedPeriod':

        return json_({

          ok: true,

          data:
            setSelectedPeriod(

              adminKey,

              clean_(
                payload.requestId
              ),

              clean_(
                payload.periodId
              )

            )

        });


      case 'assignPoseDay':

        return json_({

          ok: true,

          data:
            assignPoseDay(

              adminKey,

              clean_(
                payload.requestId
              ),

              clean_(
                payload.day
              )

            )

        });


      case 'setRequestStatus':

        return json_({

          ok: true,

          data:
            setRequestStatus(

              adminKey,

              clean_(
                payload.requestId
              ),

              clean_(
                payload.status
              )

            )

        });


      case 'setPeriodOpen':

        return json_({

          ok: true,

          data:
            setPeriodOpen(

              adminKey,

              clean_(
                payload.periodId
              ),

              Boolean(
                payload.open
              )

            )

        });


      default:

        return json_({

          ok: false,

          error:
            'Action POST inconnue.'

        });

    }

  }

  catch (err) {

    return jsonError_(err);

  }

}


function parsePostBody_(e) {

  if (
    !e ||
    !e.postData ||
    !e.postData.contents
  ) {

    throw new Error(
      'Corps POST manquant.'
    );

  }


  try {

    return JSON.parse(
      e.postData.contents
    );

  }

  catch (err) {

    throw new Error(
      'Corps JSON invalide.'
    );

  }

}


function json_(data) {

  return ContentService
    .createTextOutput(
      JSON.stringify(data)
    )
    .setMimeType(
      ContentService
        .MimeType
        .JSON
    );

}


function jsonError_(err) {

  return json_({

    ok: false,

    error:
      err &&
      err.message
        ? err.message
        : String(err)

  });

}


/* ============================================================
   DONNÉES CLIENT
   ============================================================ */

function getClientBootstrap() {

  return {

    closed:
      isClientFormClosed_(),

    cutoff:
      formatDate_(
        getParam_(
          'DATE_FERMETURE'
        ),
        'dd/MM/yyyy'
      ),

    periods:
      getPublicPeriods(),

    successMessage:
      clean_(
        getParam_(
          'SUCCESS_MESSAGE'
        )
      )

  };

}


function getPublicPeriods() {

  const requests =
    requestCounts_();


  const confirmed =
    confirmedCounts_();


  return periodRows_()

    .map(
      r => {

        const id =
          clean_(
            r[0]
          );


        const capacity =
          Number(
            r[4]
          ) || 0;


        const confirmedCount =
          confirmed[id] || 0;


        const remaining =
          Math.max(
            capacity -
            confirmedCount,
            0
          );


        const requestCount =
          requests[id] || 0;


        return {

          id: id,

          label:
            clean_(
              r[1]
            ),

          saturday:
            formatDate_(
              r[2],
              'dd/MM/yyyy'
            ),

          sunday:
            formatDate_(
              r[3],
              'dd/MM/yyyy'
            ),

          capacity:
            capacity,

          confirmed:
            confirmedCount,

          remaining:
            remaining,

          requestCount:
            requestCount,

          highDemand:
            requestCount >= 4,

          open:
            clean_(
              r[5]
            )
              .toUpperCase()
              === 'OUI',

          full:
            remaining <= 0,

          order:
            Number(
              r[6]
            ) || 999

        };

      }
    )

    .filter(
      p =>
        p.id &&
        p.open
    )

    .sort(
      (
        a,
        b
      ) =>
        a.order -
        b.order
    );

}


/* ============================================================
   ENREGISTRER UNE DEMANDE
   ============================================================ */

function submitReservation(payload) {

  if (
    isClientFormClosed_()
  ) {

    throw new Error(
      'Les demandes de réservation sont maintenant closes.'
    );

  }


  const data =
    normalizeReservation_(
      payload
    );


  validateReservation_(
    data
  );


  const periods =
    data.periodIds
      .map(
        id =>
          getPeriod_(id)
      )
      .filter(Boolean);


  if (
    periods.length !==
    data.periodIds.length
  ) {

    throw new Error(
      'Une des dates sélectionnées n’est plus disponible.'
    );

  }


  const confirmed =
    confirmedCounts_();


  periods.forEach(
    period => {

      if (
        !period.open
      ) {

        throw new Error(
          'Une des dates sélectionnées n’est plus disponible.'
        );

      }


      if (
        (
          confirmed[
            period.id
          ] || 0
        )
        >=
        period.capacity
      ) {

        throw new Error(
          'Une des dates sélectionnées est maintenant complète.'
        );

      }

    }
  );


  const id =
    makeRequestId_();


  const labels =
    periods.map(
      p => p.label
    );


  getSheet_(
    VN.SHEETS.BOOKINGS
  )
    .appendRow([

      new Date(),

      id,

      data.company,

      data.contact,

      data.email,

      data.phone,

      data.services
        .join(' + '),

      data.periodIds
        .join('|'),

      labels
        .join(' | '),

      VN.STATUS.NEW,

      '',

      '',

      '',

      '',

      data.comment,

      'WEB'

    ]);


  /*
   * Notification interne Jess.
   * Un problème de notification
   * ne doit jamais annuler
   * une réservation déjà créée.
   */
  try {

    sendOwnerNotification_({

      id: id,

      company:
        data.company,

      contact:
        data.contact,

      email:
        data.email,

      phone:
        data.phone,

      services:
        data.services,

      periodLabels:
        labels,

      comment:
        data.comment

    });

  }

  catch (error) {

    console.error(
      'Erreur notification admin :',
      error
    );

  }


  /*
   * Accusé client.
   * Même principe :
   * la réservation reste enregistrée
   * même si Gmail rencontre un problème.
   */
  if (
    clean_(
      getParam_(
        'CLIENT_CONFIRMATION_EMAIL'
      )
    )
      .toUpperCase()
      === 'OUI'
  ) {

    try {

      sendClientAcknowledgement_({

        id: id,

        contact:
          data.contact,

        email:
          data.email,

        periodLabels:
          labels

      });

    }

    catch (error) {

      console.error(
        'Erreur email client :',
        error
      );

    }

  }


  return {

    id: id,

    message:
      clean_(
        getParam_(
          'SUCCESS_MESSAGE'
        )
      )

  };

}


function normalizeReservation_(p) {

  p = p || {};


  return {

    company:
      clean_(
        p.company
      ),

    contact:
      clean_(
        p.contact
      ),

    email:
      clean_(
        p.email
      )
        .toLowerCase(),

    phone:
      clean_(
        p.phone
      ),

    services:
      Array.isArray(
        p.services
      )
        ?
        [
          ...new Set(
            p.services
              .map(
                clean_
              )
              .filter(Boolean)
          )
        ]
        :
        [],

    periodIds:
      Array.isArray(
        p.periodIds
      )
        ?
        [
          ...new Set(
            p.periodIds
              .map(
                clean_
              )
              .filter(Boolean)
          )
        ]
        :
        [],

    comment:
      clean_(
        p.comment
      )

  };

}


function validateReservation_(d) {

  if (
    !d.company
  ) {

    throw new Error(
      'Le nom de l’entreprise est obligatoire.'
    );

  }


  if (
    !d.contact
  ) {

    throw new Error(
      'Le nom du contact est obligatoire.'
    );

  }


  if (
    !d.email ||
    !/^[^\s@]+@[^\s@]+\.[^\s@]+$/
      .test(
        d.email
      )
  ) {

    throw new Error(
      'Merci de renseigner une adresse email valide.'
    );

  }


  if (
    !d.services.length
  ) {

    throw new Error(
      'Merci de sélectionner au moins une prestation.'
    );

  }


  if (
    !d.periodIds.length
  ) {

    throw new Error(
      'Merci de sélectionner au moins un week-end.'
    );

  }

}


/* ============================================================
   INTERFACE JESS
   ============================================================ */

function getAdminDashboard(
  adminKey
) {

  requireAdmin_(
    adminKey
  );


  const periods =
    getAdminPeriods_();


  const bookings =
    getAdminBookings_();


  return {

    totals: {

      requests:
        bookings
          .filter(
            b =>
              b.status !==
              VN.STATUS.REFUSED
          )
          .length,

      quotes:
        bookings
          .filter(
            b =>
              b.status ===
              VN.STATUS.QUOTE
          )
          .length,

      confirmed:
        bookings
          .filter(
            b =>
              b.status ===
              VN.STATUS.CONFIRMED
          )
          .length,

      remaining:
        periods
          .filter(
            p =>
              p.open
          )
          .reduce(
            (
              total,
              p
            ) =>
              total +
              p.remaining,
            0
          )

    },

    periods:
      periods,

    bookings:
      bookings

  };

}


/* ============================================================
   WEEK-END RETENU
   ============================================================ */

function setSelectedPeriod(
  adminKey,
  requestId,
  periodId
) {

  requireAdmin_(
    adminKey
  );


  const row =
    findRequestRow_(
      requestId
    );


  if (!row) {

    throw new Error(
      'Demande introuvable.'
    );

  }


  const sheet =
    getSheet_(
      VN.SHEETS.BOOKINGS
    );


  const requestedIds =
    splitIds_(
      sheet
        .getRange(
          row,
          8
        )
        .getValue()
    );


  if (
    periodId &&
    !requestedIds
      .includes(
        periodId
      )
  ) {

    throw new Error(
      'Ce week-end ne fait pas partie des disponibilités du client.'
    );

  }


  const period =
    periodId
      ?
      getPeriod_(
        periodId
      )
      :
      null;


  sheet
    .getRange(
      row,
      11
    )
    .setValue(
      periodId || ''
    );


  sheet
    .getRange(
      row,
      12
    )
    .setValue(
      period
        ? period.label
        : ''
    );


  return getAdminDashboard(
    adminKey
  );

}


/* ============================================================
   JOUR DE POSE
   ============================================================ */

function assignPoseDay(
  adminKey,
  requestId,
  day
) {

  requireAdmin_(
    adminKey
  );


  day =
    clean_(
      day
    )
      .toUpperCase();


  if (
    ![
      '',
      'SAMEDI',
      'DIMANCHE'
    ]
      .includes(
        day
      )
  ) {

    throw new Error(
      'Jour invalide.'
    );

  }


  const row =
    findRequestRow_(
      requestId
    );


  if (!row) {

    throw new Error(
      'Demande introuvable.'
    );

  }


  getSheet_(
    VN.SHEETS.BOOKINGS
  )
    .getRange(
      row,
      13
    )
    .setValue(
      day
    );


  return getAdminDashboard(
    adminKey
  );

}


/* ============================================================
   CHANGEMENT DE STATUT
   ============================================================ */

function setRequestStatus(
  adminKey,
  requestId,
  newStatus
) {

  requireAdmin_(
    adminKey
  );


  if (
    !Object
      .values(
        VN.STATUS
      )
      .includes(
        newStatus
      )
  ) {

    throw new Error(
      'Statut invalide.'
    );

  }


  const sheet =
    getSheet_(
      VN.SHEETS.BOOKINGS
    );


  const row =
    findRequestRow_(
      requestId
    );


  if (!row) {

    throw new Error(
      'Demande introuvable.'
    );

  }


  if (
    newStatus ===
    VN.STATUS.CONFIRMED
  ) {

    const selectedPeriodId =
      clean_(
        sheet
          .getRange(
            row,
            11
          )
          .getValue()
      );


    const poseDay =
      clean_(
        sheet
          .getRange(
            row,
            13
          )
          .getValue()
      );


    if (
      !selectedPeriodId
    ) {

      throw new Error(
        'Choisis d’abord le week-end retenu.'
      );

    }


    if (
      !poseDay
    ) {

      throw new Error(
        'Choisis d’abord samedi ou dimanche.'
      );

    }


    const lock =
      LockService
        .getScriptLock();


    lock.waitLock(
      10000
    );


    try {

      const period =
        getPeriod_(
          selectedPeriodId
        );


      if (
        !period ||
        !period.open
      ) {

        throw new Error(
          'Ce week-end n’est plus disponible.'
        );

      }


      const currentStatus =
        clean_(
          sheet
            .getRange(
              row,
              10
            )
            .getValue()
        );


      if (
        currentStatus !==
        VN.STATUS.CONFIRMED
      ) {

        const currentConfirmed =
          confirmedCounts_()[
            selectedPeriodId
          ] || 0;


        if (
          currentConfirmed >=
          period.capacity
        ) {

          throw new Error(
            'Impossible de confirmer : ce week-end est déjà complet.'
          );

        }

      }


      sheet
        .getRange(
          row,
          10
        )
        .setValue(
          newStatus
        );


      SpreadsheetApp
        .flush();

    }

    finally {

      lock.releaseLock();

    }

  }

  else {

    sheet
      .getRange(
        row,
        10
      )
      .setValue(
        newStatus
      );

  }


  return getAdminDashboard(
    adminKey
  );

}


/* ============================================================
   OUVRIR / FERMER UNE PÉRIODE
   ============================================================ */

function setPeriodOpen(
  adminKey,
  periodId,
  open
) {

  requireAdmin_(
    adminKey
  );


  const period =
    getPeriod_(
      periodId
    );


  if (!period) {

    throw new Error(
      'Week-end introuvable.'
    );

  }


  getSheet_(
    VN.SHEETS.PERIODS
  )
    .getRange(
      period.row,
      6
    )
    .setValue(
      open
        ? 'OUI'
        : 'NON'
    );


  return getAdminDashboard(
    adminKey
  );

}


/* ============================================================
   PÉRIODES ADMIN
   ============================================================ */

function getAdminPeriods_() {

  const requests =
    requestCounts_();


  const confirmed =
    confirmedCounts_();


  return periodRows_()

    .map(
      r => {

        const id =
          clean_(
            r[0]
          );


        const capacity =
          Number(
            r[4]
          ) || 0;


        const confirmedCount =
          confirmed[id] || 0;


        const requestCount =
          requests[id] || 0;


        return {

          id: id,

          label:
            clean_(
              r[1]
            ),

          capacity:
            capacity,

          confirmed:
            confirmedCount,

          remaining:
            Math.max(
              capacity -
              confirmedCount,
              0
            ),

          requestCount:
            requestCount,

          highDemand:
            requestCount >= 4,

          open:
            clean_(
              r[5]
            )
              .toUpperCase()
              === 'OUI',

          order:
            Number(
              r[6]
            ) || 999

        };

      }
    )

    .filter(
      p =>
        p.id
    )

    .sort(
      (
        a,
        b
      ) =>
        a.order -
        b.order
    );

}


/* ============================================================
   DEMANDES ADMIN
   ============================================================ */

function getAdminBookings_() {

  const sheet =
    getSheet_(
      VN.SHEETS.BOOKINGS
    );


  if (
    sheet.getLastRow() < 2
  ) {

    return [];

  }


  return sheet

    .getRange(
      2,
      1,
      sheet.getLastRow() - 1,
      16
    )

    .getValues()

    .filter(
      r =>
        r[0]
    )

    .map(
      r => {

        const ids =
          splitIds_(
            r[7]
          );


        const labels =
          splitLabels_(
            r[8]
          );


        return {

          createdAt:
            formatDate_(
              r[0],
              'dd/MM HH:mm'
            ),

          id:
            clean_(
              r[1]
            ),

          company:
            clean_(
              r[2]
            ),

          contact:
            clean_(
              r[3]
            ),

          email:
            clean_(
              r[4]
            ),

          phone:
            clean_(
              r[5]
            ),

          services:
            clean_(
              r[6]
            ),

          periodsRequested:
            ids.map(
              (
                id,
                index
              ) => ({

                id: id,

                label:
                  labels[
                    index
                  ] ||
                  (
                    getPeriod_(id)
                    || {}
                  ).label ||
                  id

              })
            ),

          status:
            clean_(
              r[9]
            ),

          selectedPeriodId:
            clean_(
              r[10]
            ),

          selectedPeriodLabel:
            clean_(
              r[11]
            ),

          poseDay:
            clean_(
              r[12]
            ),

          removalDate:
            r[13]
              instanceof Date
              ?
              formatDate_(
                r[13],
                'dd/MM/yyyy'
              )
              :
              '',

          comment:
            clean_(
              r[14]
            )

        };

      }
    )

    .reverse();

}


/* ============================================================
   COMPTEUR DEMANDES
   ============================================================ */

/*
 * Toutes les demandes NON REFUSÉES
 * comptent sur tous les week-ends
 * cochés par le client.
 *
 * Cela correspond à :
 * "X demandes déjà reçues".
 */
function requestCounts_() {

  const sheet =
    getSheet_(
      VN.SHEETS.BOOKINGS
    );


  const out = {};


  if (
    sheet.getLastRow() < 2
  ) {

    return out;

  }


  sheet
    .getRange(
      2,
      1,
      sheet.getLastRow() - 1,
      16
    )
    .getValues()
    .forEach(
      r => {

        const status =
          clean_(
            r[9]
          );


        if (
          status ===
          VN.STATUS.REFUSED
        ) {

          return;

        }


        splitIds_(
          r[7]
        )
          .forEach(
            id => {

              out[id] =
                (
                  out[id] || 0
                ) + 1;

            }
          );

      }
    );


  return out;

}


/* ============================================================
   COMPTEUR PLACES RÉELLEMENT BLOQUÉES
   ============================================================ */

function confirmedCounts_() {

  const sheet =
    getSheet_(
      VN.SHEETS.BOOKINGS
    );


  const out = {};


  if (
    sheet.getLastRow() < 2
  ) {

    return out;

  }


  sheet
    .getRange(
      2,
      1,
      sheet.getLastRow() - 1,
      16
    )
    .getValues()
    .forEach(
      r => {

        if (
          clean_(
            r[9]
          )
          !==
          VN.STATUS.CONFIRMED
        ) {

          return;

        }


        const selectedPeriodId =
          clean_(
            r[10]
          );


        if (
          selectedPeriodId
        ) {

          out[
            selectedPeriodId
          ] =
            (
              out[
                selectedPeriodId
              ] || 0
            ) + 1;

        }

      }
    );


  return out;

}


/* ============================================================
   EMAIL ADMIN / JESS
   ============================================================ */

function sendOwnerNotification_(b) {

  const to =
    clean_(
      getParam_(
        'NOTIFICATION_EMAIL'
      )
    );


  if (
    !to ||
    to ===
    'A_RENSEIGNER'
  ) {

    return;

  }


  /*
   * Référence dans l'objet :
   * permet de retrouver précisément
   * cette notification dans Gmail.
   */
  const subject =
    'Nouvelle demande Vitrines de Noël — ' +
    b.company +
    ' — ' +
    b.id;


  MailApp.sendEmail({

    to: to,

    subject:
      subject,

    name:
      clean_(
        getParam_(
          'APP_NAME'
        )
      ) ||
      'Vitrines de Noël',

    body: [

      'Nouvelle demande Vitrines de Noël',

      '',

      'Entreprise : ' +
      b.company,

      'Contact : ' +
      b.contact,

      'Email : ' +
      b.email,

      'Téléphone : ' +
      (
        b.phone ||
        'Non renseigné'
      ),

      '',

      'Prestation(s) : ' +
      b.services
        .join(' + '),

      'Disponibilités : ' +
      b.periodLabels
        .join(', '),

      '',

      'Commentaire : ' +
      (
        b.comment ||
        'Aucun'
      ),

      '',

      'Référence : ' +
      b.id

    ]
      .join('\n')

  });


  /*
   * Gmail range parfois un mail
   * envoyé à soi-même uniquement
   * dans "Messages envoyés".
   *
   * On le replace donc automatiquement
   * dans la boîte de réception
   * si NOTIFICATION_EMAIL correspond
   * à la boîte qui exécute Apps Script.
   */
  surfaceAdminNotification_(

    to,

    b.id

  );

}


/* ============================================================
   REPLACER LA NOTIFICATION ADMIN DANS LA BOÎTE DE RÉCEPTION
   ============================================================ */

function surfaceAdminNotification_(
  to,
  requestId
) {

  const destination =
    clean_(
      to
    )
      .toLowerCase();


  let aliases = [];


  try {

    aliases =
      GmailApp
        .getAliases();

  }

  catch (error) {

    console.error(
      'Impossible de lire les alias Gmail :',
      error
    );

  }


  const ownAddresses = [

    Session
      .getEffectiveUser()
      .getEmail(),

    Session
      .getActiveUser()
      .getEmail(),

    ...aliases

  ]
    .map(
      email =>
        clean_(
          email
        )
          .toLowerCase()
    )
    .filter(Boolean);


  /*
   * Si NOTIFICATION_EMAIL pointe
   * vers une vraie autre boîte,
   * Gmail la recevra normalement :
   * rien à déplacer.
   */
  if (
    !ownAddresses
      .includes(
        destination
      )
  ) {

    return;

  }


  /*
   * Gmail peut prendre quelques
   * millisecondes pour indexer
   * le message envoyé.
   */
  let threads = [];


  for (
    let attempt = 0;
    attempt < 3;
    attempt++
  ) {

    Utilities.sleep(
      800
    );


    threads =
      GmailApp
        .search(
          'in:sent "' +
          requestId +
          '" "Nouvelle demande Vitrines de Noël"',
          0,
          5
        );


    if (
      threads.length
    ) {

      break;

    }

  }


  if (
    !threads.length
  ) {

    console.log(
      'Notification admin envoyée mais thread Gmail non retrouvé immédiatement : ' +
      requestId
    );

    return;

  }


  const thread =
    threads[0];


  const labelName =
    'Vitrines Noël - Nouvelles demandes';


  const label =
    GmailApp
      .getUserLabelByName(
        labelName
      )
      ||
      GmailApp
        .createLabel(
          labelName
        );


  /*
   * On veut que Jess voie
   * réellement une nouvelle alerte.
   */
  thread.moveToInbox();

  thread.markUnread();

  thread.addLabel(
    label
  );

}


/* ============================================================
   EMAIL CLIENT
   ============================================================ */

function sendClientAcknowledgement_(b) {

  const lines = [

    'Bonjour ' +
    b.contact +
    ',',

    '',

    'Votre demande Vitrines de Noël est bien enregistrée.',

    '',

    'Disponibilités indiquées : ' +
    b.periodLabels
      .join(', '),

    '',

    'Cette demande ne vaut pas confirmation définitive du créneau.',

    'La date sera bloquée après validation de votre demande et confirmation de notre part.',

    '',

    'On revient vers vous très vite pour confirmer votre intervention et vous transmettre les éléments nécessaires à la suite.',

    '',

    'Merci de votre confiance.',

    '',

    'Jessica Lacquiez',

    'Vitrines de Noël',

    '',

    'Référence de votre demande : ' +
    b.id

  ];


  MailApp.sendEmail({

    to:
      b.email,

    subject:
      'Votre demande Vitrines de Noël a bien été enregistrée',

    name:
      clean_(
        getParam_(
          'APP_NAME'
        )
      ) ||
      'Vitrines de Noël',

    body:
      lines
        .join('\n')

  });

}


/* ============================================================
   AUTORISATION GMAIL
   ============================================================ */

/*
 * À lancer UNE FOIS manuellement
 * après avoir collé ce nouveau code.
 *
 * Google demandera éventuellement
 * une autorisation supplémentaire
 * pour Gmail.
 */
function autoriserNotificationsAdmin() {

  GmailApp
    .getAliases();


  SpreadsheetApp
    .getUi()
    .alert(

      'Vitrines de Noël',

      'Autorisation Gmail effectuée.\n\nLes notifications administrateur peuvent maintenant être replacées dans la boîte de réception.',

      SpreadsheetApp
        .getUi()
        .ButtonSet
        .OK

    );

}


/* ============================================================
   ACCÈS AU GOOGLE SHEET
   ============================================================ */

function ss_() {

  const id =
    PropertiesService
      .getScriptProperties()
      .getProperty(
        'SPREADSHEET_ID'
      );


  if (!id) {

    throw new Error(
      'Le Sheet n’est pas initialisé. Lance setupVitrinesNoel une première fois.'
    );

  }


  return SpreadsheetApp
    .openById(
      id
    );

}


function getSheet_(name) {

  const sheet =
    ss_()
      .getSheetByName(
        name
      );


  if (!sheet) {

    throw new Error(
      'Onglet "' +
      name +
      '" introuvable.'
    );

  }


  return sheet;

}


/* ============================================================
   HELPERS
   ============================================================ */

function isClientFormClosed_() {

  const cutoff =
    getParam_(
      'DATE_FERMETURE'
    );


  return (
    cutoff
      instanceof Date
      ?
      new Date() >
      cutoff
      :
      false
  );

}


function periodRows_() {

  const sheet =
    getSheet_(
      VN.SHEETS.PERIODS
    );


  if (
    sheet.getLastRow() < 2
  ) {

    return [];

  }


  return sheet
    .getRange(
      2,
      1,
      sheet.getLastRow() - 1,
      8
    )
    .getValues()
    .filter(
      r =>
        clean_(
          r[0]
        )
    );

}


function getPeriod_(id) {

  const rows =
    periodRows_();


  for (
    let i = 0;
    i < rows.length;
    i++
  ) {

    const r =
      rows[i];


    if (
      clean_(
        r[0]
      ) ===
      id
    ) {

      return {

        id:
          clean_(
            r[0]
          ),

        label:
          clean_(
            r[1]
          ),

        capacity:
          Number(
            r[4]
          ) || 0,

        open:
          clean_(
            r[5]
          )
            .toUpperCase()
            === 'OUI',

        row:
          i + 2

      };

    }

  }


  return null;

}


function findRequestRow_(id) {

  const sheet =
    getSheet_(
      VN.SHEETS.BOOKINGS
    );


  if (
    sheet.getLastRow() < 2
  ) {

    return null;

  }


  const ids =
    sheet
      .getRange(
        2,
        2,
        sheet.getLastRow() - 1,
        1
      )
      .getValues();


  for (
    let i = 0;
    i < ids.length;
    i++
  ) {

    if (
      clean_(
        ids[i][0]
      ) ===
      clean_(
        id
      )
    ) {

      return i + 2;

    }

  }


  return null;

}


function splitIds_(value) {

  return clean_(
    value
  )
    ?
    clean_(
      value
    )
      .split('|')
      .map(
        clean_
      )
      .filter(Boolean)
    :
    [];

}


function splitLabels_(value) {

  return clean_(
    value
  )
    ?
    clean_(
      value
    )
      .split('|')
      .map(
        clean_
      )
      .filter(Boolean)
    :
    [];

}


function readParamsFromSheet_(sheet) {

  const out = {};


  if (
    sheet.getLastRow() < 2
  ) {

    return out;

  }


  sheet
    .getRange(
      2,
      1,
      sheet.getLastRow() - 1,
      2
    )
    .getValues()
    .forEach(
      r => {

        if (
          clean_(
            r[0]
          )
        ) {

          out[
            clean_(
              r[0]
            )
          ] =
            r[1];

        }

      }
    );


  return out;

}


function getParam_(key) {

  const sheet =
    getSheet_(
      VN.SHEETS.PARAMS
    );


  if (
    sheet.getLastRow() < 2
  ) {

    return '';

  }


  const rows =
    sheet
      .getRange(
        2,
        1,
        sheet.getLastRow() - 1,
        2
      )
      .getValues();


  for (
    const r of rows
  ) {

    if (
      clean_(
        r[0]
      ) ===
      key
    ) {

      return r[1];

    }

  }


  return '';

}


function clean_(value) {

  return String(
    value == null
      ? ''
      : value
  )
    .trim();

}


function formatDate_(
  value,
  pattern
) {

  if (
    !(
      value
      instanceof Date
    )
  ) {

    return clean_(
      value
    );

  }


  return Utilities
    .formatDate(

      value,

      clean_(
        getParam_(
          'TIMEZONE'
        )
      ) ||
      'Europe/Paris',

      pattern

    );

}


function makeRequestId_() {

  return (

    'DEM-' +

    Utilities
      .formatDate(

        new Date(),

        clean_(
          getParam_(
            'TIMEZONE'
          )
        ) ||
        'Europe/Paris',

        'yyyyMMdd-HHmmss'

      )

    +

    '-' +

    Math.floor(
      Math.random()
      * 900
      + 100
    )

  );

}


/* ============================================================
   ADMIN KEY
   ============================================================ */

function isAdminKeyValid_(
  candidate
) {

  const expected =
    clean_(
      getParam_(
        'ADMIN_KEY'
      )
    );


  return Boolean(

    expected &&

    candidate &&

    clean_(
      candidate
    ) ===
    expected

  );

}


function requireAdmin_(
  candidate
) {

  if (
    !isAdminKeyValid_(
      candidate
    )
  ) {

    throw new Error(
      'Accès administrateur refusé.'
    );

  }

}


function showAdminKey() {

  SpreadsheetApp
    .getUi()
    .alert(

      'Clé admin',

      clean_(
        getParam_(
          'ADMIN_KEY'
        )
      ),

      SpreadsheetApp
        .getUi()
        .ButtonSet
        .OK

    );

}