/******************************************************
 * RaphNLab Kids MINT Workshop Management
 *
 * Phase 1:
 * Google Forms
 *      ↓
 * Google Sheets
 *      ↓
 * Teilnehmer-ID
 *      ↓
 * Teilnehmer-Sheet
 *      ↓
 * Bestätigungs-E-Mail
 ******************************************************/

const CONFIG = {

  // Tabellenblätter
  FORM_SHEET_NAME: 'Form_Responses',
  PARTICIPANTS_SHEET_NAME: 'Teilnehmer',
  WORKSHOPS_SHEET_NAME: 'Workshops',

  // Teilnehmer-ID
  PARTICIPANT_ID_PREFIX: 'MINT',

  // E-Mail
  SEND_CONFIRMATION_EMAIL: true,

  // Organisation
  ORGANIZATION_NAME: 'RaphNLab',

  // Kontakt
  CONTACT_EMAIL: 's.ngoufack@raphnlab.de',
  CONTACT_PHONE: '0176 42732196',
  WEBSITE: 'https://raphnlab.github.io/website/'
};


/**
 * Wird automatisch aufgerufen,
 * wenn eine neue Google-Forms-Antwort eingeht.
 */
function onFormSubmit(e) {

  try {

    if (!e || !e.range) {
      throw new Error(
        'onFormSubmit wurde ohne gültiges Event aufgerufen.'
      );
    }

    const sheet = e.range.getSheet();

    Logger.log(
      'Neue Formularantwort auf Sheet: ' +
      sheet.getName()
    );

    processRegistration(e);

  } catch (error) {

    Logger.log(
      'Fehler in onFormSubmit: ' +
      error.stack
    );

    throw error;
  }
}


/**
 * Verarbeitet eine neue Anmeldung.
 */
function processRegistration(e) {

  const spreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();

  const responseSheet =
    e.range.getSheet();

  const responseRow =
    e.range.getRow();


  /**************************************************
   * 1. Formularantwort auslesen
   **************************************************/

  const headers =
    responseSheet
      .getRange(
        1,
        1,
        1,
        responseSheet.getLastColumn()
      )
      .getValues()[0];

  const values =
    responseSheet
      .getRange(
        responseRow,
        1,
        1,
        responseSheet.getLastColumn()
      )
      .getValues()[0];

  const data =
    createFormData(headers, values);


  Logger.log(
    'Verarbeite Formularzeile: ' +
    responseRow
  );


  /**************************************************
   * 2. Teilnehmer-Sheet laden
   **************************************************/

  const participantSheet =
    getOrCreateParticipantsSheet(
      spreadsheet
    );


  /**************************************************
   * 3. Prüfen, ob Teilnehmer bereits existiert
   **************************************************/

  const existingParticipant =
    findParticipantByFormRow(
      participantSheet,
      responseRow
    );


  /**************************************************
   * FALL A:
   *
   * Teilnehmer existiert bereits
   **************************************************/

  if (existingParticipant) {

    Logger.log(
      'Teilnehmer existiert bereits.'
    );


    const emailStatus =
      participantSheet
        .getRange(
          existingParticipant.row,
          18
        )
        .getValue();


    /*
     * Wenn die E-Mail bereits erfolgreich
     * versendet wurde, ist nichts mehr zu tun.
     */

    if (
      normalizeText(emailStatus) ===
      'gesendet'
    ) {

      Logger.log(
        'Bestätigungs-E-Mail wurde bereits gesendet.'
      );

      return;
    }


    /*
     * Teilnehmer existiert,
     * aber E-Mail wurde noch nicht versendet.
     *
     * Wir holen die bereits vorhandene
     * Teilnehmer-ID und senden die Mail erneut.
     */

    const participantId =
      participantSheet
        .getRange(
          existingParticipant.row,
          1
        )
        .getValue();


    const workshopId =
      participantSheet
        .getRange(
          existingParticipant.row,
          2
        )
        .getValue();


    const workshop =
      getWorkshopById(
        spreadsheet,
        workshopId
      );


    if (!workshop) {

      throw new Error(
        'Workshop ' +
        workshopId +
        ' wurde nicht gefunden.'
      );
    }


    if (
      CONFIG.SEND_CONFIRMATION_EMAIL &&
      data.email
    ) {

      sendConfirmationEmail(
        data,
        workshop,
        participantId
      );


      participantSheet
        .getRange(
          existingParticipant.row,
          18
        )
        .setValue('GESENDET');


      Logger.log(
        'Bestätigungs-E-Mail erneut gesendet.'
      );
    }

    /************************************************
    * Teilnehmerliste aktualisieren
    ************************************************/

    try {

      updateParticipantList();

      Logger.log(
        'Teilnehmerliste erfolgreich aktualisiert.'
      );

    } catch (error) {

      Logger.log(
        'Fehler beim Aktualisieren der Teilnehmerliste: ' +
        error.message
      );
    }
    return;
  }


  /**************************************************
   * FALL B:
   *
   * Neuer Teilnehmer
   **************************************************/


  const workshop =
    getOpenWorkshop(
      spreadsheet
    );


  if (!workshop) {

    throw new Error(
      'Kein Workshop mit Status "Offen" gefunden.'
    );
  }


  /**************************************************
   * 4. Teilnehmer-ID erzeugen
   **************************************************/

  const participantId =
    generateParticipantId(
      participantSheet,
      workshop.id
    );


  /**************************************************
   * 5. Teilnehmer speichern
   **************************************************/

  const participantRow = [

    participantId,

    workshop.id,

    data.firstName,

    data.lastName,

    data.age,

    data.experience,

    data.parentFirstName,

    data.parentLastName,

    data.email,

    data.phone,

    data.registration,

    data.photoConsent,

    data.photoUsage,

    data.notes,

    data.support,

    data.timestamp,

    'ANGEMELDET',

    '',

    responseRow
  ];


  participantSheet
    .appendRow(
      participantRow
    );


  const newParticipantRow =
    participantSheet.getLastRow();


  /**************************************************
   * 6. Bestätigungs-E-Mail
   **************************************************/

  if (
    CONFIG.SEND_CONFIRMATION_EMAIL &&
    data.email
  ) {

    try {

      sendConfirmationEmail(
        data,
        workshop,
        participantId
      );


      participantSheet
        .getRange(
          newParticipantRow,
          18
        )
        .setValue('GESENDET');


      Logger.log(
        'Bestätigungs-E-Mail erfolgreich gesendet.'
      );


    } catch (error) {

      /*
       * Teilnehmer bleibt gespeichert.
       *
       * E-Mail-Status bleibt leer,
       * sodass wir die Mail später erneut
       * versenden können.
       */

      participantSheet
        .getRange(
          newParticipantRow,
          18
        )
        .setValue(
          'FEHLER: ' +
          error.message
        );


      Logger.log(
        'E-Mail konnte nicht gesendet werden: ' +
        error.message
      );


      throw error;
    }
  }


  Logger.log(
    'Teilnehmer erfolgreich registriert: ' +
    participantId
  );
}


function getWorkshopById(
  spreadsheet,
  workshopId
) {

  const sheet =
    spreadsheet.getSheetByName(
      CONFIG.WORKSHOPS_SHEET_NAME
    );


  if (!sheet) {

    throw new Error(
      'Das Sheet "Workshops" wurde nicht gefunden.'
    );
  }


  const lastRow =
    sheet.getLastRow();


  if (lastRow < 2) {
    return null;
  }


  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        7
      )
      .getValues();


  for (const row of values) {

    if (
      String(row[0]).trim() ===
      String(workshopId).trim()
    ) {

      return {

        id: row[0],

        title: row[1],

        date: row[2],

        startTime: row[3],

        endTime: row[4],

        location: row[5],

        status: row[6]
      };
    }
  }


  return null;
}

/**
 * Erstellt ein Objekt aus den Formularantworten.
 *
 * Die Spalten werden anhand ihrer Überschrift
 * erkannt und nicht anhand der Spaltennummer.
 */
function createFormData(headers, values) {

  /*
   * Google Forms / Form Responses 1
   *
   * Die aktuelle Tabelle verwendet leider
   * "Column 1 ... Column 15" als Header.
   *
   * Deshalb verwenden wir hier bewusst die
   * bekannten Spaltenpositionen.
   */

  return {

    // Spalte 1
    timestamp: values[0],

    // Spalte 2
    email: values[1],

    // Spalte 3
    firstName: values[2],

    // Spalte 4
    lastName: values[3],

    // Spalte 5
    age: values[4],

    // Spalte 6
    experience: values[5],

    // Spalte 7
    parentFirstName: values[6],

    // Spalte 8
    parentLastName: values[7],

    // Spalte 9
    phone: values[8],

    // Spalte 10
    photoConsent: values[9],

    // Spalte 11
    photoUsage: values[10],

    // Spalte 12
    notes: values[11],

    // Spalte 13
    support: values[12],

    // Spalte 14
    registration: values[13],

    // Spalte 15
    confirmation: values[14]

  };
}


function repairExistingParticipants() {

  const spreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();

  const participantSheet =
    spreadsheet.getSheetByName(
      'Teilnehmer'
    );

  const responseSheet =
    spreadsheet.getSheetByName(
      'Form Responses 1'
    );


  if (!participantSheet) {

    throw new Error(
      'Sheet "Teilnehmer" wurde nicht gefunden.'
    );
  }


  if (!responseSheet) {

    throw new Error(
      'Sheet "Form Responses 1" wurde nicht gefunden.'
    );
  }


  const lastParticipantRow =
    participantSheet.getLastRow();


  if (lastParticipantRow < 2) {

    Logger.log(
      'Keine Teilnehmer vorhanden.'
    );

    return;
  }


  const participantData =
    participantSheet
      .getRange(
        2,
        1,
        lastParticipantRow - 1,
        19
      )
      .getValues();


  let repaired = 0;
  let emailsSent = 0;


  for (
    let i = 0;
    i < participantData.length;
    i++
  ) {

    const participantRow =
      i + 2;


    const participantId =
      participantData[i][0];


    const formRow =
      Number(
        participantData[i][18]
      );


    if (
      !participantId ||
      !formRow ||
      formRow < 2
    ) {

      continue;
    }


    /**********************************************
     * Formularantwort lesen
     **********************************************/

    const values =
      responseSheet
        .getRange(
          formRow,
          1,
          1,
          15
        )
        .getValues()[0];


    const data =
      createFormData(
        [],
        values
      );


    /**********************************************
     * Workshop
     **********************************************/

    const workshop =
      getWorkshopById(
        spreadsheet,
        participantData[i][1]
      );


    if (!workshop) {

      Logger.log(
        'Workshop nicht gefunden für ' +
        participantId
      );

      continue;
    }


    /**********************************************
     * Teilnehmerdaten aktualisieren
     *
     * Spalten 3-16
     **********************************************/

    participantSheet
      .getRange(
        participantRow,
        3,
        1,
        14
      )
      .setValues([[
        data.firstName,
        data.lastName,
        data.age,
        data.experience,
        data.parentFirstName,
        data.parentLastName,
        data.email,
        data.phone,
        data.registration,
        data.photoConsent,
        data.photoUsage,
        data.notes,
        data.support,
        data.timestamp
      ]]);


    repaired++;


    /**********************************************
     * Bestätigungs-E-Mail
     **********************************************/

    const emailStatus =
      String(
        participantData[i][17] || ''
      ).trim();


    if (
      CONFIG.SEND_CONFIRMATION_EMAIL &&
      data.email &&
      emailStatus !== 'GESENDET'
    ) {

      try {

        sendConfirmationEmail(
          data,
          workshop,
          participantId
        );


        participantSheet
          .getRange(
            participantRow,
            18
          )
          .setValue(
            'GESENDET'
          );


        emailsSent++;


        Logger.log(
          'E-Mail gesendet: ' +
          participantId +
          ' → ' +
          data.email
        );


      } catch (error) {

        participantSheet
          .getRange(
            participantRow,
            18
          )
          .setValue(
            'FEHLER: ' +
            error.message
          );


        Logger.log(
          'E-Mail Fehler bei ' +
          participantId +
          ': ' +
          error.message
        );
      }
    }
  }


  Logger.log(
    'Reparatur abgeschlossen.'
  );


  Logger.log(
    'Teilnehmer repariert: ' +
    repaired
  );


  Logger.log(
    'E-Mails gesendet: ' +
    emailsSent
  );


  SpreadsheetApp
    .getUi()
    .alert(

      'Reparatur abgeschlossen.\n\n' +

      'Teilnehmer aktualisiert: ' +
      repaired +
      '\n' +

      'Bestätigungs-E-Mails gesendet: ' +
      emailsSent

    );
}


/**
 * Sucht einen Wert anhand eines Teils
 * der Spaltenüberschrift.
 */
function getValueByHeader(
  headers,
  values,
  searchText
) {

  const normalizedSearch =
    normalizeText(searchText);

  const index =
    headers.findIndex(header => {

      return normalizeText(header)
        .includes(normalizedSearch);

    });

  if (index === -1) {

    Logger.log(
      'Spalte nicht gefunden: ' +
      searchText
    );

    return '';
  }

  return values[index];
}


/**
 * Normalisiert Text für robuste Vergleiche.
 */
function normalizeText(text) {

  return String(text || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ');
}


/**
 * Erstellt das Teilnehmer-Sheet,
 * falls es noch nicht existiert.
 */
function getOrCreateParticipantsSheet(
  spreadsheet
) {

  let sheet =
    spreadsheet.getSheetByName(
      CONFIG.PARTICIPANTS_SHEET_NAME
    );

  if (!sheet) {

    sheet =
      spreadsheet.insertSheet(
        CONFIG.PARTICIPANTS_SHEET_NAME
      );

    const headers = [

      'Teilnehmer-ID',
      'Workshop-ID',
      'Vorname',
      'Nachname',
      'Alter',
      'Erfahrung',
      'Eltern-Vorname',
      'Eltern-Nachname',
      'E-Mail',
      'Telefon',
      'Anmeldung',
      'Foto-Einwilligung',
      'Foto-Verwendung',
      'Besondere Hinweise',
      'Unterstützung',
      'Timestamp',
      'Status',
      'Bestätigungs-E-Mail',
      'Form-Zeile'
    ];

    sheet
      .getRange(1, 1, 1, headers.length)
      .setValues([headers]);
  }

  return sheet;
}


/**
 * Sucht einen Teilnehmer anhand
 * der ursprünglichen Formularzeile.
 */
function findParticipantByFormRow(
  sheet,
  formRow
) {

  const lastRow =
    sheet.getLastRow();

  if (lastRow < 2) {
    return null;
  }

  const values =
    sheet
      .getRange(
        2,
        19,
        lastRow - 1,
        1
      )
      .getValues();

  for (let i = 0; i < values.length; i++) {

    if (
      Number(values[i][0]) ===
      Number(formRow)
    ) {

      return {
        row: i + 2
      };
    }
  }

  return null;
}


/**
 * Liefert den ersten Workshop mit Status "Offen".
 */
function getOpenWorkshop(
  spreadsheet
) {

  const sheet =
    spreadsheet.getSheetByName(
      CONFIG.WORKSHOPS_SHEET_NAME
    );

  if (!sheet) {

    throw new Error(
      'Das Sheet "Workshops" wurde nicht gefunden.'
    );
  }

  const lastRow =
    sheet.getLastRow();

  if (lastRow < 2) {
    return null;
  }

  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        7
      )
      .getValues();

  for (const row of values) {

    const status =
      normalizeText(row[6]);

    if (status === 'offen') {

      return {

        id: row[0],

        title: row[1],

        date: row[2],

        startTime: row[3],

        endTime: row[4],

        location: row[5],

        status: row[6]
      };
    }
  }

  return null;
}


/**
 * Erzeugt eine neue Teilnehmer-ID.
 *
 * Beispiel:
 *
 * MINT-2026-001
 * MINT-2026-002
 * MINT-2026-003
 */
function generateParticipantId(
  participantSheet,
  workshopId
) {

  const year =
    new Date().getFullYear();

  const prefix =
    CONFIG.PARTICIPANT_ID_PREFIX +
    '-' +
    year;

  const lastRow =
    participantSheet.getLastRow();

  let nextNumber = 1;

  if (lastRow >= 2) {

    const ids =
      participantSheet
        .getRange(
          2,
          1,
          lastRow - 1,
          1
        )
        .getValues();

    let highest = 0;

    ids.forEach(row => {

      const id =
        String(row[0] || '');

      const match =
        id.match(
          new RegExp(
            '^' +
            prefix +
            '-(\\d+)$'
          )
        );

      if (match) {

        const number =
          parseInt(
            match[1],
            10
          );

        if (number > highest) {
          highest = number;
        }
      }
    });

    nextNumber =
      highest + 1;
  }

  return (
    prefix +
    '-' +
    String(nextNumber)
      .padStart(3, '0')
  );
}


/**
 * Sendet die Anmeldebestätigung.
 */
function sendConfirmationEmail(
  data,
  workshop,
  participantId
  )
{

  /**************************************************
   * 1. Brevo API-Key laden
   **************************************************/

  const apiKey =
    PropertiesService
      .getScriptProperties()
      .getProperty('BREVO_API_KEY');


  if (!apiKey) {

    throw new Error(
      'BREVO_API_KEY wurde nicht gefunden.'
    );
  }


  /**************************************************
   * 2. Daten vorbereiten
   **************************************************/

  const workshopDate =
    formatDate(
      workshop.date
    );


  const parentName =
    (
      data.parentFirstName +
      ' ' +
      data.parentLastName
    ).trim();


  const childName =
    (
      data.firstName +
      ' ' +
      data.lastName
    ).trim();


  /**************************************************
   * 3. HTML-E-Mail erzeugen
   **************************************************/

  const htmlBody =
    createConfirmationHtml(
      data,
      workshop,
      participantId
    );


  /**************************************************
   * 4. Plain-Text-Version erzeugen
   **************************************************/

  const textBody =
    createConfirmationText(
      data,
      workshop,
      participantId
    );


  /**************************************************
   * 5. Brevo Payload
   **************************************************/

  const payload = {

    sender: {
      name: 'RaphNLab',

      email: 's.ngoufack@raphnlab.de'
    },
    to: [
      {
        email: data.email,

        name: parentName
      }
    ],


    replyTo: {
      email: CONFIG.CONTACT_EMAIL,
      name: 'RaphNLab'
    },

    subject:
      'Anmeldung bestätigt – ' +
      workshop.title +
      ' am ' +
      workshopDate,
    htmlContent:
      htmlBody,
    textContent:
      textBody,
    tags: [
      'raphnlab',
      'kids-mint-workshop',
      'registration'
    ]
  };


  /**************************************************
   * 6. E-Mail über Brevo senden
   **************************************************/

  const response =
    UrlFetchApp.fetch(
      'https://api.brevo.com/v3/smtp/email',
      {
        method:
          'post',
        contentType:
          'application/json',
        headers: {
          'accept':
            'application/json',
          'api-key':
            apiKey
        },
        payload:
          JSON.stringify(payload),
        muteHttpExceptions:
          true
      }
    );


  /**************************************************
   * 7. Antwort auswerten
   **************************************************/

  const responseCode =
    response.getResponseCode();

  const responseBody =
    response.getContentText();

  Logger.log(
    'Brevo HTTP Status: ' +
    responseCode
  );

  Logger.log(
    'Brevo Response: ' +
    responseBody
  );


  /**************************************************
   * 8. Fehler behandeln
   **************************************************/

  if (responseCode < 200 || responseCode >= 300) {
    throw new Error(
      'Brevo E-Mail-Versand fehlgeschlagen. ' +
      'HTTP ' + responseCode + ': ' + responseBody
    );
  }


  /**************************************************
   * 9. Erfolgreichen Versand protokollieren
   **************************************************/

  let messageId = '';
  try {
    const result =
      JSON.parse(
        responseBody
      );

    messageId =
      result.messageId ||
      '';

  } catch (error) {

    Logger.log(
      'Brevo Response konnte nicht als JSON gelesen werden.'
    );
  }


  Logger.log(
    'Bestätigungs-E-Mail erfolgreich über Brevo gesendet.'
  );


  if (messageId) {

    Logger.log(
      'Brevo Message-ID: ' +
      messageId
    );
  }


  return messageId;
}



function createConfirmationHtml(
  data,
  workshop,
  participantId
) {

  const parentName =
    escapeHtml(
      (
        data.parentFirstName +
        ' ' +
        data.parentLastName
      ).trim()
    );


  const childName =
    escapeHtml(
      (
        data.firstName +
        ' ' +
        data.lastName
      ).trim()
    );


  const safeParticipantId =
    escapeHtml(
      participantId
    );


  const safeWorkshopTitle =
    escapeHtml(
      workshop.title
    );


  const safeWorkshopDate =
    escapeHtml(
      formatDate(
        workshop.date
      )
    );


  const safeStartTime =
    escapeHtml(
      formatDate(workshop.startTime)
    );


  const safeEndTime =
    escapeHtml(
      formatDate(workshop.endTime)
    );


  const safeLocation =
    escapeHtml(
      workshop.location
    );


  return `

<!DOCTYPE html>
  <html>
    <head>
    <meta charset="UTF-8">
    </head>

    <body style="
      margin:0;
      padding:0;
      background:#f4f6f8;
      font-family:Arial,Helvetica,sans-serif;
    ">

      <table
        width="100%"
        cellpadding="0"
        cellspacing="0"
        style="
          background:#f4f6f8;
          padding:30px 10px;
        "
      >
        <tr>
          <td align="center">
            <table
              width="600"
              cellpadding="0"
              cellspacing="0"
              style="
                max-width:600px;
                background:#ffffff;
                border-radius:12px;
                overflow:hidden;
                box-shadow:
                  0 2px 8px
                  rgba(0,0,0,0.08);
              ">

              <!-- HEADER -->
                <tr>
                  <td
                    style="
                      padding:28px 30px;
                      background:#1f2937;
                      color:#ffffff;
                      text-align:center;
                    "
                  >
                    <div style="
                      font-size:26px;
                      font-weight:bold;
                    "> RaphNLab </div>


                    <div style="
                      margin-top:8px;
                      font-size:15px;
                      opacity:0.9;
                    "> Kids MINT Workshop </div>
                  </td>
                </tr>


              <!-- CONTENT -->

              <tr>
                <td style="
                  padding:35px;
                ">
                  <h1 style="
                    margin:0 0 20px 0;
                    font-size:25px;
                    color:#1f2937;
                  "> Anmeldung bestätigt 🎉 </h1>

                  <p style="
                    font-size:16px;
                    line-height:1.6;
                    color:#374151;
                  "> Guten Tag ${parentName}, </p>

                  <p style="
                    font-size:16px;
                    line-height:1.6;
                    color:#374151;
                  "> vielen Dank für die Anmeldung von <strong> ${childName} </strong> zum Kids MINT Workshop. </p>

                  <!-- PARTICIPANT -->

                  <table
                    width="100%"
                    cellpadding="0"
                    cellspacing="0"
                    style="
                      margin:25px 0;
                      background:#f8fafc;
                      border-radius:8px;
                    "
                  >
                    <tr>
                      <td style="
                        padding:20px;
                      ">

                        <div style="
                          font-size:13px;
                          color:#6b7280;
                          text-transform:uppercase;
                          letter-spacing:0.5px;
                        "> Teilnehmer </div>

                        <div style="
                          margin-top:7px;
                          font-size:18px;
                          font-weight:bold;
                          color:#111827;
                        "> ${childName} </div>

                        <div style="
                          margin-top:8px;
                          font-size:14px;
                          color:#4b5563;
                        "> Teilnehmer-ID: <strong> ${safeParticipantId} </strong> </div>
                      </td>
                    </tr>

                  </table>


                  <!-- WORKSHOP -->

                  <h2 style="
                    font-size:18px;
                    color:#1f2937;
                    margin-top:30px;
                  "> Workshop </h2>

                  <table
                    width="100%"
                    cellpadding="0"
                    cellspacing="0"
                    style="
                      font-size:15px;
                      color:#374151;
                      line-height:1.8;
                    "
                  >

                    <tr>
                      <td width="90"> <strong>Titel:</strong> </td>
                      <td> ${safeWorkshopTitle}</td>
                    </tr>


                    <tr>
                      <td> <strong>Datum:</strong> </td>
                      <td> ${safeWorkshopDate} </td>
                    </tr>


                    <tr>
                      <td> <strong>Zeit:</strong> </td> 
                      <td> ${safeStartTime} - ${safeEndTime} Uhr </td>
                    </tr>

                    <tr>
                      <td> <strong>Ort:</strong> </td>
                      <td> ${safeLocation} </td>
                    </tr>

                  </table>

                  <!-- ABLAUF -->

                  <h2 style="
                    font-size:18px;
                    color:#1f2937;
                    margin-top:30px;
                  "> Ablauf </h2>

                  <table
                    width="100%"
                    cellpadding="0"
                    cellspacing="0"
                    style="
                      font-size:15px;
                      color:#374151;
                      line-height:1.8;
                    "
                  >
                    <ul>
                      <li>12:00 - 12:30 Uhr: Ankunft und gegenseitiges Kennenlernen</li>
                      <li>12:30 - 13:15 Uhr: 3D sehen ohne Brille Erleben</li>
                      <li>13:15 - 14:00 Uhr: Austausch mit den Kindern über den eigenen beruflichen Werdegang und persönliche Erfahrungen</li>
                      <li>14:00 - 15:45 Uhr: Praktische Übungen mit elektronischen Bauteilen und Konstruktionen</li>
                      <li>15:45 - 16:00 Uhr: Teilnahmeschein Vergabe und Verabschiedung</li>
                    </ul>
                  </table>

                  <p style="
                    margin-top:25px;
                    padding:15px;
                    background:#f0fdf4;
                    border-radius:8px;
                    color:#166534;
                    font-size:15px;
                    line-height:1.5;
                  "> <strong> Die Teilnahme am Workshop ist kostenlos.</strong> </p>


                  <p style="
                    font-size:16px;
                    line-height:1.6;
                    color:#374151;
                  "> Sollte Ihr Kind nach der Anmeldung verhindert sein, informieren Sie uns bitte so früh wie möglich. </p>


                  <p style="
                    font-size:16px;
                    line-height:1.6;
                    color:#374151;
                  "> Wir freuen uns auf <strong> ${childName} </strong> und einen spannenden Workshop! 🚀 </p>
                </td>
              </tr>


              <!-- FOOTER -->
              <tr>

                <td style="
                  padding:25px 35px;
                  background:#f8fafc;
                  border-top:1px solid #e5e7eb;
                ">
                  <div style="
                    font-weight:bold;
                    color:#1f2937;
                  "> RaphNLab </div>

                  <div style="
                    margin-top:8px;
                    font-size:13px;
                    line-height:1.7;
                    color:#6b7280;
                  "> 
                      E-Mail: ${escapeHtml(CONFIG.CONTACT_EMAIL)} 
                    <br>
                      Telefon: ${escapeHtml(CONFIG.CONTACT_PHONE)}
                    <br>
                      Website: <a href=${escapeHtml(CONFIG.WEBSITE)}>raphnlab.de</a>
                  </div>

                  <div style="
                    margin-top:15px;
                    font-size:12px;
                    color:#9ca3af;
                  "> Diese E-Mail wurde automatisch nach der Anmeldung zum Kids MINT Workshop versendet. </div>

                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </body>
  </html>
`;
}

function createConfirmationText(
  data,
  workshop,
  participantId
) {

  const workshopDate =
    formatDate(
      workshop.date
    );


  const childName =
    (
      data.firstName +
      ' ' +
      data.lastName
    ).trim();


  const parentName =
    (
      data.parentFirstName +
      ' ' +
      data.parentLastName
    ).trim();


  return (

    'Guten Tag ' + parentName + ',\n\n' +

    'vielen Dank für die Anmeldung von ' +
    childName + ' zum Kids MINT Workshop.\n\n' +

    'TEILNEHMER\n' +
    '------------------------------\n' +
    'Name: ' + childName +  '\n' +
    'Teilnehmer-ID: ' + participantId +  '\n\n' +
    'WORKSHOP\n' + 
    '------------------------------\n' +
    'Titel: ' + workshop.title + '\n' +
    'Datum: ' + workshopDate + '\n' +
    'Zeit: ' + formatDate(workshop.startTime) + ' - ' +  formatDate(workshop.endTime) + '\n' +
    'Ort: ' +  workshop.location + '\n' +
    'ABLAUF: ' +
    '------------------------------\n' +
    '-> 12:00 - 12:30 Uhr: Ankunft und gegenseitiges Kennenlernen' +
    '-> 12:30 - 13:15 Uhr: 3D sehen ohne Brille Erleben' +
    '-> 13:15 - 14:00 Uhr: Austausch mit den Kindern über den eigenen beruflichen Werdegang und persönliche Erfahrungen' +
    '-> 14:00 - 15:45 Uhr: Praktische Übungen mit elektronischen Bauteilen und Konstruktionen ' +
    '-> 15:45 - 16:00 Uhr: Teilnahmeschein Vergabe und Verabschiedung' +
    '\n\n' +

    'Die Teilnahme am Workshop ist kostenlos.\n\n' +

    'Sollte Ihr Kind nach der Anmeldung verhindert sein, ' + 'informieren Sie uns bitte so früh wie möglich.\n\n' +

    'Wir freuen uns auf ' + childName + ' und einen spannenden Workshop!\n\n' +

    'Viele Grüße\n' +
    'RaphNLab\n\n' +

    'E-Mail: ' +
    CONFIG.CONTACT_EMAIL +
    '\n' +

    'Telefon: ' +
    CONFIG.CONTACT_PHONE +

    'Website: ' +
    CONFIG.WEBSITE

  );
}


function escapeHtml(value) {

  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

/**
 * Formatiert ein Datum für die E-Mail.
 */
function formatDate(date) {

  if (!date) {
    return '';
  }

  if (
    Object.prototype.toString
      .call(date) ===
    '[object Date]'
  ) {

    return Utilities
      .formatDate(
        date,
        Session.getScriptTimeZone(),
        'dd.MM.yyyy'
      );
  }

  return String(date);
}


/**
 * Testfunktion.
 *
 * Damit können wir überprüfen,
 * ob das Script grundsätzlich läuft.
 */
function testConfiguration() {

  const spreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();

  const workshop =
    getOpenWorkshop(
      spreadsheet
    );

  if (!workshop) {

    Logger.log(
      'FEHLER: Kein offener Workshop gefunden.'
    );

    return;
  }

  Logger.log(
    'Workshop gefunden: ' +
    workshop.id +
    ' - ' +
    workshop.title
  );

  Logger.log(
    'Datum: ' +
    formatDate(workshop.date)
  );

  Logger.log(
    'Ort: ' +
    workshop.location
  );
}


/**
 * Erstellt den automatischen Trigger.
 *
 * Diese Funktion muss nur EINMAL
 * manuell ausgeführt werden.
 */
function setupTrigger() {

  const spreadsheet =
    SpreadsheetApp.getActive();

  const triggers =
    ScriptApp.getProjectTriggers();

  // Bereits vorhandenen Trigger entfernen,
  // damit kein doppelter Trigger entsteht.
  triggers.forEach(trigger => {

    if (
      trigger.getHandlerFunction() ===
      'onFormSubmit'
    ) {

      ScriptApp.deleteTrigger(
        trigger
      );
    }
  });


  ScriptApp.newTrigger(
    'onFormSubmit'
  )
    .forSpreadsheet(
      spreadsheet
    )
    .onFormSubmit()
    .create();


  Logger.log(
    'Formular-Trigger erfolgreich erstellt.'
  );
}

function testConfirmationEmail() {

  const testData = {

    email:
      'ngoufackss@yahoo.fr',

    parentFirstName:
      'Test',

    parentLastName:
      'Familie',

    firstName:
      'Max',

    lastName:
      'Mustermann'
  };


  const testWorkshop = {

    id:
      'MINT-2026-01',

    title:
      'Kids MINT Workshop',

    date:
      new Date(2026, 9, 10),

    startTime:
      '12:00',

    endTime:
      '16:00',

    location:
      'Robert-Bosch-Straße 33, 73430 Aalen'
  };


  const messageId =
    sendConfirmationEmail(
      testData,
      testWorkshop,
      'MINT-2026-TEST'
    );


  Logger.log(
    'Test erfolgreich. Brevo Message-ID: ' +
    messageId
  );
}


function onOpen() {

  const ui = SpreadsheetApp.getUi();

  ui.createMenu('RaphNLab')
    .addItem(
      'Teilnehmerliste aktualisieren',
      'updateParticipantList'
    )
    .addItem(
      'Teilnehmerliste anzeigen',
      'showParticipantList'
    )
    .addSeparator()
    .addItem(
      'Systemstatus',
      'showSystemStatus'
    )
    .addToUi();
}

function updateParticipantList() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sourceSheet =
    ss.getSheetByName('Teilnehmer');

  const workshopSheet =
    ss.getSheetByName('Workshops');


  if (!sourceSheet) {
    throw new Error(
      'Das Sheet "Teilnehmer" wurde nicht gefunden.'
    );
  }


  if (!workshopSheet) {
    throw new Error(
      'Das Sheet "Workshops" wurde nicht gefunden.'
    );
  }


  /************************************************
   * Teilnehmerdaten lesen
   ************************************************/

  const participantData =
    sourceSheet
      .getDataRange()
      .getValues();


  if (participantData.length < 2) {

    SpreadsheetApp.getUi().alert(
      'Es sind noch keine Teilnehmer vorhanden.'
    );

    return;
  }


  const participantHeaders =
    participantData[0];


  const participantColumnIndex =
    getColumnIndices(
      participantHeaders
    );


  /************************************************
   * Workshopdaten lesen
   ************************************************/

  const workshopData =
    workshopSheet
      .getDataRange()
      .getValues();


  if (workshopData.length < 2) {

    throw new Error(
      'Im Sheet "Workshops" sind keine Workshops vorhanden.'
    );
  }


  const workshopHeaders =
    workshopData[0];


  const workshopColumnIndex =
    getColumnIndices(
      workshopHeaders
    );


  /************************************************
   * Workshops nach Workshop-ID indizieren
   ************************************************/

  const workshopsById = {};


  for (
    let i = 1;
    i < workshopData.length;
    i++
  ) {

    const row =
      workshopData[i];


    const workshopId =
      String(
        getValue(
          row,
          workshopColumnIndex,
          'Workshop-ID'
        ) || ''
      ).trim();


    if (!workshopId) {
      continue;
    }


    workshopsById[workshopId] = {

      title:
        getValue(
          row,
          workshopColumnIndex,
          'Titel'
        ),

      date:
        getValue(
          row,
          workshopColumnIndex,
          'Datum'
        ),

      startTime:
        getValue(
          row,
          workshopColumnIndex,
          'Beginn'
        ),

      endTime:
        getValue(
          row,
          workshopColumnIndex,
          'Ende'
        )

    };
  }


  /************************************************
   * Bereits gesetzte Verwaltungswerte sichern
   ************************************************/

  const existingManagement =
    getExistingManagementValues(
      ss,
      'Teilnehmerliste'
    );


  /************************************************
   * Ausgabe vorbereiten
   ************************************************/

  const output = [];


  output.push([

    'Teilnehmer-ID',
    'Kind',
    'Workshop',
    'Datum',
    'Zeit',
    'Elternteil',
    'E-Mail',
    'Status',
    'Teilnahme',
    'Zertifikat'

  ]);


  /************************************************
   * Teilnehmer verarbeiten
   ************************************************/

  for (
    let i = 1;
    i < participantData.length;
    i++
  ) {

    const row =
      participantData[i];


    // Leere Zeilen ignorieren
    if (
      row.every(
        value =>
          value === '' ||
          value === null
      )
    ) {
      continue;
    }


    /**********************************************
     * Teilnehmer-ID
     **********************************************/

    const participantId =
      String(
        getValue(
          row,
          participantColumnIndex,
          'Teilnehmer-ID'
        ) || ''
      ).trim();


    if (!participantId) {
      continue;
    }


    /**********************************************
     * Workshop-ID
     **********************************************/

    const workshopId =
      String(
        getValue(
          row,
          participantColumnIndex,
          'Workshop-ID'
        ) || ''
      ).trim();


    /**********************************************
     * Teilnehmer
     **********************************************/

    const firstName =
      getValue(
        row,
        participantColumnIndex,
        'Vorname'
      ) || '';


    const lastName =
      getValue(
        row,
        participantColumnIndex,
        'Nachname'
      ) || '';


    const parentFirstName =
      getValue(
        row,
        participantColumnIndex,
        'Elternteil Vorname'
      ) || '';


    const parentLastName =
      getValue(
        row,
        participantColumnIndex,
        'Elternteil Nachname'
      ) || '';


    const email =
      getValue(
        row,
        participantColumnIndex,
        'E-Mail'
      ) || '';


    const status =
      getValue(
        row,
        participantColumnIndex,
        'Status'
      ) || '';


    /**********************************************
     * Namen zusammensetzen
     **********************************************/

    const childName =
      (
        firstName +
        ' ' +
        lastName
      ).trim();


    const parentName =
      (
        parentFirstName +
        ' ' +
        parentLastName
      ).trim();


    /**********************************************
     * Workshop anhand Workshop-ID suchen
     **********************************************/

    const workshop =
      workshopsById[workshopId];


    let workshopTitle = '';
    let workshopDate = '';
    let startTime = '';
    let endTime = '';


    if (workshop) {

      workshopTitle =
        workshop.title || '';

      workshopDate =
        workshop.date || '';

      startTime =
        workshop.startTime || '';

      endTime =
        workshop.endTime || '';

    }
    else {

      Logger.log(
        'Workshop-ID nicht gefunden: ' +
        workshopId +
        ' für Teilnehmer ' +
        participantId
      );

    }


    /**********************************************
     * Zeit formatieren
     **********************************************/

    let time = '';


    if (
      startTime ||
      endTime
    ) {

      time =
        `${startTime} – ${endTime}`;

    }


    /**********************************************
     * Verwaltungswerte erhalten
     **********************************************/

    const management =
      existingManagement[
        participantId
      ] || {};


    /**********************************************
     * Teilnehmerliste
     **********************************************/

    output.push([

      participantId,

      childName,

      workshopTitle,

      formatDate(
        workshopDate
      ),

      time,

      parentName,

      email,

      status,

      management.participation ||
        'OFFEN',

      management.certificate ||
        'OFFEN'

    ]);
  }


  /************************************************
   * Ziel-Sheet
   ************************************************/

  let targetSheet =
    ss.getSheetByName(
      'Teilnehmerliste'
    );


  if (!targetSheet) {

    targetSheet =
      ss.insertSheet(
        'Teilnehmerliste'
      );
  }


  /************************************************
   * Bestehenden Inhalt löschen
   ************************************************/

  targetSheet.clear();


  /************************************************
   * Filterbereich
   ************************************************/

  targetSheet
    .getRange('A1')
    .setValue('Workshop:');


  targetSheet
    .getRange('A2')
    .setValue('Status:');


  targetSheet
    .getRange('A1:A2')
    .setFontWeight('bold');


  targetSheet
    .getRange('B1')
    .setValue('ALLE');


  targetSheet
    .getRange('B2')
    .setValue('ALLE');


  /************************************************
   * Daten ab Zeile 4
   ************************************************/

  targetSheet
    .getRange(
      4,
      1,
      output.length,
      output[0].length
    )
    .setValues(
      output
    );


  /************************************************
   * Formatierung
   ************************************************/

  formatParticipantList(
    targetSheet,
    output.length + 3
  );


  /************************************************
   * Verwaltungs-Dropdowns
   ************************************************/

  createManagementDropdowns(
    targetSheet,
    output.length + 3
  );


  /************************************************
   * Filter-Dropdowns
   ************************************************/

  createParticipantFilters(
    targetSheet,
    output
  );


  /************************************************
   * Log
   ************************************************/

  Logger.log(
    'Teilnehmerliste erfolgreich aktualisiert. ' +
    'Teilnehmer: ' +
    (output.length - 1)
  );


  SpreadsheetApp
    .getUi()
    .alert(
      'Teilnehmerliste erfolgreich aktualisiert.\n\n' +
      'Teilnehmer: ' +
      (output.length - 1)
    );
}

function getExistingManagementValues(
  ss,
  sheetName
) {

  const result = {};
  const sheet =
    ss.getSheetByName(sheetName);

  if (!sheet) {
    return result;
  }

  const lastRow =
    sheet.getLastRow();


  if (lastRow < 5) {
    return result;
  }

  const data =
    sheet
      .getRange(
        5,
        1,
        lastRow - 4,
        10
      )
      .getValues();


  data.forEach(row => {
    const participantId =
      String(row[0]).trim();

    if (!participantId) {
      return;
    }

    result[participantId] = {
      participation:
        row[8] || 'OFFEN',
      certificate:
        row[9] || 'OFFEN'

    };
  });

  return result;
}

function createParticipantFilters(
  sheet,
  output
) {

  const workshops = new Set();
  const statuses = new Set();


  for (
    let i = 1;
    i < output.length;
    i++
  ) {

    const workshop =
      output[i][2];

    const status =
      output[i][7];


    if (workshop) {
      workshops.add(
        String(workshop)
      );
    }


    if (status) {
      statuses.add(
        String(status)
      );
    }
  }


  const workshopValues = [
    'ALLE',
    ...Array.from(workshops).sort()
  ];


  const statusValues = [
    'ALLE',
    ...Array.from(statuses).sort()
  ];


  /************************************************
   * Workshop Dropdown
   ************************************************/

  const workshopRule =
    SpreadsheetApp
      .newDataValidation()
      .requireValueInList(
        workshopValues,
        true
      )
      .setAllowInvalid(false)
      .build();


  sheet
    .getRange('B1')
    .setDataValidation(
      workshopRule
    );


  /************************************************
   * Status Dropdown
   ************************************************/

  const statusRule =
    SpreadsheetApp
      .newDataValidation()
      .requireValueInList(
        statusValues,
        true
      )
      .setAllowInvalid(false)
      .build();


  sheet
    .getRange('B2')
    .setDataValidation(
      statusRule
    );


  /************************************************
   * Filteränderung überwachen
   ************************************************/

  // Die eigentliche Filterlogik wird
  // über onEdit(e) umgesetzt.
}

function onEdit(e) {

  if (!e || !e.range) {
    return;
  }

  const sheet = e.range.getSheet();

  if (sheet.getName() !== 'Teilnehmerliste') {
    return;
  }

  const row = e.range.getRow();
  const column =  e.range.getColumn();

  /************************************************
   * Nur B1 oder B2
   ************************************************/

  if (column !== 2 || (row !== 1 && row !== 2)) {
    return;
  }

  applyParticipantFilter(
    sheet
  );
}

function applyParticipantFilter(
  sheet
) {

  const workshopFilter =
    String(
      sheet
        .getRange('B1')
        .getValue()
    ).trim();


  const statusFilter =
    String(
      sheet
        .getRange('B2')
        .getValue()
    ).trim();


  const lastRow = sheet.getLastRow();


  if (lastRow < 5) {
    return;
  }


  const dataRange =
    sheet.getRange(
      5,
      1,
      lastRow - 4,
      10
    );


  const data = dataRange.getValues();


  for (let i = 0; i < data.length; i++) {

    const workshop = String(data[i][2]).trim();
    const status = String(data[i][7]).trim();


    const workshopMatches =
      workshopFilter === 'ALLE' ||
      workshop === workshopFilter;


    const statusMatches = statusFilter === 'ALLE' || status === statusFilter;

    const visible = workshopMatches && statusMatches;

    sheet.showRows(i + 5);

    if (!visible) {
      sheet.hideRows(i + 5);
    }
  }
}

function getColumnIndices(headers) {

  const indices = {};

  headers.forEach(
    (header, index) => {
      indices[
        String(header).trim()
      ] = index;

    }
  );

  return indices;
}

function getValue(
  row,
  columnIndex,
  columnName
) {

  const index =
    columnIndex[columnName];

  if (
    index === undefined
  ) {

    return '';
  }
  return row[index];
}

function formatParticipantList(
  sheet,
  numberOfRows
) {

  const numberOfColumns = 10;

  const headerRow = 4;


  /************************************************
   * Header
   ************************************************/

  const headerRange =
    sheet.getRange(
      headerRow,
      1,
      1,
      numberOfColumns
    );


  headerRange
    .setFontWeight('bold');


  headerRange
    .setHorizontalAlignment(
      'center'
    );


  /************************************************
   * Daten
   ************************************************/

  if (
    numberOfRows > headerRow
  ) {

    sheet
      .getRange(
        headerRow + 1,
        1,
        numberOfRows - headerRow,
        numberOfColumns
      )
      .setVerticalAlignment(
        'middle'
      );
  }


  /************************************************
   * Spaltenbreiten
   ************************************************/

  const widths = [
    150, // Teilnehmer-ID
    180, // Kind
    220, // Workshop
    100, // Datum
    120, // Zeit
    180, // Elternteil
    240, // E-Mail
    130, // Status
    150, // Teilnahme
    130  // Zertifikat
  ];

  widths.forEach(
    (width, index) => {

      sheet.setColumnWidth(
        index + 1,
        width
      );

    }
  );

  /************************************************
   * Kopfzeilen fixieren
   ************************************************/

  sheet.setFrozenRows(
    headerRow
  );
}

function createManagementDropdowns(
  sheet,
  numberOfRows
) {

  if (numberOfRows < 6) {
    return;
  }


  const statusRule =
    SpreadsheetApp
      .newDataValidation()
      .requireValueInList(
        [
          'ANGEMELDET',
          'BESTÄTIGT',
          'ABGESAGT'
        ],
        true
      )
      .setAllowInvalid(false)
      .build();


  const participationRule =
    SpreadsheetApp
      .newDataValidation()
      .requireValueInList(
        [
          'OFFEN',
          'TEILGENOMMEN',
          'NICHT ERSCHIENEN'
        ],
        true
      )
      .setAllowInvalid(false)
      .build();


  const certificateRule =
    SpreadsheetApp
      .newDataValidation()
      .requireValueInList(
        [
          'OFFEN',
          'ERSTELLT',
          'GESENDET'
        ],
        true
      )
      .setAllowInvalid(false)
      .build();


  const rows =
    numberOfRows - 5;


  // Status – Spalte H
  sheet
    .getRange(
      6,
      8,
      rows,
      1
    )
    .setDataValidation(
      statusRule
    );


  // Teilnahme – Spalte I
  sheet
    .getRange(
      6,
      9,
      rows,
      1
    )
    .setDataValidation(
      participationRule
    );


  // Zertifikat – Spalte J
  sheet
    .getRange(
      6,
      10,
      rows,
      1
    )
    .setDataValidation(
      certificateRule
    );
}


function showParticipantList() {

  SpreadsheetApp.getUi().alert(
    'Die Teilnehmerliste wird in Phase 2.2 eingerichtet.'
  );
}


function showSystemStatus() {

  SpreadsheetApp.getUi().alert(
    'RaphNLab Teilnehmerverwaltung\n\n' +
    'Phase 1: ✓ Abgeschlossen\n' +
    'Phase 2: ✓ Gestartet\n' +
    'Phase 3: ○ Offen\n' +
    'Phase 4: ○ Offen'
  );
}
