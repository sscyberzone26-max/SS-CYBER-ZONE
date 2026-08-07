const SPREADSHEET_ID =
  "1L9qM7ioM7x1hFzriVwcNiZ6Tj_W1MFf_9vbzHKYrZ0g";

const TOKEN_SHEET = "Sheet1";
const SETTINGS_SHEET = "Settings";


/* =====================================================
   GET
===================================================== */

function doGet(e) {

  try {

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    /* GET SETTINGS */

    if (
      e &&
      e.parameter &&
      e.parameter.action === "getSettings"
    ) {

      const settings = getSettings_(ss);

      return ContentService
        .createTextOutput(
          JSON.stringify({
            status: "success",
            settings: settings
          })
        )
        .setMimeType(ContentService.MimeType.JSON);
    }


    /* GET TOKEN DATA */

    const sheet =
      ss.getSheetByName(TOKEN_SHEET);

    if (!sheet) {

      return ContentService
        .createTextOutput(
          JSON.stringify({
            status: "error",
            message: "Sheet1 not found"
          })
        )
        .setMimeType(ContentService.MimeType.JSON);
    }


    const data =
      sheet.getDataRange().getValues();


    if (data.length <= 1) {

      return ContentService
        .createTextOutput("[]")
        .setMimeType(ContentService.MimeType.JSON);
    }


    const headers = data.shift();


    const result = data.map(function(row) {

      let obj = {};

      headers.forEach(function(header, index) {

        obj[header] = row[index];

      });

      return obj;

    });


    return ContentService
      .createTextOutput(
        JSON.stringify(result)
      )
      .setMimeType(ContentService.MimeType.JSON);


  } catch (error) {

    return json_({
      status: "error",
      message: error.toString()
    });

  }

}



/* =====================================================
   SETTINGS
===================================================== */

function getSettings_(ss) {

  let sheet =
    ss.getSheetByName(SETTINGS_SHEET);


  /* CREATE SETTINGS SHEET */

  if (!sheet) {

    sheet =
      ss.insertSheet(SETTINGS_SHEET);


    sheet.getRange("A1:B8").setValues([

      ["Setting", "Value"],

      ["openMode", "always"],

      ["startTime", ""],

      ["endTime", ""],

      ["startAmPm", "AM"],

      ["endAmPm", "AM"],

      ["slotLimit", ""],

      ["timeSlotLimit", ""]

    ]);

  }


  const data =
    sheet.getRange("A1:B8").getValues();


  let settings = {};


  for (let i = 1; i < data.length; i++) {

    const key = data[i][0];

    const value = data[i][1];


    if (key) {

      settings[key] = value;

    }

  }


  /* NUMBER CONVERSION */

  if (
    settings.slotLimit !== "" &&
    settings.slotLimit !== null &&
    settings.slotLimit !== undefined
  ) {

    settings.slotLimit =
      Number(settings.slotLimit);

  }


  if (
    settings.timeSlotLimit !== "" &&
    settings.timeSlotLimit !== null &&
    settings.timeSlotLimit !== undefined
  ) {

    settings.timeSlotLimit =
      Number(settings.timeSlotLimit);

  }


  return settings;

}



/* =====================================================
   SAVE SETTINGS
===================================================== */

function saveSettings_(ss, settings) {

  let sheet =
    ss.getSheetByName(SETTINGS_SHEET);


  if (!sheet) {

    getSettings_(ss);

    sheet =
      ss.getSheetByName(SETTINGS_SHEET);

  }


  let slotLimit = "";

  let timeSlotLimit = "";


  if (
    settings.slotLimit !== undefined &&
    settings.slotLimit !== null &&
    settings.slotLimit !== ""
  ) {

    slotLimit =
      Number(settings.slotLimit);

  }


  if (
    settings.timeSlotLimit !== undefined &&
    settings.timeSlotLimit !== null &&
    settings.timeSlotLimit !== ""
  ) {

    timeSlotLimit =
      Number(settings.timeSlotLimit);

  }


  sheet.getRange("A1:B8").setValues([

    ["Setting", "Value"],

    [
      "openMode",
      settings.openMode || "always"
    ],

    [
      "startTime",
      settings.startTime || ""
    ],

    [
      "endTime",
      settings.endTime || ""
    ],

    [
      "startAmPm",
      settings.startAmPm || "AM"
    ],

    [
      "endAmPm",
      settings.endAmPm || "AM"
    ],

    [
      "slotLimit",
      slotLimit
    ],

    [
      "timeSlotLimit",
      timeSlotLimit
    ]

  ]);


  return getSettings_(ss);

}



/* =====================================================
   POST
===================================================== */

function doPost(e) {

  try {

    const ss =
      SpreadsheetApp.openById(
        SPREADSHEET_ID
      );


    if (!e || !e.postData || !e.postData.contents) {

      throw new Error(
        "POST data not found"
      );

    }


    const data =
      JSON.parse(
        e.postData.contents
      );



    /* =================================================
       ADMIN SETTINGS SAVE
    ================================================= */

    if (
      data.action === "saveSettings"
    ) {

      const settings =
        saveSettings_(
          ss,
          data.settings || {}
        );


      return json_({

        status: "success",

        settings: settings

      });

    }



    /* =================================================
       TOKEN GENERATION
    ================================================= */

    const sheet =
      ss.getSheetByName(TOKEN_SHEET);


    if (!sheet) {

      throw new Error(
        "Sheet1 not found"
      );

    }


    const name =
      String(
        data.name || ""
      ).trim();


    const phone =
      String(
        data.phone || ""
      ).trim();


    const address =
      String(
        data.address || ""
      ).trim();


    const visit =
      String(
        data.visit || ""
      ).trim();


    const aadhaar =
      String(
        data.aadhaar || ""
      ).trim();



    /* =================================================
       PHONE / AADHAAR DUPLICATE CHECK
    ================================================= */

    const allData =
      sheet.getDataRange().getValues();


    if (allData.length > 1) {

      const headers =
        allData[0];


      const phoneIndex =
        headers.indexOf("Phone");


      const aadhaarIndex =
        headers.indexOf("Aadhaar");



      /* PHONE CHECK */

      if (phoneIndex !== -1) {

        for (
          let i = 1;
          i < allData.length;
          i++
        ) {

          if (
            String(
              allData[i][phoneIndex]
            ).trim() === phone &&
            phone !== ""
          ) {

            return json_({

              status:
                "duplicate_phone"

            });

          }

        }

      }



      /* AADHAAR CHECK */

      if (aadhaarIndex !== -1) {

        for (
          let i = 1;
          i < allData.length;
          i++
        ) {

          if (
            String(
              allData[i][aadhaarIndex]
            ).trim() === aadhaar &&
            aadhaar !== ""
          ) {

            return json_({

              status:
                "duplicate_aadhaar"

            });

          }

        }

      }

    }



    /* =================================================
       TOKEN NUMBER
    ================================================= */

    const token =
      getNextToken_(sheet);



    /* =================================================
       ADD ROW
    ================================================= */

    const now =
      Utilities.formatDate(
        new Date(),
        Session.getScriptTimeZone(),
        "dd/MM/yyyy HH:mm:ss"
      );


    sheet.appendRow([

      token,

      name,

      phone,

      address,

      visit,

      aadhaar,

      now

    ]);



    return json_({

      status: "success",

      token: token

    });


  } catch (error) {

    return json_({

      status: "error",

      message: error.toString()

    });

  }

}



/* =====================================================
   GENERATE NEXT TOKEN
===================================================== */

function getNextToken_(sheet) {

  const today =
    Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "ddMMyyyy"
    );


  let serial = 1;


  const lastRow =
    sheet.getLastRow();


  if (lastRow > 1) {

    const values =
      sheet
        .getRange(
          2,
          1,
          lastRow - 1,
          1
        )
        .getValues();


    for (
      let i = values.length - 1;
      i >= 0;
      i--
    ) {

      const oldToken =
        String(values[i][0] || "");


      if (
        oldToken.includes(today)
      ) {

        const parts =
          oldToken.split("/");


        if (parts.length >= 2) {

          const oldSerial =
            parseInt(
              parts[parts.length - 1],
              10
            );


          if (!isNaN(oldSerial)) {

            serial =
              oldSerial + 1;

          }

        }


        break;

      }

    }

  }


  return (
    "SScyberzone " +
    today +
    "/" +
    ("000" + serial).slice(-3)
  );

}



/* =====================================================
   JSON RESPONSE
===================================================== */

function json_(data) {

  return ContentService

    .createTextOutput(
      JSON.stringify(data)
    )

    .setMimeType(
      ContentService.MimeType.JSON
    );

}