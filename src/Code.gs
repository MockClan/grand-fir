/**
 * Change these to match the column names you are using for email
 * recipient addresses and email sent column.
 */
const RECIPIENT_COL = 'Email Address';
const EMAIL_SENT_COL = 'Follow-up';

/**
 * Creates the menu item "Mail Merge" for user to run scripts on drop-down.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Troop Tools')
      .addItem('Send Confirmation Emails', 'sendConfirmationEmails')
      .addItem('Resolve USPS Addresses & Geocodes', 'resolveAddresses')
  // .addItem("Normlize Phone Numbers", "normalizePhoneNumbers")
      .addToUi();
}

/**
 * Sends the Christmas Tree confirmation emails for the current worksheet.
 */
function sendConfirmationEmails() {
  sendEmails('[CONFIRMATION] Christmas Tree Recycling with BSA Troop 246');
}

/**
 * Performs a Google Maps Geocode lookup to resolve the geocode for a given address.
 * @param {string} address The address to lookup
 * @return {object} An object repesenting the response with latitude and longitude.
 */
function getGeoLocation(address) {
  try {
    var loc = Maps.newGeocoder().geocode(address);
    if (loc.results[0].geometry.location_type == 'ROOFTOP') {
      return {
        latitude: loc.results[0].geometry.location.lat,
        longitude: loc.results[0].geometry.location.lng
      };
    } else {
      return null;
    }
  } catch (err) {
    throw new Error(err.message);
  }
}

/**
 * Function attempts to normalize the Phone Numbers column for the current worksheet.
 * @param {sheet} sheet The worksheet to process (the current active by default if not specified)
 */
function normalizePhoneNumbers(sheet = SpreadsheetApp.getActiveSheet()) {
  const data = sheet.getDataRange().getValues();

  // Assumes row 1 contains our column headings
  const heads = data.shift();

  // Gets the index of the column named 'Email Status' (Assumes header names are unique)
  // @see http://ramblings.mcpher.com/Home/excelquirks/gooscript/arrayfunctions
  const geocodeColIdx = heads.indexOf('Phone Number');

  // Converts 2d array into an object array
  // See https://stackoverflow.com/a/22917499/1027723
  // For a pretty version, see https://mashe.hawksey.info/?p=17869/#comment-184945
  const obj = data.map((r) =>
    heads.reduce((o, k, i) => ((o[k] = r[i] || ''), o), {})
  );

  // Creates an array to record sent emails
  const out = [];

  // Loops through all the rows of data
  obj.forEach(function(row, rowIdx) {
    // Only process rows where the geoCode hasn't been resolved
    if (row['Phone Number'] != '') {
      try {
        var originalPhoneNumber = row['Phone Number'];
        var phoneNumber = originalPhoneNumber
            .replace('/\)/g', '')
            .replace('/\)/g', '')
            .replace('/\s/g', '')
            .replace('/-/g', '');

        Logger.log('@' + (rowIdx + 2) + ' ' + originalPhoneNumber + ' -> ' + phoneNumber);
        out.push([parseInt(phoneNumber)]);
      } catch (e) {
        Logger.log(e.message);
        out.push([row['Phone Number']]);
      }
    } else {
      out.push([row['Phone Number']]);
    }
  });

  // Updates the sheet with new data
  sheet.getRange(2, geocodeColIdx + 1, out.length).setValues(out);
}

/**
 * Performs both the full USPS address lookup and geocode lookup for rows with
 * empty values for the given sheet.
 * @param {sheet} sheet The worksheet to process (current active worksheet it not specified).
 */
function resolveAddresses(sheet = SpreadsheetApp.getActiveSheet()) {
  resolveUspsAddresses(sheet);
  resolveGeoLocations(sheet);
}

/**
 * Iterates through all of the uspsStreet and uspsZip columns for the given sheet,
 * resolves the USPS defined address, and populates those columns with the results.
 * @param {sheet} sheet The worksheet to process (current active worksheet it not specified).
 */
function resolveUspsAddresses(sheet = SpreadsheetApp.getActiveSheet()) {
  const data = sheet.getDataRange().getValues();

  // Assumes row 1 contains our column headings
  const heads = data.shift();

  // Gets the index of the column named 'Email Status' (Assumes header names are unique)
  // @see http://ramblings.mcpher.com/Home/excelquirks/gooscript/arrayfunctions
  const streetColIdx = heads.indexOf('uspsStreet');
  const zipColIdx = heads.indexOf('uspsZip');

  // Converts 2d array into an object array
  // See https://stackoverflow.com/a/22917499/1027723
  // For a pretty version, see https://mashe.hawksey.info/?p=17869/#comment-184945
  const obj = data.map((r) =>
    heads.reduce((o, k, i) => ((o[k] = r[i] || ''), o), {})
  );

  // Loops through all the rows of data
  obj.forEach(function(row, rowIdx) {
    if (row['Street Address'] != '' && (row['uspsStreet'] == '' || row['uspsZip'] == '')) {
      Logger.log('Update USPS Address for row ' + rowIdx + ': ' +JSON.stringify({
        address: row['Street Address'],
        city: 'Colorado Springs',
        state: 'CO',
        zip: row['ZIP Code']
      }));

      try {
        var result = Usps.addressLookup(row['Street Address'], 'Colorado Springs', 'CO');
        Logger.log('Response: ' + JSON.stringify(result));
        if (result !== null) {
          sheet.getRange(rowIdx+2, streetColIdx+1).setValue(result.address2);
          sheet.getRange(rowIdx+2, zipColIdx+1).setValue(result.zip5 + '-' + result.zip4);
        }
      } catch (e) {
        throw new Error(e.message);
      }
    }
  });
}

/**
 * Iterates through the geocode columne for the given sheet, resolves the geocode using google maps,
 * and populates those columns with the results.
 * @param {sheet} sheet The worksheet to process (current active worksheet it not specified).
 */
function resolveGeoLocations(sheet = SpreadsheetApp.getActiveSheet()) {
  // Gets the data from the passed sheet
  const dataRange = sheet.getDataRange();
  const data = dataRange.getDisplayValues();
  const heads = data.shift();
  const geocodeColIdx = heads.indexOf('geocode');

  // Converts the current table values into a 2-D array into an object array
  const obj = data.map((r) =>
    heads.reduce((o, k, i) => ((o[k] = r[i] || ''), o), {})
  );

  // The array used to populate the geocode column with updated values
  const out = [];

  // Loops through all the rows of data
  obj.forEach(function(row, rowIdx) {
    // Only process rows where the geoCode hasn't been resolved
    if (
      row['geocode'] == '' &&
      row['uspsStreet'] != '' &&
      row['uspsZip'] != ''
    ) {
      try {
        var geocode = getGeoLocation(row['uspsStreet'] + ', ' + row['uspsZip']);
        if (geocode == null) {
          out.push(['<ERROR>']);
        } else {
          out.push([geocode.latitude + ',' + geocode.longitude]);
        }
      } catch (e) {
        out.push([e.message]);
      }
    } else {
      out.push([row['geocode']]);
    }
  });

  // Updates the sheet with new data
  sheet.getRange(2, geocodeColIdx + 1, out.length).setValues(out);
}
