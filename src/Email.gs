/**
 * Sends emails from sheet data.
 * @param {string} subjectLine (optional) for the email draft message
 * @param {Sheet} sheet to read data from
 */
function sendEmails(subjectLine, sheet = SpreadsheetApp.getActiveSheet()) {
  // option to skip browser prompt if you want to use this code in other projects
  if (!subjectLine) {
    subjectLine = Browser.inputBox(
        'Mail Merge',
        'Type or copy/paste the subject line of the Gmail ' +
        'draft message you would like to mail merge with:',
        Browser.Buttons.OK_CANCEL
    );

    if (subjectLine === 'cancel' || subjectLine == '') {
      // If no subject line, finishes up
      return;
    }
  }

  // Gets the draft Gmail message to use as a template
  const emailTemplate = getGmailTemplateFromDrafts_(subjectLine);

  const dataRange = sheet.getDataRange();
  const data = dataRange.getDisplayValues();

  // Assumes row 1 contains our column headings
  const heads = data.shift();

  // Gets the index of the column named 'Email Status' (Assumes header names are unique)
  // @see http://ramblings.mcpher.com/Home/excelquirks/gooscript/arrayfunctions
  const emailSentColIdx = heads.indexOf(EMAIL_SENT_COL);

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
    // Only sends emails if email_sent cell is blank and not hidden by a filter
    if (row[RECIPIENT_COL] != '' && row[EMAIL_SENT_COL] == '') {
      try {
        const msgObj = fillInTemplateFromObject_(emailTemplate.message, row);

        // If you need to send emails with unicode/emoji characters change GmailApp for MailApp
        // Uncomment advanced parameters as needed (see docs for limitations)
        GmailApp.sendEmail(row[RECIPIENT_COL], msgObj.subject, msgObj.text, {
          htmlBody: msgObj.html,
          // bcc: 'a.bbc@email.com',
          // cc: 'a.cc@email.com',
          // from: 'an.alias@email.com',
          name: 'BSA Troop 246',
          replyTo: 'recycletrees@bsatroop246.org',
          attachments: emailTemplate.attachments,
          inlineImages: emailTemplate.inlineImages
        });
        // Edits cell to record email sent date
        const now = new Date();
        out.push([
          'Mail Merge @ ' +
            now.getMonth() +
            '/' +
            now.getDay() +
            '/' +
            now.getFullYear()
        ]);
      } catch (e) {
        // modify cell to record error
        out.push([e.message]);
      }
    } else {
      out.push([row[EMAIL_SENT_COL]]);
    }
  });

  // Updates the sheet with new data
  sheet.getRange(2, emailSentColIdx + 1, out.length).setValues(out);

  /**
   * Get a Gmail draft message by matching the subject line.
   * @param {string} subjectLine to search for draft message
   * @return {object} containing the subject, plain and html message body and attachments
   */
  function getGmailTemplateFromDrafts_(subjectLine) {
    try {
      // get drafts
      const drafts = GmailApp.getDrafts();
      // filter the drafts that match subject line
      const draft = drafts.filter(subjectFilter_(subjectLine))[0];
      // get the message object
      const msg = draft.getMessage();

      // Handles inline images and attachments so they can be included in the merge
      // Based on https://stackoverflow.com/a/65813881/1027723
      // Gets all attachments and inline image attachments
      const allInlineImages = draft
          .getMessage()
          .getAttachments({
            includeInlineImages: true,
            includeAttachments: false
          });
      const attachments = draft
          .getMessage()
          .getAttachments({includeInlineImages: false});
      const htmlBody = msg.getBody();

      // Creates an inline image object with the image name as key
      // (can't rely on image index as array based on insert order)
      const imgObject = allInlineImages.reduce(
          (obj, i) => ((obj[i.getName()] = i), obj),
          {}
      );

      // Regexp searches for all img string positions with cid
      const imgexp = RegExp('<img.*?src="cid:(.*?)".*?alt="(.*?)"[^>]+>', 'g');
      const matches = [...htmlBody.matchAll(imgexp)];

      // Initiates the allInlineImages object
      const inlineImagesObj = {};
      // built an inlineImagesObj from inline image matches
      matches.forEach(
          (match) => (inlineImagesObj[match[1]] = imgObject[match[2]])
      );

      return {
        message: {
          subject: subjectLine,
          text: msg.getPlainBody(),
          html: htmlBody
        },
        attachments: attachments,
        inlineImages: inlineImagesObj
      };
    } catch (e) {
      throw new Error('Oops - can\'t find Gmail draft');
    }

    /**
     * Filter draft objects with the matching subject linemessage by matching the subject line.
     * @param {string} subjectLine to search for draft message
     * @return {object} GmailDraft object
     */
    function subjectFilter_(subjectLine) {
      return function(element) {
        if (element.getMessage().getSubject() === subjectLine) {
          return element;
        }
      };
    }
  }

  /**
   * Fill template string with data object
   * @see https://stackoverflow.com/a/378000/1027723
   * @param {string} template string containing {{}} markers which are replaced with data
   * @param {object} data object used to replace {{}} markers
   * @return {object} message replaced with data
   */
  function fillInTemplateFromObject_(template, data) {
    // We have two templates one for plain text and the html body
    // Stringifing the object means we can do a global replace
    let templateString = JSON.stringify(template);

    // Token replacement
    templateString = templateString.replace(/{{[^{}]+}}/g, (key) => {
      return escapeData_(data[key.replace(/[{}]+/g, '')] || '');
    });
    return JSON.parse(templateString);
  }

  /**
   * Escape cell data to make JSON safe
   * @see https://stackoverflow.com/a/9204218/1027723
   * @param {string} str to escape JSON special characters from
   * @return {string} escaped string
   */
  function escapeData_(str) {
    return str
        .replace(/[\\]/g, '\\\\')
        .replace(/[\"]/g, '\\"')
        .replace(/[\/]/g, '\\/')
        .replace(/[\b]/g, '\\b')
        .replace(/[\f]/g, '\\f')
        .replace(/[\n]/g, '\\n')
        .replace(/[\r]/g, '\\r')
        .replace(/[\t]/g, '\\t');
  }
}
