// FormBuilder.gs
// Builds the branching SM Form:
//   Store Info + Audit Type (branches here)
//     -> Opening: 6 section pages, each ending in optional Remarks
//     -> Closing: 1 section page, ending in optional Remarks
//   -> Final Remarks -> Confirmation
//
// Every checklist question is a required Yes/No/Not Applicable choice.
// Items flagged image:true get a required file-link question beneath
// them — Google Forms requires responders to be signed into a Google
// account for native file uploads, so this uses a text link field instead.

function buildForm_() {
  var form = FormApp.create(CONFIG.FORM_TITLE);
  form.setDescription(CONFIG.FORM_DESC);
  form.setCollectEmail(false);
  form.setProgressBar(true);
  form.setConfirmationMessage(CONFIG.CONFIRMATION_MSG);
  form.setShowLinkToRespondAgain(true);

  // Store Information page
  form.addTextItem().setTitle(Q.SM).setRequired(true);
  form.addListItem().setTitle(Q.STORE).setChoiceValues(CONFIG.STORES).setRequired(true);
  form.addDateItem().setTitle(Q.DATE).setRequired(true);
  var auditTypeItem = form.addMultipleChoiceItem().setTitle(Q.AUDIT_TYPE).setRequired(true);

  // Opening section pages, tracked for branching
  var openingPageBreaks = [];
  SECTIONS_OPENING.forEach(function (sec, idx) {
    var pb = form.addPageBreakItem().setTitle(sec.name)
      .setHelpText(buildSectionHelpText_(sec));
    openingPageBreaks.push(pb);
    buildSectionQuestions_(form, sec);
  });

  // Closing section page
  var closingPageBreaks = [];
  SECTIONS_CLOSING.forEach(function (sec) {
    var pb = form.addPageBreakItem().setTitle(sec.name)
      .setHelpText(buildSectionHelpText_(sec));
    closingPageBreaks.push(pb);
    buildSectionQuestions_(form, sec);
  });

  // Final remarks + confirmation
  var finalPage = form.addPageBreakItem().setTitle('Final remarks')
    .setHelpText('Anything else worth flagging about this shift overall?');
  form.addParagraphTextItem().setTitle(Q.FINAL_REMARKS);

  // Branching: Audit Type sends the responder to the first page of
  // whichever branch they picked.
  auditTypeItem.setChoices([
    auditTypeItem.createChoice('Opening', openingPageBreaks[0]),
    auditTypeItem.createChoice('Closing', closingPageBreaks[0])
  ]);
  // The last Opening page must explicitly skip past the Closing page
  // (next in creation order) to Final Remarks, or Forms' default
  // "next page in sequence" would route an Opening responder into the
  // Closing questions. Closing already falls through to Final Remarks
  // by default since it's next in creation order, so no override needed there.
  openingPageBreaks[openingPageBreaks.length - 1].setGoToPage(finalPage);

  return form;
}

function buildSectionHelpText_(sec) {
  var txt = 'Every question on this page is mandatory (Yes / No / Not Applicable). ' +
    'Required photos must be uploaded before you can continue.';
  if (sec.note) txt += '\n\nNote: ' + sec.note;
  return txt;
}

function buildSectionQuestions_(form, sec) {
  sec.items.forEach(function (item) {
    form.addMultipleChoiceItem()
      .setTitle(item.text)
      .setChoiceValues(['Yes', 'No', 'Not Applicable'])
      .setRequired(true);
    if (item.image) {
      form.addTextItem()
        .setTitle(item.text + ' — Photo Link' + (item.imageNote ? ' (' + item.imageNote + ')' : ''))
        .setHelpText('Upload the photo to Google Drive or Google Photos, set sharing to "Anyone with the link", then paste that link here.')
        .setRequired(true);
    }
  });
  form.addParagraphTextItem().setTitle(sec.name + ' — Remarks (optional)');
}
