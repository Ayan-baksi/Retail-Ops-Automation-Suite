// FormBuilder.gs
// Builds the Google Form: Welcome -> Store Information -> one page per
// checklist section -> Final Remarks -> Confirmation. Each section page
// carries a title plus a help-text instruction line (Forms has no
// separate "instructions" field type, so this is the standard equivalent).

function buildForm_() {
  var form = FormApp.create(CONFIG.FORM_TITLE);
  form.setDescription(CONFIG.FORM_DESC);
  form.setCollectEmail(false); // free-text identity, no org restriction
  form.setProgressBar(true);
  form.setConfirmationMessage(CONFIG.CONFIRMATION_MSG);
  form.setShowLinkToRespondAgain(true);

  // TODO: once CONFIG.LOGO_FILE_ID is set, apply it here via
  // form.setImage(DriveApp.getFileById(CONFIG.LOGO_FILE_ID).getBlob())

  // Store Information page
  form.addPageBreakItem()
    .setTitle('Store information')
    .setHelpText('Tell us who is visiting and where, before the checklist begins.');
  form.addTextItem().setTitle(Q.AM).setRequired(true);
  form.addTextItem().setTitle(Q.CLUSTER).setRequired(true);
  form.addListItem().setTitle(Q.STORE).setChoiceValues(CONFIG.STORES).setRequired(true);
  form.addDateItem().setTitle(Q.DATE).setRequired(true);

  // One page per checklist section
  SECTIONS.forEach(function (sec) {
    form.addPageBreakItem()
      .setTitle(sec.name)
      .setHelpText('Score each item 1 (compliant) or 0 (not compliant). Leave a row blank only if it genuinely does not apply to this store.');
    form.addGridItem()
      .setTitle(sec.name + ' — checklist')
      .setRows(sec.items)
      .setColumns(CONFIG.SCORE_COLS)
      .setRequired(false);
    form.addParagraphTextItem().setTitle(sec.name + ' — Remarks (optional)');
  });

  // Final remarks + confirmation
  form.addPageBreakItem()
    .setTitle('Final remarks')
    .setHelpText('Anything else worth flagging about this visit overall?');
  form.addParagraphTextItem().setTitle(Q.FINAL_REMARKS);

  return form;
}
