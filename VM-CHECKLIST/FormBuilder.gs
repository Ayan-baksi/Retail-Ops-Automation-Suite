// FormBuilder.gs
// Builds the VM Form: Store Info page, then one page per checklist
// section (each item rated 1-5 or Not Applicable, all required), a
// remarks box, then Final Remarks and Confirmation. No branching
// needed — every VM visit uses the same checklist.

function buildForm_() {
  var form = FormApp.create(CONFIG.FORM_TITLE);
  form.setDescription(CONFIG.FORM_DESC);
  form.setCollectEmail(false);
  form.setProgressBar(true);
  form.setConfirmationMessage(CONFIG.CONFIRMATION_MSG);
  form.setShowLinkToRespondAgain(true);

  form.addTextItem().setTitle(Q.VM).setRequired(true);
  form.addListItem().setTitle(Q.STORE).setChoiceValues(CONFIG.STORES).setRequired(true);
  form.addTextItem().setTitle(Q.SM).setRequired(true);
  form.addDateItem().setTitle(Q.DATE).setRequired(true);
  form.addListItem().setTitle(Q.ACTIVITY)
    .setChoiceValues(['Store Visit', 'Follow-up Visit', 'Special Audit', 'Training Visit'])
    .setRequired(true);

  SECTIONS.forEach(function (sec) {
    form.addPageBreakItem().setTitle(sec.name)
      .setHelpText('Rate every item 1 (Poor) to 5 (Excellent), or Not Applicable. All items on this page are mandatory.');
    sec.items.forEach(function (item) {
      form.addMultipleChoiceItem().setTitle(item).setChoiceValues(SCORE_CHOICES).setRequired(true);
    });
    form.addParagraphTextItem().setTitle(sec.name + ' — Remarks (optional)');
  });

  form.addPageBreakItem().setTitle('Final remarks')
    .setHelpText('Anything else worth flagging about this visit overall?');
  form.addParagraphTextItem().setTitle(Q.FINAL_REMARKS);

  return form;
}
