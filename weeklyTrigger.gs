function weeklyTrigger() {
  //https://developers.google.com/apps-script/reference/script/clock-trigger-builder
 ScriptApp.newTrigger("main")
  .timeBased()
  .everyWeeks(1)
  .onWeekDay(ScriptApp.WeekDay.MONDAY)
  .inTimezone("America/Los_Angeles")
  .create();
}
