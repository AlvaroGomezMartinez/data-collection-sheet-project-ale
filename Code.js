/**
 * Copies notes from a source range to multiple target ranges in a Google Sheets document.
 * This is a utility function used once to copy notes from the first day's check-in occurrences to all other relevant ranges in the Week_Template sheet.
 */
function copyNotesToMultipleNamedRanges() {
  var spreadsheetId = "1szS07tM9EM9e8vdmS3AzcFkv3mXDd91jzeGxnBwk7H8";
  var sheetName = "Week_Template";
  var ss = SpreadsheetApp.openById(spreadsheetId);
  var sheet = ss.getSheetByName(sheetName);

  // Get the source range and its notes
  var sourceRange = ss.getRangeByName("day1CheckInOccurencesRatingsNotes");
  var sourceNotes = sourceRange.getNotes();

  // Define the target range names where notes will be copied
  var targetRangeNames = [
    "day11stOccurencesRatingsNotes", "day12ndOccurencesRatingsNotes", "day13rdOccurencesRatingsNotes", "day14thOccurencesRatingsNotes", "day15thOccurencesRatingsNotes", "day16thOccurencesRatingsNotes", "day17thOccurencesRatingsNotes", "day18thOccurencesRatingsNotes", "day19thOccurencesRatingsNotes", "day1CheckOutOccurencesRatingsNotes", "day21stOccurencesRatingsNotes", "day22ndOccurencesRatingsNotes", "day23rdOccurencesRatingsNotes", "day24thOccurencesRatingsNotes", "day25thOccurencesRatingsNotes", "day26thOccurencesRatingsNotes", "day27thOccurencesRatingsNotes", "day28thOccurencesRatingsNotes", "day29thOccurencesRatingsNotes", "day2CheckInOccurencesRatingsNotes", "day2CheckOutOccurencesRatingsNotes", "day31stOccurencesRatingsNotes", "day32ndOccurencesRatingsNotes", "day33rdOccurencesRatingsNotes", "day34thOccurencesRatingsNotes", "day35thOccurencesRatingsNotes", "day36thOccurencesRatingsNotes", "day37thOccurencesRatingsNotes", "day38thOccurencesRatingsNotes", "day39thOccurencesRatingsNotes", "day3CheckInOccurencesRatingsNotes", "day3CheckOutOccurencesRatingsNotes", "day41stOccurencesRatingsNotes", "day42ndOccurencesRatingsNotes", "day43rdOccurencesRatingsNotes", "day44thOccurencesRatingsNotes", "day45thOccurencesRatingsNotes", "day46thOccurencesRatingsNotes", "day47thOccurencesRatingsNotes", "day48thOccurencesRatingsNotes", "day49thOccurencesRatingsNotes", "day4CheckInOccurencesRatingsNotes", "day4CheckOutOccurencesRatingsNotes", "day51stOccurencesRatingsNotes", "day52ndOccurencesRatingsNotes", "day53rdOccurencesRatingsNotes", "day54thOccurencesRatingsNotes", "day55thOccurencesRatingsNotes", "day56thOccurencesRatingsNotes", "day57thOccurencesRatingsNotes", "day58thOccurencesRatingsNotes", "day59thOccurencesRatingsNotes", "day5CheckInOccurencesRatingsNotes", "day5CheckOutOccurencesRatingsNotes"
  ];

  // Iterate over each target range and copy the notes
  targetRangeNames.forEach(function(rangeName) {
    var targetRange = ss.getRangeByName(rangeName);
    var targetNotes = targetRange.getNotes();

    // Copy notes cell by cell
    for (var i = 0; i < sourceNotes.length; i++) {
      for (var j = 0; j < sourceNotes[i].length; j++) {
        targetNotes[i][j] = sourceNotes[i][j];
      }
    }
    targetRange.setNotes(targetNotes);
  });
}