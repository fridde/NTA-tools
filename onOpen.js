function onOpen() {
  let S = new SpreadsheetHelper(SpreadsheetApp.getActiveSpreadsheet());

  const menu = [
    ["Sammanställ kursdeltagare", "compileCourseParticipants"],
    ["Lägg till nya användare från bokningen", "addNewUsersFromKompetenskontroll"],
    ["Sammanställ lådstatus", "compileBoxStatus"],
    ["Fördela lådor på bokande", "compileRentalProposal"],
  ];

  S.createMenu("NTA-funktioner", menu);
}

function compileCourseParticipants() {
  const N = new NTA(new SpreadsheetHelper(SpreadsheetApp.getActiveSpreadsheet()));

  N.compileCourseParticipants();
}

function addNewUsersFromKompetenskontroll() {
  const N = new NTA(new SpreadsheetHelper(SpreadsheetApp.getActiveSpreadsheet()));

  N.addNewUsersFromKompetenskontroll();
}

function compileBoxStatus() {
  const N = new NTA(new SpreadsheetHelper(SpreadsheetApp.getActiveSpreadsheet()));

  N.compileBoxStatus();
}

function compileRentalProposal() {
  const N = new NTA(new SpreadsheetHelper(SpreadsheetApp.getActiveSpreadsheet()));

  N.compileRentalProposal();
}
