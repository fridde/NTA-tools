function onOpen() {

  let S = new SpreadsheetHelper(SpreadsheetApp.getActiveSpreadsheet());
  
  const menu = [
    ['Sammanställ kursdeltagare', 'compileCourseParticipants'],
    ['Lägg till nya användare från bokningen', 'addNewUsersFromKompetenskontroll'],
    ['Sammanställ lådstatus', 'compileBoxStatus']
  ];
  
  S.createMenu('NTA-funktioner', menu); 
}

function compileCourseParticipants()
{
  let N = new NTA(new SpreadsheetHelper(SpreadsheetApp.getActiveSpreadsheet()));

  N.compileCourseParticipants();
}

function addNewUsersFromKompetenskontroll()
{
  let N = new NTA(new SpreadsheetHelper(SpreadsheetApp.getActiveSpreadsheet()));

  N.addNewUsersFromKompetenskontroll();
}

function compileBoxStatus()
{
  let N = new NTA(new SpreadsheetHelper(SpreadsheetApp.getActiveSpreadsheet()));

  N.compileBoxStatus();
}