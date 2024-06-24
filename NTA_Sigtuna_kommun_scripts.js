class NTA 
{
  /**
   * @param {SpreadsheetHelper} sh
   */
  constructor(sh)
  {
    this.S = sh;
    this.Logger = BetterLog.useSpreadsheet();
  }

  compileCourseParticipants()
  {   
    let utdata = [];

    this.getKompetensKontroll().forEach(row => {
      let topic = this.getTopicFromLongName(row[2]);
      let mails = this.extractMailsFromRow(row);

      mails.forEach(mail => {
        let user = this.findUser(mail);
        if(user === undefined){
          this.S.alert("User with adress " + mail + " doesn't exist.");
          return;
        }

        [topic, 'I'].forEach(t => {
          if(!(this.userHasCompetence(user, t))){
            let newRow = [];
            newRow.push(user['Förnamn'] + ' ' + user['Efternamn']);
            newRow.push(mail);
            newRow.push(this.getSchoolFromId(user['Skola'])); //Skola
            newRow.push('Sigtuna'); // Kommun
            newRow.push(t); //Kurs
            
            utdata.push(newRow);
          }
        });      
      });
    });
    utdata = this.S.uniqueArray(utdata);
    this.S.insertValuesAt('Utdata', [1,1], utdata);
    
  }

  addNewUsersFromKompetenskontroll()
  {
    let maxId = Math.max(...this.getPersoner().map(p => p['id']));

    this.getKompetensKontroll().forEach(row => {
      let mails = this.extractMailsFromRow(row);      

      mails.forEach(mail => {
        if(this.findUser(mail) === undefined){
          maxId = maxId +1;
          this.personer.push({
            'id': maxId, 
            'Förnamn': '', 
            'Efternamn': '', 
            'Skola': this.getSchoolFromLongName(row[1]), 
            'Mejl': mail
            });
        };
      });
    });
    this.S.insertValuesAt('Personer', [2,1], personer);
  }

  compileBoxStatus()
  {
     let updates = this.getBoxUpdates().map(update => {
      update['Låda'] = this.standardizeBoxId(update['Låda']);
      update['Datum'] = new Date(update['Datum']);
      return update;
    });

    const givenDateText = SpreadsheetApp.getUi().prompt('Vilket datum?').getResponseText();

    if(givenDateText !== ''){
      givenDate = new Date(givenDateText);
      updates = updates.filter(row => {
        return row['Datum'].getTime() <= givenDate.getTime();
      });
    }

    updates.sort((row1, row2) => row2['Datum'].getTime() - row1['Datum'].getTime());
    
    const output = [];
    this.getBoxes().map(row => this.standardizeBoxId(row['id'])).forEach(id => {
      const newestUpdate = updates.find(update => update['Låda'] === id);
      const newestRental = updates.find(update => update['Låda'] === id && update['Status'] === 'uthyrd');
      //BetterLog.log(newestRental);
      if(newestUpdate === undefined){
        this.S.alert('Lådan med id "' +id + '" har aldrig uppdaterats.');
      }
      const statusObj = {'id': this.prettifyBoxId(id), 'Status': newestUpdate['Status'], 'Låntagare': ''};
      if(newestRental !== undefined){
        const u = this.getUserById(Number(newestRental['Person']));
        statusObj['Låntagare'] = `${u['Förnamn']} ${u['Efternamn']}, ${u['Skola'].toUpperCase()}`;
      }       

      output.push(statusObj);
    });
      
    const html = HtmlService.createHtmlOutput(this.S.convertToHTMLTable(this.S.convertToGridWithHeaders(output)));
    html.setHeight(700).setWidth(400);

    SpreadsheetApp.getUi().showSidebar(html);
  }

  extractMailsFromRow(row)
  {
      const mails = [row[0]];
      if(row[4].length > 0){
        mails.push(...(row[4].split(';')));
      }
      return mails.map(m => m.trim().toLowerCase());
  }

  /**
   * @param {String} irregularId
   * 
   * @return {String} - The regular lowercase id without any special characters
   */
  standardizeBoxId(irregularId)
  {
    const re = /\W/g;
    return irregularId.toLowerCase().replaceAll(re, '').replaceAll('_', '');   // the lowerscore is unfortunately part of \w , so we have to remove it separately
  }

  /**
   * @param {String} regularId - a lowercase-id without any special symbols in the format aa11b (small letters followed by numbers and, optional, letters)
   * 
   * @return {String} a well formatted id in the format AA:11:B
   */
  prettifyBoxId(regularId)
  {
    const parts = Array.from(regularId.matchAll(/(\D+)(\d+)(\D*)/g))[0];
    //this.Logger.log(JSON.stringify(parts));
    parts.shift();

    return parts.filter(p => p.length > 0).map(p => p.toUpperCase()).join(':');
  }


  /**
   * @param {string} mail - The mail adress, trimmed and lowercase
   * 
   * @return {Object|undefined}
   */
  findUser(mail)
  {
    const foundUsers = this.getPersoner().filter(person => person["Mejl"].toLowerCase() === mail);

    if(foundUsers.length > 2){
      this.S.alert("Flera användare hittades för " + mail);
      return;
    }
    if(foundUsers.length === 0){
      return;
    }
    return foundUsers.shift();
  }

  /**
   * @param {Number} id - The user id
   * 
   * @return {Object} 
   */
  getUserById(id)
  {
    return this.getPersoner().find(user => user['id'] === id);
  }

  userHasCompetence(user, topic){
    return this.getKompetens().find(row => {
      return (row['Person_id'] === user['id']) && (row['Tema_id'] === topic)
    }) !== undefined;
  }

  getTopicFromLongName(longName){
    let foundTopic = this.getTeman().find(topic => topic['Temanamn'] === longName);
    if(foundTopic === undefined){
      this.S.alert(longName + " finns inte med som tema.");
      return;
    }
    return foundTopic["id"];
  }

  getSchoolFromLongName(longName){
    let foundSchool = this.getSkolor().find(school => school['Alias'] === longName);
  
    if(foundSchool === undefined){
      this.S.alert(longName + " finns inte med som skola.");
      return;
    }
    return foundSchool["id"];
  }

  getSchoolFromId(schoolId){
    let foundSchool = this.getSkolor().find(school => school['id'] === schoolId);
    
    if(foundSchool === undefined){
      this.S.alert(schoolId + " finns inte med som skola.");
      return;
    }
    return foundSchool["Alias"];
  }

  

  /**
   * @return {Array<Object>}
   */
  getPersoner()
  {
    if(this.personer === undefined) {
      this.personer = this.S.firstRowAsHeader(this.S.getNamedValues('Personer'));
    }
    return this.personer;  
  }

  /**
   * @return {Array<Object>}
   */
  getKompetens()
  {
    if(this.kompetens === undefined) {
     this.kompetens = this.S.firstRowAsHeader(this.S.getNamedValues('Kompetens'));
    }
    return this.kompetens;
  }

  /**
   * @return {Array<Object>}
   */
  getTeman()
  {
    if(this.teman === undefined) {
     this.teman = this.S.firstRowAsHeader(this.S.getNamedValues('Teman'));
    }
    return this.teman;
  }

  /**
   * @return {Array<Object>}
   */
  getSkolor()
  {
    if(this.skolor === undefined) {
     this.skolor = this.S.firstRowAsHeader(this.S.getNamedValues('Skolor'));
    }
    return this.skolor;
  }

  /**
   * @return {Array<Object>}
   */
  getKompetensKontroll()
  {
    if(this.kompetenskontroll === undefined) {
     this.kompetenskontroll = this.S.removeIfEmpty(this.S.getNamedValues('Kompetenskontroll'));
    }
    return this.kompetenskontroll;
  }

   /**
   * @return {Array<Object>}
   */
  getBoxes()
  {
    if(this.boxes === undefined) {
      this.boxes = this.S.firstRowAsHeader(this.S.removeIfEmpty(this.S.getNamedValues('Lådor')));
    }
    return this.boxes;  
  }

  /**
   * @return {Array<Object>}
   */
  getBoxUpdates()
  {
    if(this.boxUpdates === undefined) {
      this.boxUpdates = this.S.firstRowAsHeader(this.S.removeIfEmpty(this.S.getNamedValues('Låduppdateringar')));
    }
    return this.boxUpdates;  
  }
}



