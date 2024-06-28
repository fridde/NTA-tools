class NTA {
  /**
   * @param {SpreadsheetHelper} sh
   */
  constructor(sh) {
    this.S = sh;
    this.Logger = BetterLog.useSpreadsheet();
  }

  compileCourseParticipants() {
    let utdata = [];

    this.getKompetensKontroll().forEach(row => {
      let topic = this.getTopicFromLongName(row[2]);
      let mails = this.extractMailsFromRow(row);

      mails.forEach(mail => {
        let user = this.findUser(mail);
        if (user === undefined) {
          this.S.alert("User with adress " + mail + " doesn't exist.");
          return;
        }

        [topic, "I"].forEach(t => {
          if (!this.userHasCompetence(user, t)) {
            let newRow = [];
            newRow.push(user["Förnamn"] + " " + user["Efternamn"]);
            newRow.push(mail);
            newRow.push(this.getSchoolFromId(user["Skola"])); //Skola
            newRow.push("Sigtuna"); // Kommun
            newRow.push(t); //Kurs

            utdata.push(newRow);
          }
        });
      });
    });
    utdata = this.S.uniqueArray(utdata);
    this.S.insertValuesAt("Utdata", [1, 1], utdata);
  }

  addNewUsersFromKompetenskontroll() {
    let maxId = Math.max(...this.getPersoner().map(p => p["id"]));

    this.getKompetensKontroll().forEach(row => {
      let mails = this.extractMailsFromRow(row);

      mails.forEach(mail => {
        if (this.findUser(mail) === undefined) {
          maxId = maxId + 1;
          this.personer.push({
            id: maxId,
            Förnamn: "",
            Efternamn: "",
            Skola: this.getSchoolFromLongName(row[1]),
            Mejl: mail,
          });
        }
      });
    });
    this.S.insertValuesAt("Personer", [2, 1], personer);
  }

  compileBoxStatus() {
    let updates = this.getBoxUpdates().map(update => {
      update["Låda"] = this.standardizeBoxId(update["Låda"]);
      update["Datum"] = new Date(update["Datum"]);
      return update;
    });

    const givenDateText = SpreadsheetApp.getUi().prompt("Vilket datum?").getResponseText();

    if (givenDateText !== "") {
      const givenDate = new Date(givenDateText);
      updates = updates.filter(row => {
        return row["Datum"].getTime() <= givenDate.getTime();
      });
    }

    updates.sort((row1, row2) => row2["Datum"].getTime() - row1["Datum"].getTime());

    const output = [];
    this.getBoxes()
      .map(row => this.standardizeBoxId(row["id"]))
      .forEach(id => {
        const newestUpdate = updates.find(update => update["Låda"] === id);
        const newestRental = updates.find(update => update["Låda"] === id && update["Status"].startsWith("uthyrning"));
        //BetterLog.log(newestRental);
        if (newestUpdate === undefined) {
          this.S.alert('Lådan med id "' + id + '" har aldrig uppdaterats.');
        }
        const statusObj = {
          id: this.prettifyBoxId(id),
          Status: newestUpdate["Status"],
          Låntagare: "",
        };
        if (newestRental !== undefined && newestRental["Status"] === "uthyrning start") {
          const u = this.getUserById(Number(newestRental["Person"]));
          statusObj["Låntagare"] = `${u["Förnamn"]} ${u["Efternamn"]}, ${u["Skola"].toUpperCase()}`;
        }

        output.push(statusObj);
      });

    const html = HtmlService.createHtmlOutput(this.S.convertToHTMLTable(this.S.convertToGridWithHeaders(output)));
    html.setHeight(700).setWidth(400);

    SpreadsheetApp.getUi().showSidebar(html);
  }

  compileRentalProposal() {
    const rentalDateText = SpreadsheetApp.getUi().prompt("Vilket datum för status av uthyrningen?").getResponseText();
    const rentalDate = rentalDateText.length > 0 ? new Date(rentalDateText) : new Date();

    const termin = rentalDate.getMonth < 6 ? 'Vårtermin' : 'Hösttermin';

    const boxesByTopic = this.getBoxesByTopic();

    let rentals = [];

    this.getCurrentBookings()
    .filter(booking => booking['Termin'] === termin)
    .forEach(booking => {
      const topic = this.getTopicFromLongName(booking["Tema"]);
      let nrBoxes = Number(booking["Antal lådor"]);
      nrBoxes = isNaN(nrBoxes) ? 0 : nrBoxes;
      [...Array(nrBoxes).keys()]
        .forEach(() => {
          const availableBoxes = boxesByTopic[topic];
          if(availableBoxes.length === 0){
            this.S.alert('Inte tillräckligt med lådor för tema ' + topic);
          }
          const reservedBoxes = [boxesByTopic[topic].shift()];
          const boxNr = this.extractBoxPartsFromId(reservedBoxes[0]['id'], false)['number'];
          boxesByTopic[topic].forEach((box, index) => {
            if(this.extractBoxPartsFromId(box['id'], false)['number'] === boxNr){
              reservedBoxes.push(box);
              boxesByTopic[topic].splice(index, 1);
            }
          });
          rentals.concat(reservedBoxes.map(box => {
            return {
              'Låda': box['id'],
              'Datum': rentalDate.toISOString().substring(0,10),
              'Status': 'uthyrning start',
              'Person': this.findUser(booking['Mejl'].trim().toLowerCase())['id']
            };
          })); 
        });
    });
    this.S.alert(JSON.stringify(rentals));
  }

  /**
   *
   * @param {Array} row
   * @returns {Array}
   */
  extractMailsFromRow(row) {
    const mails = [row[0]];
    if (row[4].length > 0) {
      mails.push(...row[4].split(";"));
    }
    return mails.map(m => m.trim().toLowerCase());
  }

  /**
   * @param {String} irregularId
   *
   * @return {String} - The regular lowercase id without any special characters
   */
  standardizeBoxId(irregularId) {
    const re = /\W/g;
    return irregularId.toLowerCase().replaceAll(re, "").replaceAll("_", ""); // the lowerscore is unfortunately part of \w , so we have to remove it separately
  }

  /**
   * @param {String} regularId - a lowercase-id without any special symbols in the format aa11b (small letters followed by numbers and, optional, letters)
   *
   * @return {String} a well formatted id in the format AA:11:B
   */
  prettifyBoxId(regularId) {
    let parts = this.extractBoxPartsFromId(regularId);

    return Object.values(parts).filter(p => p !== null).join(':');
  }

  /**
   * @param {String} id - The standardized or prettified box id
   * @param {Boolean} [isRegular=true] - indicating if the id is in regular form or not (regular = 'kf01a', prettified = 'KF:01:A')
   * 
   * @returns {Object} an object with the parts topic, number and letter 
   */
  extractBoxPartsFromId(id, isRegular = true) {
    id = isRegular ? id : this.standardizeBoxId(id);
    
    let parts = Array.from(id.matchAll(/(\D+)(\d+)(\D*)/g))[0];
    parts.shift();

    parts = parts.filter(p => p.length > 0).map(p => p.toUpperCase());

    return {
      topic: parts[0],
      number: parts[1],
      letter: parts[2] ?? null
    };
  }

  /**
   * @param {string} mail - The mail adress, trimmed and lowercase
   *
   * @return {Object|undefined}
   */
  findUser(mail) {
    const foundUsers = this.getPersoner().filter(person => person["Mejl"].toLowerCase() === mail);

    if (foundUsers.length > 2) {
      this.S.alert("Flera användare hittades för " + mail);
      return;
    }
    if (foundUsers.length === 0) {
      return;
    }
    return foundUsers.shift();
  }

  /**
   * @param {Number} id - The user id
   *
   * @return {Object}
   */
  getUserById(id) {
    return this.getPersoner().find(user => user["id"] === id);
  }

  /**
   *
   * @param {Object} user
   * @param {String} topic
   * @returns {Boolean}
   */
  userHasCompetence(user, topic) {
    return (
      this.getKompetens().find(row => {
        return row["Person_id"] === user["id"] && row["Tema_id"] === topic;
      }) !== undefined
    );
  }

  /**
   *
   * @param {String} longName
   *
   * @returns {String} - the short 2-3 letter long uppercase topic abbreviation
   */
  getTopicFromLongName(longName) {
    let foundTopic = this.getTeman().find(topic => topic["Temanamn"] === longName);
    if (foundTopic === undefined) {
      this.S.alert(longName + " finns inte med som tema.");
      return;
    }
    return foundTopic["id"];
  }

  getSchoolFromLongName(longName) {
    let foundSchool = this.getSkolor().find(school => school["Alias"] === longName);

    if (foundSchool === undefined) {
      this.S.alert(longName + " finns inte med som skola.");
      return;
    }
    return foundSchool["id"];
  }

  getSchoolFromId(schoolId) {
    let foundSchool = this.getSkolor().find(school => school["id"] === schoolId);

    if (foundSchool === undefined) {
      this.S.alert(schoolId + " finns inte med som skola.");
      return;
    }
    return foundSchool["Alias"];
  }

  /**
   * @return {Array<Object>}
   */
  getPersoner() {
    if (this.personer === undefined) {
      this.personer = this.S.firstRowAsHeader(this.S.getNamedValues("Personer"));
    }
    return this.personer;
  }

  /**
   * @return {Array<Object>}
   */
  getKompetens() {
    if (this.kompetens === undefined) {
      this.kompetens = this.S.firstRowAsHeader(this.S.getNamedValues("Kompetens"));
    }
    return this.kompetens;
  }

  /**
   * @return {Array<Object>}
   */
  getTeman() {
    if (this.teman === undefined) {
      this.teman = this.S.firstRowAsHeader(this.S.getNamedValues("Teman"));
    }
    return this.teman;
  }

  /**
   * @return {Array<Object>}
   */
  getSkolor() {
    if (this.skolor === undefined) {
      this.skolor = this.S.firstRowAsHeader(this.S.getNamedValues("Skolor"));
    }
    return this.skolor;
  }

  /**
   * @return {Array<Object>}
   */
  getKompetensKontroll() {
    if (this.kompetenskontroll === undefined) {
      this.kompetenskontroll = this.S.removeIfEmpty(this.S.getNamedValues("Kompetenskontroll"));
    }
    return this.kompetenskontroll;
  }



  /**
   * @return {Array<Object>}
   */
  getBoxes() {
    if (this.boxes === undefined) {
      this.boxes = this.S.firstRowAsHeader(this.S.removeIfEmpty(this.S.getNamedValues("Lådor")));
    }
    return this.boxes;
  }

  getBoxesByTopic() {
    if (this.boxesByTopic === undefined) {
      this.boxesByTopic = this.S.groupBy(this.getBoxes(), "Tema");
    }
    return this.boxesByTopic;
  }

  /**
   * @return {Array<Object>}
   */
  getBoxUpdates() {
    if (this.boxUpdates === undefined) {
      this.boxUpdates = this.S.firstRowAsHeader(this.S.removeIfEmpty(this.S.getNamedValues("Låduppdateringar")));
    }
    return this.boxUpdates;
  }

  getCurrentBookings() {
    if (this.currentBookings === undefined) {
      this.currentBookings = this.S.firstRowAsHeader(
        this.S.removeIfEmpty(this.S.getValuesFromSheet("Aktuella_bokningar"))
      );
    }

    return this.currentBookings;
  }
}
