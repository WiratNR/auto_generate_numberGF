const Name = {
  HEADER: "รหัสสินค้า",
  PREFIX: "SKU",
  LENGTH: 5,
}

class App {
  constructor() {
    this.ss = SpreadsheetApp.getActive()
    this.sheet = this.getLinkedSheet()
    if (!this.sheet) {
      throw Error(`ไม่มีลิ้งไปยัง Google Sheet`)
    }
    this.form = FormApp.openByUrl(this.sheet.getFormUrl())
    this.message = this.form.getConfirmationMessage()
    this.uidRegex = new RegExp(`${Name.PREFIX}\\d{${Name.LENGTH}}`, 'gi');

  }
  createUidBynumber(number) {
    return Name.PREFIX + (10 ** Name.LENGTH + number).toString().slice(-Name.LENGTH)
  }

  getLinkedSheet() {
    return this.ss.getSheets().find(sheet => sheet.getFormUrl())
  }

  getUidFromConfirmationMessage() {
    const message = this.form.getConfirmationMessage()
    const result = message.match(this.uidRegex)
    if (!result)
      throw Error(`no Name ${this.uidRegex}`)
    return result[0]
  }

  createNextUid(curentUid) {
    const nextUidNumber = Number(curentUid.replace(Name.PREFIX, "")) + 1
    return this.createUidBynumber(nextUidNumber)
  }

  saveCurrentUid(uid, rowStart) {
    const [headers] = this.sheet.getDataRange().getDisplayValues()
    let uidHeadIndex = headers.indexOf(Name.HEADER)
    if (uidHeadIndex === -1) {
      uidHeadIndex = headers.length
      this.sheet.getRange(1, uidHeadIndex + 1).setValue(Name.HEADER)
    }
    this.sheet.getRange(rowStart, uidHeadIndex + 1).setValue(uid)
  }

  updateConfirmationMessage(nextUid) {
    const message = this.message.replace(this.uidRegex, nextUid)
    this.form.setConfirmationMessage(message)
  }

  run(e) {
    const { rowStart } = e.range
    const curentUid = this.getUidFromConfirmationMessage()
    this.saveCurrentUid(curentUid, rowStart)
    const nextUid = this.createNextUid(curentUid)
    this.updateConfirmationMessage(nextUid)
  }
}


function _onFormSubmit(e) {
  new App().run(e)
}
