/**
 * Webb Alum Kings ballot -> Google Sheet
 *
 * Setup (~5 min):
 * 1. New Google Sheet. Row 1 headers: Timestamp | Voter ID | Name | Entry Vote | Additional Vote | Additional At
 * 2. Extensions -> Apps Script. Paste this file. Set SECRET below to match config.js on the site.
 * 3. Deploy -> New deployment -> Web app -> Execute as: Me, Who has access: Anyone
 * 4. Copy the web app URL into config.js as endpoint. Push config.js to GitHub (or edit on GitHub).
 * 5. Open results.html on your site (or the Sheet) to see votes.
 */

var SECRET = 'wak-ballot-2026'

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents)
    if (body.secret !== SECRET) {
      return jsonOut({ ok: false, error: 'unauthorized' })
    }

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]
    ensureHeaders(sheet)

    var voterId = String(body.voterId || '')
    if (!voterId) {
      return jsonOut({ ok: false, error: 'missing voterId' })
    }

    var name = String(body.name || '')
    var row = findVoterRow(sheet, voterId)

    if (body.type === 'entry') {
      if (row.index > 0) {
        return jsonOut({ ok: true, duplicate: true })
      }
      sheet.appendRow([
        new Date(),
        voterId,
        name,
        'yes',
        '',
        '',
      ])
      return jsonOut({ ok: true })
    }

    if (body.type === 'additional') {
      var vote = body.vote === 'no' ? 'no' : 'yes'
      if (row.index > 0) {
        var existingAdditional = sheet.getRange(row.index, 5).getValue()
        if (existingAdditional) {
          return jsonOut({ ok: true, duplicate: true })
        }
        sheet.getRange(row.index, 5).setValue(vote)
        sheet.getRange(row.index, 6).setValue(new Date())
        if (name && !sheet.getRange(row.index, 3).getValue()) {
          sheet.getRange(row.index, 3).setValue(name)
        }
        return jsonOut({ ok: true })
      }
      sheet.appendRow([new Date(), voterId, name, 'yes', vote, new Date()])
      return jsonOut({ ok: true })
    }

    return jsonOut({ ok: false, error: 'unknown type' })
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) })
  }
}

function doGet(e) {
  var params = e && e.parameter ? e.parameter : {}
  if (params.secret !== SECRET) {
    return jsonOut({ ok: false, error: 'unauthorized' })
  }

  if (params.action === 'stats') {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]
    ensureHeaders(sheet)
    var values = sheet.getDataRange().getValues()
    var ballots = 0
    var entryYes = 0
    var additionalYes = 0
    var additionalNo = 0
    var pendingAdditional = 0

    for (var i = 1; i < values.length; i++) {
      ballots++
      entryYes++
      var add = values[i][4]
      if (add === 'yes') additionalYes++
      else if (add === 'no') additionalNo++
      else pendingAdditional++
    }

    return jsonOut({
      ok: true,
      ballots: ballots,
      entryYes: entryYes,
      additionalYes: additionalYes,
      additionalNo: additionalNo,
      pendingAdditional: pendingAdditional,
      updatedAt: new Date().toISOString(),
    })
  }

  return jsonOut({ ok: true, message: 'Webb Alum Kings vote tracker' })
}

function ensureHeaders(sheet) {
  if (sheet.getLastRow() > 0) return
  sheet.appendRow([
    'Timestamp',
    'Voter ID',
    'Name',
    'Entry Vote',
    'Additional Vote',
    'Additional At',
  ])
}

function findVoterRow(sheet, voterId) {
  var values = sheet.getDataRange().getValues()
  for (var i = 1; i < values.length; i++) {
    if (values[i][1] === voterId) {
      return { index: i + 1 }
    }
  }
  return { index: 0 }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON
  )
}
