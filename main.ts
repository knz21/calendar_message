const slackToken = PropertiesService.getScriptProperties().getProperty('slack_bot_token')
const chennelId = PropertiesService.getScriptProperties().getProperty('channel_id')
const calendarId = PropertiesService.getScriptProperties().getProperty('calendar_id')
const spreadSheetId = PropertiesService.getScriptProperties().getProperty('spreadsheet_id')
const spreadSheetName = 'messages'
const columnIndex = {
    id: 0,
    title: 1,
    message: 2
}

const main = () => {
    createTrigger()
}

const getEventsForNextTwoHours = () => {
    if (!calendarId) return []
    const calendar = CalendarApp.getCalendarById(calendarId)
    const now = new Date()
    const twoHoursLater = new Date(now.getTime() + 2 * 60 * 60 * 1000)
    const events = calendar.getEvents(now, twoHoursLater)
    return events
}

const createTrigger = () => {
    if (!spreadSheetId) return
    const sheet = SpreadsheetApp.openById(spreadSheetId).getSheetByName(spreadSheetName)
    if (sheet == null) return
    const events = getEventsForNextTwoHours()
    if (events.length == 0) return
    const lastRow = sheet.getLastRow()
    let filteredMessages: string[][] = []
    if (lastRow > 0) {
        const triggeredMessages = sheet.getRange(1, 1, sheet.getLastRow(), Object.keys(columnIndex).length).getValues()
        filteredMessages = triggeredMessages.filter(row => {
            const title = row[columnIndex.title]
            return !events.some(event => event.getTitle() == title)
        })
    }
    const newMessages: string[][] = []
    events.forEach(event => {
        const trigger = ScriptApp.newTrigger('sendMessage')
            .timeBased()
            .after(1 * 60 * 1000)
            .create()
        const messageRow = []
        messageRow[columnIndex.id] = trigger.getUniqueId()
        messageRow[columnIndex.title] = event.getTitle()
        messageRow[columnIndex.message] = event.getDescription()
        newMessages.push(messageRow)
    })
    const newTrigerredMessages = filteredMessages.concat(newMessages)
    sheet.clearContents()
    sheet.getRange(1, 1, newTrigerredMessages.length, Object.keys(columnIndex).length).setValues(newTrigerredMessages)
}

const sendMessage = (e: any) => {
    if (!spreadSheetId || !e.triggerUid) return
    const sheet = SpreadsheetApp.openById(spreadSheetId).getSheetByName(spreadSheetName)
    if (!sheet) return
    const triggeredMessages = sheet.getRange(1, 1, sheet.getLastRow(), Object.keys(columnIndex).length).getValues()
    const messageRowIndex = triggeredMessages.findIndex(row => row[columnIndex.id] == e.triggerUid)
    if (messageRowIndex == -1) return
    const messageRow = triggeredMessages[messageRowIndex]
    if (!messageRow) return
    const message = messageRow[columnIndex.message]
    postMessageToSlack(message)
    sheet.deleteRow(messageRowIndex + 1)
}

const postMessageToSlack = (
    text: string,
    threadTs: string | null = null,
    channel: string | null = chennelId
) => {
    const formData = {
        token: slackToken,
        channel: channel,
        text: text,
        thread_ts: threadTs
    }
    const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
        method: 'post',
        payload: formData,
        muteHttpExceptions: true
    }
    UrlFetchApp.fetch('https://slack.com/api/chat.postMessage', options)
}