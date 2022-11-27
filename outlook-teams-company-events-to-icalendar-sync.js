

/**
 * Export your Outlook Calendar events as iCalendar *.ics file to sync with your work-life-balance private calendar.
 *
 * Bookmarklet that generates an *.ics file of your corporate events.
 *
 * Set your `userPrivateTld` to initiate an e-mail message to your address.
 */


// Robot settings
let dt = new Date(),
    userPrivateTld = 'gmail.com', // redirect to Office365 Outlook `/mail` App and create message to `<user.name>@<private-mailbox.tld>`
    userEmail = '', // company email address: `<user.name>@<corporate-mailbox.tld>`
    companyName = '', // company name @ Office365
    companyTld = '', // company email address top level domain
    dtstamp = dt.getFullYear() + (((dt.getMonth()+1) < 10) ? '0' : '') + (dt.getMonth()+1) + ((dt.getDate() < 10) ? '0' : '') + dt.getDate() + 'T' + ((dt.getHours() < 10)? '0' : '') + dt.getHours() + ((dt.getMinutes() < 10) ? '0' : '') + dt.getMinutes() + ((dt.getSeconds() < 10) ? '0' : '') + dt.getSeconds() + 'Z'
    dateToday = dt.getFullYear() + (((dt.getMonth()+1) < 10) ? '0' : '') + '-' + (dt.getMonth()+1) + '-' + ((dt.getDate() < 10) ? '0' : '') + dt.getDate(),
    dateNextYear = (dt.getFullYear() + 1) + (((dt.getMonth()+1) < 10) ? '0' : '') + '-' + (dt.getMonth()+1) + '-' + ((dt.getDate() < 10) ? '0' : '') + dt.getDate(),
    timeNow = ((dt.getHours() < 10)? '0' : '') + dt.getHours() + '-' + ((dt.getMinutes() < 10) ? '0' : '') + dt.getMinutes() + '-' + ((dt.getSeconds() < 10) ? '0' : '') + dt.getSeconds(), // + 'Z',
    reqRangeStart = dateToday, // get meetings from today ...
    reqRangeEnd = dateNextYear, // ... until next year
    reqCanaryId = null, // OWA authentication Id
    calTimeZone = 'Europe/Vienna', // 2nd best city next to #StP ;)
    appTld = 'outlook.office.com', // Office365 App Robot is enabled
    appTldSuffix = '/calendar/view/month/' // Open Calendar App or: Mail App `'/mail/'`


// The robot can only operate on the `appTld`
// URL: https://outlook.office.com/calendar/view/month
let currentLocation = document.location.href,
    appTldPrefix = 'https://'

if (!currentLocation.includes(appTld)) {
    console.log('*** Robot: I can only operate on the TLD ' + appTld)

    // redirect to app where robot can operon on:
    window.location.href = appTldPrefix + appTld + appTldSuffix
}

// Authenticate API requests
const cookieName  = 'X-OWA-CANARY', // OWA API Authentication Token
    cookieValue = `; ${document.cookie}`,
    cookieParts = cookieValue.split(`; ${cookieName}=`)

if (cookieParts.length === 2) {
    reqCanaryId = cookieParts.pop().split(';').shift()
} else {
    console.log('*** Problem detected: can not find OWA Canary Id.')

    // a problem occured - redirect to a help page
    // window.location.href = 'https://support.microsoft.com/xxx'
}

async function doDataRequest(reqCanaryId) {
    let url = 'https://outlook.office.com/owa/startupdata.ashx?app=Calendar&n=0',
        res = await fetch(url, {
            method: 'POST',
            headers: {
                "x-owa-canary": reqCanaryId,
                "action": "StartupData",
                "x-req-source": "Calendar",
            },
        })

    if (res.ok) {
        let ret = await res.json(),
            jsonDataPath = ret.findFolders.Body.ResponseMessages.Items[0].RootFolder.Folders

        // configuration
        userEmail = ret.owaUserConfig.SessionSettings.UserEmailAddress
        companyTld = ret.owaUserConfig.SessionSettings.OrganizationDomain
        companyName = ret.owaUserConfig.SessionSettings.CompanyName
        calTimeZone = ret.owaUserConfig.UserOptions.MailboxTimeZoneOffset[0]['IanaTimeZones'][0]
        calTimeZoneMS = ret.owaUserConfig.UserOptions.TimeZone

        //---debug---
        // console.log('calTimeZone', JSON.stringify(calTimeZone))
        // console.log(ret.SessionSettings.UserEmailAddress)
        // console.log(JSON.stringify(ret.findFolders.Body.ResponseMessages.Items[0].RootFolder.Folders))

        return (ret.findFolders && ret.findFolders.Body) ? jsonDataPath : {'success': false, 'msg': 'No Folder found.'}
    } else {
        return `HTTP error: ${res.status}`
    }
}

async function doCalendarRequest(reqFolderId, reqCanaryId) {

    let url = 'https://outlook.office.com/owa/service.svc?action=GetCalendarView&app=Calendar&n=0',
        res = await fetch(url, {
            method: 'POST',
            headers: {
                "x-owa-canary": reqCanaryId,
                "action": "GetCalendarView",
                "x-req-source": "Calendar",
                "x-owa-urlpostdata": "%7B%22__type%22%3A%22GetCalendarViewJsonRequest%3A%23Exchange%22%2C%22Header%22%3A%7B%22__type%22%3A%22JsonRequestHeaders%3A%23Exchange%22%2C%22RequestServerVersion%22%3A%22V2018_01_08%22%2C%22TimeZoneContext%22%3A%7B%22__type%22%3A%22TimeZoneContext%3A%23Exchange%22%2C%22TimeZoneDefinition%22%3A%7B%22__type%22%3A%22TimeZoneDefinitionType%3A%23Exchange%22%2C%22Id%22%3A%22" + encodeURIComponent(calTimeZoneMS) + "%22%7D%7D%7D%2C%22Body%22%3A%7B%22__type%22%3A%22GetCalendarViewRequest%3A%23Exchange%22%2C%22CalendarId%22%3A%7B%22__type%22%3A%22TargetFolderId%3A%23Exchange%22%2C%22BaseFolderId%22%3A%7B%22__type%22%3A%22FolderId%3A%23Exchange%22%2C%22Id%22%3A%22" + encodeURIComponent(reqFolderId) + "%22%7D%7D%2C%22RangeStart%22%3A%22" + reqRangeStart + "T00%3A00%3A00.000%22%2C%22RangeEnd%22%3A%22" + reqRangeEnd + "T00%3A00%3A00.000%22%7D%7D",
            },
            //---debug---
            // orig -- x-owa-urlpostdata: {"__type":"GetCalendarViewJsonRequest:#Exchange","Header":{"__type":"JsonRequestHeaders:#Exchange","RequestServerVersion":"V2018_01_08","TimeZoneContext":{"__type":"TimeZoneContext:#Exchange","TimeZoneDefinition":{"__type":"TimeZoneDefinitionType:#Exchange","Id":"W. Europe Standard Time"}}},"Body":{"__type":"GetCalendarViewRequest:#Exchange","CalendarId":{"__type":"TargetFolderId:#Exchange","BaseFolderId":{"__type":"FolderId:#Exchange","Id":"AQMkADYyYmNkNGMxLWEwY2MALTQ2OGMtOGYyNy1kNWM2MjE1OTJiMmQALgAAA0AwE5rbwQBOstHy6V1GgbABAKR5Uyyd0yRMtXgZ7NZuTCwAAAIBDQAAAA=="}},"RangeStart":"2022-11-28T00:00:00.000","RangeEnd":"2023-01-02T00:00:00.000"}}
        })

    if (res.ok) {
        let ret = await res.json(),
        jsonDataPath = ret.Body.Items

        return (ret.Body && ret.Body.Items) ? jsonDataPath : {'success': false, 'msg': 'No Events found.'}
    } else {
        return `HTTP error: ${res.status}`
    }
}

// date format helper
function fmtDate(date) {
    return date.replace('+01:00', '').replaceAll('-', '').replaceAll(':', '')
}

doDataRequest(reqCanaryId).then(data => {
    let folderCounter = 0, // should work ...
        folderName = ['Kalender', 'Calendar'] // not that good ...

    data.forEach(folder => {
        if (folder['__type'] == 'CalendarFolder:#Exchange' && folderName.includes(folder['DisplayName'])) {
            reqFolderId = folder.FolderId.Id
            return
        }
        /*if (folder['__type'] === 'CalendarFolder:#Exchange') {
            folderCounter++
            if (folderCounter === 1) {
                reqFolderId = folder.FolderId.Id
                return
            }
        }*/
    });

    return reqFolderId
}).then(function(reqFolderId) {
    doCalendarRequest(reqFolderId,   reqCanaryId).then(data => {
        let calendar = '',
            iCal = 'BEGIN:VCALENDAR\r\nVERSION:2.0\r\n',
            iCalEnd = 'END:VCALENDAR\r\nUID:uid+' + userEmail + '\r\nDTSTAMP:' + dtstamp + '\r\nPRODID:calendar+' + userEmail + '\r\n'

        if (data && typeof data == 'object') {
            data.forEach(event => {
                // event header
                iCal += 'BEGIN:VEVENT\r\nTRANSP:TRANSPARENT\r\nLOCATION:Meeting\r\nCLASS:PUBLIC\r\nDTSTAMP:' + dtstamp + ''

                // event details
                iCal += '\r\nSUMMARY:' + event.Subject
                iCal += '\r\nDESCRIPTION: @' + event.Organizer.Mailbox.Name
                // iCal += '\r\nURL:' + appTldPrefix + appTld + appTldSuffix

                // event metadata
                iCal += '\r\nDTSTART;TZID=Europe/Vienna:' + fmtDate(event.Start) + '\r\nDTEND;TZID=Europe/Vienna:' + fmtDate(event.End)
                iCal += '\r\nDTSTAMP:' + dtstamp + '\r\nUID:' + event.ItemId.Id + '\r\nEND:VEVENT\r\n' // TODO event.UID vs.: event.ItemId.Id
            });

            // build calendar @ *.ics file format
            calendar = iCal + iCalEnd

            //---debug---
            // console.log(calendar);
            // dev tools: copy(calendar);


            // download final iCalendar document
            var blob = new Blob([calendar], {type: 'text/calendar'}),
                a = document.createElement('a')

            var e = new MouseEvent('click', {
                view: window,
                bubbles: true,
                cancelable: false
            })

            a.download = 'company-cal-events.' + companyTld + '--' + dtstamp + '.ics'
            a.href = window.URL.createObjectURL(blob)
            a.dataset.downloadurl =  ['text/calendar', a.download, a.href].join(':')
            a.dispatchEvent(e)
            //-- end of iCalendar file download

            //-- optional
            // redirect to compose a new e-mail message when private tld is set.
            if (userPrivateTld != '') {
                let url = 'https://outlook.office.com/mail/deeplink/compose'

                url += '?to=' + userEmail.replace(companyTld, userPrivateTld)
                url += '&subject=iCal+Sync&body=iCal+' + companyName + ' - ' + dtstamp

                window.location.href = url
            }
        } else {
            // there was a problem : /
            console.log(data)
        }
    })
})
