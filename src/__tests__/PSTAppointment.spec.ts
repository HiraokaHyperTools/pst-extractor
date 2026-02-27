import { PSTFile } from '../PSTFile.class.js';
import { PSTFolder } from '../PSTFolder.class.js';
import { PSTAppointment } from '../PSTAppointment.class.js';
import { openPstFile } from '../index.js';
import { resolve } from 'path';
let pstFile: PSTFile
let folder: PSTFolder

beforeAll(async () => {
  pstFile = await openPstFile(
    resolve('./src/__tests__/testdata/mtnman1965@outlook.com.ost')
  )

  // get to Calendar folder
  let childFolders: PSTFolder[] = await (await (await pstFile.getRootFolder()).getSubFolders())
  folder = await childFolders[1] // Root - Mailbox
  childFolders = await (await folder.getSubFolders())
  folder = await childFolders[4] // IPM_SUBTREE
  childFolders = await (await folder.getSubFolders())
  folder = await childFolders[11] // Calendar
})

afterAll(async () => {
  await pstFile?.close()
})

describe('PSTAppointment tests', () => {
  it('should have a Calendar folder', () => {
    expect(folder.displayName).toEqual('Calendar')
  })

  it('should have two calendar items', async () => {
    let appt: PSTAppointment = (await folder.getEmail(0)) as PSTAppointment
    // console.log(JSON.stringify(appt, null, 2));
    expect(appt.messageClass).toEqual('IPM.Appointment')
    expect(appt.subject).toEqual('get lunch')
    expect(appt.startTime).toEqual(new Date('2018-03-04T19:00:00.000Z'))
    expect(appt.senderName).toEqual('Mountain Man')
    expect(appt.duration).toEqual(60)

    appt = (await folder.getEmail(1)) as PSTAppointment
    // console.log(JSON.stringify(appt, null, 2));
    expect(appt.messageClass).toEqual('IPM.Appointment')
    expect(appt.subject).toEqual('workout')
    expect(appt.creationTime).toEqual(new Date('2018-03-05T20:26:26.396Z'))
    expect(appt.senderName).toEqual('Mountain Man')
    expect(appt.duration).toEqual(60)
    expect(appt.transportMessageHeaders).toEqual('')
    expect(appt.location).toEqual('')
    expect(appt.allAttendees).toEqual('')
    expect(appt.toAttendees).toEqual('')
    expect(appt.ccAttendees).toEqual('')
    expect(appt.netMeetingServer).toEqual('')
    expect(appt.netMeetingOrganizerAlias).toEqual('')
    expect(appt.netMeetingDocumentPathName).toEqual('')
    expect(appt.netShowURL).toEqual('')
    expect(appt.requiredAttendees).toEqual('')
    expect(appt.recurrencePattern).toEqual(
      'every day from 10:00 AM to 11:00 AM'
    )
    expect(appt.importance).toEqual(2)
    expect(appt.busyStatus).toEqual(2)
    expect(appt.color).toEqual(0)
    expect(appt.meetingStatus).toEqual(0)
    expect(appt.responseStatus).toEqual(0)
    expect(appt.recurrenceType).toEqual(1)
    expect(appt.appointmentSequence).toEqual(0)
    expect(appt.netMeetingType).toEqual(0)
    expect(appt.localeId).toEqual(1033)
    expect(appt.startTime).toEqual(new Date('2018-03-05T17:00:00.000Z'))
    expect(appt.endTime).toEqual(new Date('2018-03-05T18:00:00.000Z'))
    expect(appt.showAsBusy).toBeTruthy()
    expect(appt.isRecurring).toBeTruthy()
    expect(appt.sendAsICAL).toBeFalsy()
    expect(appt.subType).toBeFalsy()
    expect(appt.isOnlineMeeting).toBeFalsy()
    expect(appt.appointmentCounterProposal).toBeFalsy()
    expect(appt.isSilent).toBeFalsy()
    expect(appt.recurrenceBase).toEqual(null)
    expect(appt.attendeeCriticalChange).toEqual(null)
  })
})
