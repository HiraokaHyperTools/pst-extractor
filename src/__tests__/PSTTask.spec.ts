import { openPstFile } from '../openPstFile.js';
import { PSTFile } from '../PSTFile.class.js';
import { PSTFolder } from '../PSTFolder.class.js';
import { PSTTask } from '../PSTTask.class.js';
import { resolve } from 'path';
let pstFile: PSTFile
let folder: PSTFolder

beforeAll(async () => {
  pstFile = await openPstFile(
    resolve('./src/__tests__/testdata/mtnman1965@outlook.com.ost')
  )

  // get to Tasks folder
  let childFolders: PSTFolder[] = (await (await pstFile.getRootFolder()).getSubFolders())
  folder = childFolders[1] // Root - Mailbox
  childFolders = (await folder.getSubFolders())
  folder = childFolders[4] // IPM_SUBTREE
  childFolders = (await folder.getSubFolders())
  folder = childFolders[17] // Tasks
})

afterAll(async () => {
  await pstFile.close()
})

describe('PSTTask tests', () => {
  it('should have a Tasks folder', () => {
    expect(folder.displayName).toEqual('Tasks')
  })

  it('should have two tasks', async () => {
    // fully loaded task
    let task: PSTTask = (await folder.getEmail(0)) as PSTTask
    // Log.debug1(JSON.stringify(task, null, 2));
    expect(task.messageClass).toEqual('IPM.Task')
    expect(task.subject).toEqual('fully loaded task')
    expect(task.isTaskRecurring).toBeTruthy()
    expect(task.isTaskComplete).toBeFalsy()
    expect(task.taskOwner).toEqual('Mountain Man')
    expect(task.taskStatus).toEqual(1) // started
    expect(task.percentComplete).toEqual(0.75)
    expect(task.bodyPrefix).toContain(
      'Blue category, high priority, 75% complete'
    )

    // basic task
    task = (await folder.getEmail(1)) as PSTTask
    // Log.debug1(JSON.stringify(task, null, 2));
    expect(task.messageClass).toEqual('IPM.Task')
    expect(task.subject).toEqual('basic task')
    expect(task.isTaskRecurring).toBeFalsy()
    expect(task.isTaskComplete).toBeFalsy()
    expect(task.taskOwner).toEqual('Mountain Man')
    expect(task.taskStatus).toEqual(0) // not started
    expect(task.bodyPrefix).toContain('Vanilla task, not started')
    expect(task.importance).toEqual(1)
    expect(task.percentComplete).toEqual(0)
    expect(task.taskActualEffort).toEqual(0)
    expect(task.taskEstimatedEffort).toEqual(0)
    expect(task.taskVersion).toEqual(5)
    expect(task.taskOrdinal).toEqual(-4000)
    expect(task.taskOwnership).toEqual(0)
    expect(task.acceptanceState).toEqual(0)
    expect(task.transportMessageHeaders).toEqual('')
    expect(task.taskAssigner).toEqual('')
    expect(task.taskLastUser).toEqual('')
    expect(task.isTaskComplete).toBeFalsy()
    expect(task.isTaskRecurring).toBeFalsy()
    expect(task.taskDateCompleted).toEqual(null)
  })
})
