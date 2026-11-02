import { calendar_v3 } from '@googleapis/calendar'
import { sync } from '../src/subscriptions'
import { fetchCalendarEvents } from '../src/utils'

jest.mock('../src/utils', () => ({
  ...jest.requireActual('../src/utils'),
  fetchCalendarEvents: jest.fn(),
}))

const events: calendar_v3.Schema$Event[] = [
  { id: 'a', summary: 'A', start: { date: '2026-01-01' }, end: { date: '2026-01-02' } },
  { id: 'b', summary: 'B', start: { date: '2026-01-03' }, end: { date: '2026-01-04' } },
]

const subscription = { calendarId: 'calendar', summary: 'Test', url: 'https://example.com/calendar.ics' }

const createCalendar = (insert: jest.Mock): calendar_v3.Calendar =>
  ({
    events: {
      list: jest.fn().mockResolvedValue({ data: { items: [] } }),
      insert,
      update: jest.fn(),
      delete: jest.fn(),
    },
  }) as unknown as calendar_v3.Calendar

const delayedInsert = (inserted: string[]): jest.Mock =>
  jest.fn(async ({ requestBody }) => {
    await new Promise(resolve => setTimeout(resolve, 10))
    inserted.push(requestBody.id)
  })

beforeEach(() => {
  jest.mocked(fetchCalendarEvents).mockResolvedValue(events)
  jest.spyOn(console, 'warn').mockImplementation(() => undefined)
})

afterEach(() => jest.restoreAllMocks())

describe('sync', () => {
  it('resolves after a subscription finished syncing', async () => {
    const inserted: string[] = []
    await sync(createCalendar(delayedInsert(inserted)), { requestBody: subscription })
    expect(inserted).toEqual(['a', 'b'])
  })

  it('resolves after multiple subscriptions finished syncing', async () => {
    const inserted: string[] = []
    const requestBody = [subscription, subscription] as unknown as calendar_v3.Schema$Subscription
    await sync(createCalendar(delayedInsert(inserted)), { requestBody })
    expect(inserted).toEqual(['a', 'b', 'a', 'b'])
  })

  it('rejects when an event fails to sync, after syncing the other events', async () => {
    const insert = jest.fn(async ({ requestBody }) => {
      if (requestBody.id === 'a') throw new Error('quota exceeded')
    })
    const calendar = createCalendar(insert)

    await expect(sync(calendar, { requestBody: subscription })).rejects.toThrow(
      'Failed to sync 1 event of calendar Test',
    )
    expect(insert).toHaveBeenCalledTimes(2)
  })

  it('syncs every subscription before rejecting for a failed one', async () => {
    const insert = jest.fn(async ({ calendarId }) => {
      if (calendarId === 'first') throw new Error('not found')
    })
    const requestBody = [
      { ...subscription, calendarId: 'first', summary: 'First' },
      { ...subscription, calendarId: 'second', summary: 'Second' },
    ] as unknown as calendar_v3.Schema$Subscription

    await expect(sync(createCalendar(insert), { requestBody })).rejects.toThrow(
      'Failed to sync 2 events of calendar First',
    )
    expect(insert.mock.calls.filter(([{ calendarId }]) => calendarId === 'second')).toHaveLength(2)
  })
})
