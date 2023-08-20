interface SimpleEvent {
  getStartTime(): GoogleAppsScript.Base.Date;
  getEndTime(): GoogleAppsScript.Base.Date;
}

function eventsOverlap(event1: SimpleEvent, event2: SimpleEvent) {
  return (
    event1.getStartTime() < event2.getEndTime() &&
    event2.getStartTime() < event1.getEndTime()
  );
}

class MinGuestsPredicate {
  constructor(
    private readonly guests: Iterable<string>,
    private readonly min: number,
  ) {}

  public test(): boolean {
    const guests = new Set();
    for (const guest of MinGuestsPredicate.maybeExpandGroups(this.guests)) {
      guests.add(guest);
      if (guests.size >= this.min) {
        return true;
      }
    }
    return false;
  }

  private static *maybeExpandGroups(
    emails: Iterable<string>,
  ): Iterable<string> {
    for (const email of emails) {
      try {
        const group = GroupsApp.getGroupByEmail(email);
        yield* group.getUsers().map((user) => user.getEmail());
        const subgroups = group
          .getGroups()
          .map((subgroup) => subgroup.getEmail());
        yield* MinGuestsPredicate.maybeExpandGroups(subgroups);
      } catch (e) {
        yield email;
      }
    }
  }
}

const LARGE_EVENT_MIN_GUESTS = 24;

function maybeSendAnEmail(event: GoogleAppsScript.Calendar.CalendarEvent) {
  if (!event.guestsCanSeeGuests()) {
    return;
  }
  const eventGuests = event.getGuestList().map((guest) => guest.getEmail());
  if (new MinGuestsPredicate(eventGuests, LARGE_EVENT_MIN_GUESTS).test()) {
    return;
  }
  const recipients = event.getCreators();
  recipients.push(Session.getEffectiveUser().getEmail());
  const eventStartTime = event.getStartTime().toLocaleString();
  const subject = `Auto-declined: ${event.getTitle()} @ ${eventStartTime}`;
  const body =
    "This event was automatically declined due to a conflict. Please let me " +
    "know if you're expecting me to attend and rescheduling is not an option. " +
    "Thanks!\n\nSent using https://github.com/avamsi/calendar-janitor.";
  GmailApp.sendEmail(recipients.join(","), subject, body);
}

const LOOK_AHEAD_DAYS = 14;
const IMPLICIT_DNS_BLOCK: { [id: string]: [number, number] } = {
  start: [0, 0], // 12:00 AM
  end: [8, 0], // 8:00 AM
};

function clean() {
  const startDate = new Date();
  const endDate = new Date(startDate);
  endDate.setDate(endDate.getDate() + LOOK_AHEAD_DAYS);

  let events = CalendarApp.getDefaultCalendar().getEvents(startDate, endDate);
  events = events.filter((event) => !event.isAllDayEvent());

  const dnsBlocks: SimpleEvent[] = events.filter((event) =>
    event.getTitle().startsWith("DNS"),
  );
  for (
    const date = new Date(startDate);
    date <= endDate;
    date.setDate(date.getDate() + 1)
  ) {
    dnsBlocks.push(
      new (class DnsBlock implements SimpleEvent {
        getStartTime() {
          const startTime = new Date(date);
          startTime.setHours(...IMPLICIT_DNS_BLOCK.start);
          return startTime;
        }
        getEndTime() {
          const endTime = new Date(date);
          endTime.setHours(...IMPLICIT_DNS_BLOCK.end);
          return endTime;
        }
      })(),
    );
  }

  for (const event of events) {
    if (event.getMyStatus() !== CalendarApp.GuestStatus.INVITED) {
      Logger.log(`Skipping ${event.getTitle()} (${event.getMyStatus()})`);
      continue;
    }
    if (dnsBlocks.some((dnsBlock) => eventsOverlap(event, dnsBlock))) {
      Logger.log(`Declining ${event.getTitle()} (DNS)`);
      event.setMyStatus(CalendarApp.GuestStatus.NO);
      maybeSendAnEmail(event);
    }
  }
}

function scheduleTriggers() {
  ScriptApp.newTrigger("clean")
    .forUserCalendar(Session.getEffectiveUser().getEmail())
    .onEventUpdated()
    .create();
}
