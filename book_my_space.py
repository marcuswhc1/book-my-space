"""Auto-book the flexible desk on BookMySpace for next week (Tue-Thu). Schedule it to run every Monday.

Run with:  python book_my_space.py

The flow (see run() at the bottom):
  1. Work out next week's Tue / Wed / Thu dates from today's date.
  2. Open the leave portal and find which of those dates are Approved Leave / Public Holiday / Non-workday.
     Those dates are NOT booked and are emailed to you.
  3. Open BookMySpace and book the remaining dates.
  4. Print a summary and email any booking failures.
"""
import datetime as dt
import os
import re
import time
from urllib.parse import urlparse

import win32com.client as win32
from dotenv import load_dotenv
from playwright.sync_api import Page, Playwright, TimeoutError as PWTimeout, sync_playwright

################
# Configurations
################
# Secrets and URLs live in the .env file (never commit it). load_dotenv() reads them into the environment.
load_dotenv()

BASE_URL = os.getenv("BOOKMYSPACE_BASE_URL")            # BookMySpace website
LEAVE_PORTAL_URL = os.getenv("LEAVE_PORTAL_URL")        # annual leave portal (no sign-in needed)
EMAIL = os.getenv("BOOKMYSPACE_EMAIL")                  # used on the Microsoft sign-in page
NOTIFY_EMAIL = os.getenv("BOOKMYSPACE_NOTIFY_EMAIL")    # who gets the emails (sent through the desktop Outlook)
SPACE = os.getenv("BOOKMYSPACE_SPACE")                  # the desk to book, e.g. 7E.FlexDesk.BB-016
BROWSER = os.getenv("BOOKMYSPACE_BROWSER", "chrome")
LOGIN_TIMEOUT_SEC = int(os.getenv("BOOKMYSPACE_LOGIN_TIMEOUT_SEC"))
# Days to book: Mon=0 ... Tue=1, Wed=2, Thu=3 (comma separated in .env, e.g. "1,2,3").
BOOK_WEEKDAYS = tuple(int(d.strip()) for d in os.getenv("BOOKMYSPACE_WEEKDAYS").split(","))

SEND_EMAILS = True      # False: only print, do not send any email. True: actually send emails.
CONFIRM_BOOKING = True  # False: click Cancel in the confirmation popup (test run). True: really book.
HEADLESS = True        # True hides the browser window, False shows it.

# How the leave calendar marks each type of day (the class name on the day's cell).
LEAVE_TYPES = {
    "sapMeCalendarType01": "Approved Leave",
    "sapMeCalendarType00": "Non-workday",
    "sapMeCalendarType06": "Public Holiday",
}
LEAVE_PORTAL_WAIT_SEC = int(os.getenv("LEAVE_PORTAL_WAIT_SEC"))  # the leave portal is slow to load
CALENDAR_NEXT_SELECTOR = ".sapMeCalendarNext"  # the ">" arrow at the top of the leave calendar
MAX_CALENDAR_PAGE_CLICKS = 2                   # how many times we may click ">" to reach a later month


# Raised when a period is deliberately not booked; the message is the reason shown in the summary.
class NotBooked(Exception):
    pass


##################
# Helper functions
##################
# Print a message with a timestamp.
def log(msg: str) -> None:
    print(f"[{dt.datetime.now():%H:%M:%S}] {msg}")


# Stop early with a clear message if something is missing in the .env file.
def check_config() -> None:
    required = {
        "BOOKMYSPACE_BASE_URL": BASE_URL,
        "LEAVE_PORTAL_URL": LEAVE_PORTAL_URL,
        "BOOKMYSPACE_EMAIL": EMAIL,
        "BOOKMYSPACE_NOTIFY_EMAIL": NOTIFY_EMAIL,
        "BOOKMYSPACE_SPACE": SPACE,
    }
    missing = [name for name, value in required.items() if not value]
    if missing:
        raise SystemExit(
            f"Missing settings: {', '.join(missing)}. "
            "Check that .env exists next to this script and that every line is KEY=value."
        )


# Text like "2026-10-20 (Tue)".
def day_label(day: dt.date) -> str:
    return f"{day} ({day:%a})"


# Text like "2026-10-20 (Tue) to 2026-10-22 (Thu)".
def period_label(start: dt.date, end: dt.date) -> str:
    return f"{day_label(start)} to {day_label(end)}"


# First line of an error message (keeps the summary short).
def first_line(e: Exception) -> str:
    return (str(e).strip().splitlines() or [type(e).__name__])[0]


# Send an email through the desktop Outlook.
def send_email(subject: str, lines: list[str]) -> None:
    if not SEND_EMAILS:
        return
    try:
        mail = win32.Dispatch("Outlook.Application").CreateItem(0)
        mail.To = NOTIFY_EMAIL
        mail.Subject = subject
        mail.Body = "\r\n".join(lines)
        mail.Send()
        log(f"Email sent to {NOTIFY_EMAIL}: {subject}")
    except Exception as e:
        log(f"Could not send email '{subject}': {e}")


############################################
# STEP 1: work out the dates we want to book
############################################
# Return next week's Tue / Wed / Thu (or whichever weekdays are set in BOOK_WEEKDAYS).
def get_target_dates(today: dt.date) -> list[dt.date]:
    this_monday = today - dt.timedelta(days=today.weekday())
    next_monday = this_monday + dt.timedelta(days=7)
    return [next_monday + dt.timedelta(days=d) for d in BOOK_WEEKDAYS]


# Group dates into runs of consecutive days, e.g. [Tue, Thu] -> [(Tue, Tue), (Thu, Thu)].
# BookMySpace books one date range at a time, so each run is one booking.
def group_consecutive(days: list[dt.date]) -> list[tuple[dt.date, dt.date]]:
    runs: list[list[dt.date]] = []
    for day in sorted(days):
        if runs and day - runs[-1][1] == dt.timedelta(days=1):
            runs[-1][1] = day  # extends the current run
        else:
            runs.append([day, day])  # starts a new run
    return [(first, last) for first, last in runs]


################################
# STEP 2: check the leave portal
################################
# CSS selector for one day's cell. Cell ids look like "main--S4--LRS4_DAT_CALENDAR-Tue-Oct-13-2026".
# The greyed-out days of the neighbouring month (e.g. Oct 1 shown in the Sep grid) have the same id,
# so we skip cells marked sapMeCalendarDayNotInCurrentMonth.
def leave_day_selector(day: dt.date) -> str:
    return f"[id$='-{day:%a-%b-%d-%Y}']:not(.sapMeCalendarDayNotInCurrentMonth)"


# Find a day's cell. The calendar shows two months; if the day is not shown yet, click ">" to go forward.
def find_leave_day_cell(page: Page, day: dt.date):
    cell = page.locator(leave_day_selector(day)).first
    for _ in range(MAX_CALENDAR_PAGE_CLICKS + 1):
        try:
            cell.wait_for(state="attached", timeout=3000)
            return cell
        except PWTimeout:
            log(f"  {day} is not in view, clicking the calendar's next-page arrow")
            page.locator(CALENDAR_NEXT_SELECTOR).first.click()
    raise RuntimeError(f"{day} was not found in the leave calendar")


# Open the leave portal and return {date: reason} for every date that is leave / holiday / non-workday.
def check_leave_portal(page: Page, days: list[dt.date]) -> dict[dt.date, str]:
    log("Opening the leave portal")
    page.goto(LEAVE_PORTAL_URL, wait_until="load", timeout=90000)

    # The portal may open on the tile page (click "Request Leave") or go straight to the calendar.
    request_leave = page.get_by_text("Request Leave", exact=True).first
    try:
        request_leave.wait_for(state="visible", timeout=LEAVE_PORTAL_WAIT_SEC * 1000)
        request_leave.click()
        log("- clicked 'Request Leave'")
    except PWTimeout:
        log("- 'Request Leave' tile not shown, assuming the calendar page is already open")

    try:
        page.locator("[class*='sapMeCalendar']")        .first.wait_for(state="visible", timeout=LEAVE_PORTAL_WAIT_SEC * 1000)
    except PWTimeout:
        raise RuntimeError(f"The leave calendar did not appear (at {page.url})")

    # Read the class name on each date's cell and match it against LEAVE_TYPES.
    blocked: dict[dt.date, str] = {}
    for day in days:
        classes = find_leave_day_cell(page, day).get_attribute("class") or ""
        for class_name, reason in LEAVE_TYPES.items():
            if class_name in classes:
                blocked[day] = reason
                break
        log(f"- {day_label(day)}: {blocked.get(day, 'free')}")
    return blocked


# Email the days that cannot be booked because of leave / holiday / non-workday.
def send_leave_email(blocked: dict[dt.date, str], bookable: list[dt.date]) -> None:
    lines = [f"BookMySpace run on {dt.datetime.now():%Y-%m-%d %H:%M}", "", "NOT booked (leave portal):"]
    lines += [f"  - {day_label(day)}: {reason}" for day, reason in sorted(blocked.items())]
    lines += ["", "Will still be booked:" + ("" if bookable else " none")]
    lines += [f"  - {day_label(day)}" for day in bookable]
    send_email(f"BookMySpace: {len(blocked)} day(s) not booked (leave/holiday)", lines)


#############################
# STEP 3: book on BookMySpace
#############################
# --- Sign in ---

# True only when the browser is on the BookMySpace site (the login page URL also contains its name, so check the host).
def on_app(page: Page) -> bool:
    return urlparse(page.url).hostname == "bookmyspace.infineon.com"


# Sign in: enter the email, click Next, then Yes on 'Stay signed in?'.
def sign_in(page: Page) -> None:
    page.goto(BASE_URL, wait_until="load", timeout=90000)
    email_box = page.get_by_placeholder("someone@infineon.com").or_(page.locator("input[name='loginfmt']")).first
    app_ready = page.get_by_text("Book Now!").first
    try:
        email_box.or_(app_ready).first.wait_for(state="visible", timeout=30000)
    except PWTimeout:
        raise RuntimeError(f"Neither the login page nor the dashboard appeared (at {page.url})")
    if app_ready.is_visible():
        log("No login page shown - already signed in")
        return

    log("Login page detected, signing in...")
    email_box.fill(EMAIL)
    page.get_by_role("button", name="Next").click()
    log("Email submitted...")

    # Keep checking until the dashboard shows up or the time runs out.
    deadline = time.time() + LOGIN_TIMEOUT_SEC
    while time.time() < deadline:
        if on_app(page) and app_ready.is_visible():
            log("Signed in")
            return
        if page.get_by_text("Stay signed in?").is_visible():
            page.get_by_role("button", name="Yes").click()
            log("'Stay signed in?' -> Yes")
        page.wait_for_timeout(1000)

    raise RuntimeError(f"Sign-in did not finish within {LOGIN_TIMEOUT_SEC}s (stuck at {page.url})")


# --- Booking pieces ---

# Click one day in the BookMySpace date picker (flatpickr). Each cell has an aria-label like "October 20, 2026".
def click_day(page: Page, day: dt.date) -> None:
    label = f"{day:%B} {day.day}, {day.year}"
    cell = page.locator(f".flatpickr-day[aria-label='{label}']:visible").first
    try:
        cell.wait_for(state="visible", timeout=10000)
    except PWTimeout:
        raise RuntimeError(f"{day} not found in the calendar")

    classes = cell.get_attribute("class") or ""
    if "flatpickr-disabled" in classes or "notAllowed" in classes:
        raise RuntimeError(f"{day} is not selectable in the calendar")
    cell.click()
    log(f"  calendar: clicked {day}")


# Pick SPACE in the Space filter (a virtual-select dropdown): open it, type the name, click the match.
# Returns False if the space is not offered for the chosen dates.
def select_space(page: Page) -> bool:
    toggle = page.locator(".vscomp-toggle-button:visible").last
    search = page.locator(".vscomp-search-input:visible")
    for _ in range(3):  # the widget can ignore a click made before it has finished loading
        if not search.is_visible():
            toggle.click()
        try:
            search.wait_for(state="visible", timeout=3000)
            break
        except PWTimeout:
            log("  Space dropdown did not open, retrying...")
    else:
        raise RuntimeError("Space dropdown never opened")
    search.fill(SPACE)
    option = page.locator(".vscomp-option:visible", has_text=SPACE).first
    try:
        option.wait_for(state="visible", timeout=5000)
    except PWTimeout:
        return False
    option.click()
    page.get_by_text("1 to 1 of 1 items").wait_for()  # list has been filtered down to the chosen space
    return True


# Read the days you have already booked from the "Upcoming bookings" table (first page of rows only).
def get_booked_days(page: Page) -> set[dt.date]:
    cells = page.locator("td[data-header='Day']")
    try:
        cells.first.wait_for(state="attached", timeout=3000)
    except PWTimeout:
        return set()  # no upcoming bookings
    days = set()
    for text in cells.all_text_contents():
        match = re.search(r"\d{4}-\d{2}-\d{2}", text)
        if match:
            days.add(dt.date.fromisoformat(match.group()))
    return days


# Book one run of consecutive days on BookMySpace.
# Returns text for the "Booked" summary; raises NotBooked(reason) if it was deliberately skipped.
def book_run(page: Page, start: dt.date, end: dt.date) -> str:
    log(f"--- Booking {period_label(start, end)} ---")

    log("Step 1: open the dashboard, click 'Book Now!'")
    page.goto(BASE_URL, wait_until="networkidle", timeout=90000)
    if not on_app(page):
        raise RuntimeError(f"Redirected away from BookMySpace (SSO?): {page.url}")
    page.get_by_role("button", name="Book Now!").or_(page.get_by_role("link", name="Book Now!")).first.click()

    log("Step 2: location is pre-filled, click 'Next'")
    page.get_by_role("button", name="Next", exact=True).click()

    log("Step 3: skip days that are already booked, then select the dates")
    days = [start + dt.timedelta(days=i) for i in range((end - start).days + 1)]
    already = get_booked_days(page)
    free = [d for d in days if d not in already]
    if not free:
        log("- Every day of this period is already booked")
        return f"{period_label(start, end)} - already booked"
    if len(free) != (free[-1] - free[0]).days + 1:
        raise NotBooked("Existing bookings split the period into separate parts")
    note = ""
    if len(free) < len(days):
        had = [day_label(d) for d in days if d in already]
        log(f"- Already booked: {had}; booking the remaining days only")
        note = f" - newly booked; already had {', '.join(had)}"
    start, end = free[0], free[-1]
    click_day(page, start)
    click_day(page, end)  # the range picker needs a second click even for a one-day booking
    page.locator("button", has_text="Search").click()  # a tooltip <div role=button> also says 'Search', so use <button>

    expected = f"{start:%Y-%m-%d} to {end:%Y-%m-%d}"
    space_tab = page.get_by_role("tab", name="Space List").or_(page.get_by_text("Space List", exact=True)).first
    try:
        space_tab.wait_for(state="visible", timeout=15000)
    except PWTimeout:
        raise RuntimeError(f"Search results for {expected} never appeared - date selection probably failed")
    log(f"- Searched {expected}")

    log(f"Step 4: look for {SPACE} in the Space box")
    space_tab.click()
    if not select_space(page):
        raise NotBooked(f"{SPACE} is not available for these dates")
    log(f"- {SPACE} is available, clicking Book")
    # Only look inside the Space List table: the hidden Upcoming bookings tab also has rows with the space name.
    space_list = page.locator("[data-block='MainFlow.WB_BookingAvailableDeskList']")
    book_button = space_list.locator("tr", has_text=SPACE).locator("button", has_text="Book").first
    try:
        book_button.wait_for(state="visible", timeout=10000)
    except PWTimeout:
        text = page.evaluate("document.body.innerText")
        shown = " ".join(text[text.find("Space List"):][:400].split())
        raise RuntimeError(f"No Book button for {SPACE}. Page shows: {shown}")
    book_button.click()

    log("Step 5: confirm the booking in the popup")
    dialog = page.locator("[role='dialog']").filter(has_text="Booking Confirmation").first
    dialog.wait_for(state="visible", timeout=15000)
    if not CONFIRM_BOOKING:
        dialog.locator("button", has_text="Cancel").click()
        raise NotBooked("test run: CONFIRM_BOOKING is False, so the confirmation popup was cancelled")
    dialog.locator("button", has_text="Book").click()
    dialog.wait_for(state="hidden", timeout=30000)
    log(f"- Booked {SPACE} for {expected}")
    return period_label(start, end) + note


###########
# Main flow
###########
# Print the final summary to the console.
def print_summary(booked: list[str], leave_blocked: dict[dt.date, str], failed: list[tuple[str, str]]) -> None:
    log("===== SUMMARY =====")
    log("Booked:" + ("" if booked else " none"))
    for label in booked:
        log(f"  - {label}")
    log("Not booked (leave portal):" + ("" if leave_blocked else " none"))
    for day, reason in sorted(leave_blocked.items()):
        log(f"  - {day_label(day)}: {reason}")
    log("Failed:" + ("" if failed else " none"))
    for label, reason in failed:
        log(f"  - {label}: {reason}")


def run(playwright: Playwright) -> None:
    # STEP 1: today's date -> next week's Tue / Wed / Thu.
    today = dt.date.today()
    target_days = get_target_dates(today)
    log(f"Today: {today} ({today:%A}); target days: {', '.join(day_label(d) for d in target_days)}")

    browser = playwright.chromium.launch(channel=BROWSER, headless=HEADLESS)
    context = browser.new_context()
    page = context.new_page()
    page.set_default_timeout(20000)

    booked: list[str] = []
    failed: list[tuple[str, str]] = []  # (what, reason) for bookings that errored or were skipped
    leave_blocked: dict[dt.date, str] = {}

    try:
        # STEP 2: leave portal first. Leave / holiday / non-workday days are removed from the list.
        try:
            leave_blocked = check_leave_portal(page, target_days)
        except Exception as e:
            # If we cannot tell which days are leave, do not book anything.
            log(f"ERROR checking the leave portal: {e}")
            failed.append((period_label(target_days[0], target_days[-1]), f"leave check failed, nothing booked: {first_line(e)}"))
            target_days = []

        bookable = [d for d in target_days if d not in leave_blocked]
        if leave_blocked:
            send_leave_email(leave_blocked, bookable)

        # STEP 3: BookMySpace for the remaining days (one booking per run of consecutive days).
        if bookable:
            try:
                sign_in(page)
                for start, end in group_consecutive(bookable):
                    try:
                        booked.append(book_run(page, start, end))
                    except NotBooked as e:
                        failed.append((period_label(start, end), str(e)))
                        log(f"NOT BOOKED {period_label(start, end)}: {e}")
                    except Exception as e:
                        failed.append((period_label(start, end), first_line(e)))
                        log(f"ERROR for {period_label(start, end)}: {e}")
            except Exception as e:
                log(f"ERROR during sign-in: {e}")
                failed += [(period_label(s, t), f"sign-in failed: {first_line(e)}") for s, t in group_consecutive(bookable)]
        else:
            log("No days left to book")
    finally:
        context.close()
        browser.close()

    print_summary(booked, leave_blocked, failed)

    # Booking failures get their own email (leave/holiday days were already emailed above).
    if failed:
        lines = [f"BookMySpace run on {dt.datetime.now():%Y-%m-%d %H:%M}", "", "Failed:"]
        lines += [f"  - {label}: {reason}" for label, reason in failed]
        lines += ["", "Booked:" + ("" if booked else " none")]
        lines += [f"  - {label}" for label in booked]
        send_email(f"BookMySpace: {len(failed)} booking failure(s)", lines)


if __name__ == "__main__":
    check_config()
    with sync_playwright() as pw:
        run(pw)
