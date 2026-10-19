# BookMySpace Desk Booking

## The Problem

The office’s hot-desking policy requires employees to book desks manually every two weeks. This repetitive process is time-consuming and can lead to missed bookings, leaving employees without a desk when they arrive at the office.

## The Solution

`book_my_space.py` automates desk bookings by checking the configured dates for the upcoming week against your leave calendar. It skips non-working days, public holidays, and days when you’re on annual leave, then books your selected desk for the remaining dates.

If a booking issue occurs, the script sends you an email notification via Outlook.

## Running it

### Prerequisites

- Windows with Python installed.
- Google Chrome installed. The script launches Chrome through Playwright.
- Desktop Outlook configured with an account that can send email. Outlook is used for booking notifications.
- Access to the BookMySpace and leave-calendar websites.

### Set up

Open PowerShell in the project folder and create a virtual environment, then install the dependencies:

```text
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```


Set the values in `.env` beside `book_my_space.py`:

| Setting | Description |
| --- | --- |
| `BOOKMYSPACE_BASE_URL` | URL of the BookMySpace site. |
| `LEAVE_PORTAL_URL` | URL of the leave calendar. |
| `BOOKMYSPACE_EMAIL` | Email address used to sign in to BookMySpace. |
| `BOOKMYSPACE_NOTIFY_EMAIL` | Recipient for booking and skipped-date notifications sent through Outlook. |
| `BOOKMYSPACE_SPACE` | Desk identifier to book. |
| `BOOKMYSPACE_WEEKDAYS` | Comma-separated weekdays to book: Monday is `0`, Tuesday `1`, through Sunday `6`. For Tuesday to Thursday, use `1,2,3`. |
| `BOOKMYSPACE_BROWSER` | Installed Chrome channel to launch; normally `chrome`. |
| `BOOKMYSPACE_LOGIN_TIMEOUT_SEC` | Maximum time, in seconds, to wait for sign-in to complete. |
| `LEAVE_PORTAL_WAIT_SEC` | Maximum time, in seconds, to wait for the leave calendar to load. |


### Run manually

From the project folder, with the virtual environment activated, run:

```powershell
python .\book_my_space.py
```

The script calculates the selected weekdays for the next calendar week. Review the run summary and any Outlook notifications to confirm the result.

### Schedule with Windows Task Scheduler

Create a **Basic Task** in Task Scheduler and configure it to run on a schedule. In **Action**, choose **Start a program** and set:

- **Program/script:** the full path to `.venv\Scripts\python.exe` in this project.
- **Add arguments:** the full path to `book_my_space.py`, enclosed in quotes if the path contains spaces.
- **Start in:** the project folder containing `book_my_space.py` and `.env`.

In the task's general settings, choose **Run only when the user is logged on** so the task can access desktop Outlook. Make sure the computer is awake and connected to the network at the scheduled time.