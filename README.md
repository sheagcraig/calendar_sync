# Google Calendar Sync Setup Instructions

This Google Apps Script automatically syncs events from your personal Google Calendar to your work Google Calendar as "Busy" blocks, protecting your privacy while preventing scheduling conflicts.

## Features

- ✅ **Smart Event Syncing**: Syncs personal calendar events to work calendar as "Busy" blocks
- ✅ **Recurring Events Support**: Properly handles complex recurring event patterns (daily, weekly, monthly)
- ✅ **Drive Time Management**: Automatically adds 30-minute buffer blocks before/after events with locations
- ✅ **Weekend Filtering**: Optionally skip syncing weekend events (configurable)
- ✅ **Availability Filtering**: Optionally skip events marked as "Free" instead of "Busy" (configurable)
- ✅ **Visual Organization**: Apply colored labels to synced events for easy identification
- ✅ **Intelligent Cleanup**: Removes orphaned events and handles filtered event cleanup
- ✅ **Change Detection**: Detects and handles event changes (time updates, deletions, location changes)
- ✅ **Robust Error Handling**: Comprehensive logging and error recovery
- ✅ **Manual Management**: Reset and diagnostic functions

## Setup Steps

### 1. Create a New Google Apps Script Project

1. Go to [Google Apps Script](https://script.google.com)
2. Click "New Project"
3. Replace the default code with the contents of `calendar-sync.gs`
4. Save the project with a meaningful name (e.g., "Personal Calendar Sync")

### 2. Configure Calendar IDs

You need to find the Calendar IDs for both your personal and work calendars:

#### Finding Your Personal Calendar ID:
1. Open [Google Calendar](https://calendar.google.com)
2. In the left sidebar, find your personal calendar
3. Click the three dots next to it → "Settings and sharing"
4. Scroll down to "Calendar ID" and copy it
5. It will look like: `your-email@gmail.com` or `abc123def456@group.calendar.google.com`

#### Finding Your Work Calendar ID:
- If using your primary work calendar: use `'primary'`
- For a specific work calendar: follow the same steps as above

#### Update the Script:
In the script, update these lines at the top:
```javascript
const PERSONAL_CALENDAR_ID = 'your-personal-calendar@gmail.com'; // Replace with actual ID
const WORK_CALENDAR_ID = 'primary'; // Use 'primary' or specific calendar ID
```

### 3. Test Calendar Access

1. In the Apps Script editor, select the `testCalendarAccess` function from the dropdown
2. Click the "Run" button (▶️)
3. Grant permissions when prompted:
   - Allow access to your Google Calendar
   - Allow access to view and manage your calendars
4. Check the execution log to confirm both calendars are accessible

### 4. Run Initial Sync

1. Select the `syncCalendars` function from the dropdown
2. Click "Run" to perform the first sync
3. Check your work calendar to verify "Busy" events were created
4. Review the execution log for any errors

### 5. Set Up Automatic Triggers

To run the sync automatically:

1. In the Apps Script editor, click the clock icon (⏰) in the left sidebar ("Triggers")
2. Click "+ Add Trigger"
3. Configure the trigger:
   - **Function to run**: `syncCalendars`
   - **Event source**: Time-driven
   - **Type of time based trigger**: Minutes timer
   - **Select minute interval**: Every 15 minutes (or your preference)
4. Save the trigger

**Recommended sync frequency:**
- Every 15-30 minutes for active use
- Every hour for less frequent changes
- Consider your Google Apps Script quotas

## Configuration Options

You can customize these settings in the script:

### Basic Configuration
```javascript
const PERSONAL_CALENDAR_ID = 'sheagcraig@gmail.com'; // Your personal calendar ID
const WORK_CALENDAR_ID = 'primary'; // Your work calendar ID
const SYNC_DAYS_AHEAD = 30;    // How many days ahead to sync
const SYNC_DAYS_BEHIND = 7;    // How many days behind to sync (for updates)
const BUSY_EVENT_TITLE = 'Busy'; // Title for blocked time events
```

### Advanced Features
```javascript
// Drive Time Management
const DRIVE_TIME_MINUTES = 30; // Minutes to add before/after events with locations

// Visual Organization
const SYNC_EVENT_COLOR = CalendarApp.EventColor.PALE_BLUE; // Color for synced events

// Filtering Options
const CLEANUP_ORPHANED_EVENTS = true; // Automatically clean up untracked "Busy" events
const ONLY_SYNC_BUSY_EVENTS = true;   // Skip events marked as "Free" availability
```

### Weekend Event Filtering
To **enable weekend syncing**, comment out these lines in the main sync loop:
```javascript
// Skip weekend events (Saturday = 6, Sunday = 0)
// Comment out the next 6 lines if you want to sync weekend events
// const eventDay = personalEvent.getStartTime().getDay();
// if (eventDay === 0 || eventDay === 6) {
//   console.log(`Skipping weekend event: "${personalEvent.getTitle()}" on ${personalEvent.getStartTime().toDateString()}`);
//   weekendEventsToCleanup.add(eventInstanceKey);
//   continue;
// }
```

## Usage Notes

### What Gets Synced
- **Weekday events** from your personal calendar (weekends skipped by default)
- **Events marked as "Busy"** (events marked as "Free" are skipped by default)
- **All event types**: single events, recurring events, all-day events
- Events are created as "Busy" blocks with correct start/end times
- **No personal details** are copied (just the time blocks)

### Drive Time Management
- **Events with locations** automatically get 30-minute buffer blocks:
  - **"Drive to" event**: 30 minutes before the main event
  - **Main event**: The actual event time
  - **"Drive from" event**: 30 minutes after the main event
- **Events without locations** get only the main "Busy" block
- All related events are **linked together** for updates/deletions

### Event Management
- **New events**: Automatically created on next sync
- **Changed events**: Times, locations, and availability updated automatically
- **Location changes**: Drive time events added/removed as needed
- **Deleted events**: Removed from work calendar (including associated drive time)
- **Recurring events**: Each instance handled separately with unique tracking
- **Weekend filtering**: Previously synced weekend events are automatically removed
- **Availability changes**: Events that change from "Busy" to "Free" are removed

### Visual Organization
- **Colored events**: All synced events use the configured color (default: Pale Blue)
- **Consistent titles**: All events titled "Busy" for privacy
- **Descriptive details**: Event descriptions indicate auto-sync and drive time purpose

### Privacy Protection
- **No personal information** is synced (titles, descriptions, attendees, etc.)
- **Location privacy**: Only used to determine if drive time is needed
- **All work calendar events** are simply titled "Busy"
- **Original personal events** remain completely private

## Troubleshooting

### Common Issues

**"Cannot access calendar" error:**
- Verify calendar IDs are correct
- Ensure you have access to both calendars
- Check that calendars aren't deleted or hidden

**Permission errors:**
- Re-run the authorization process
- Make sure you grant all requested permissions
- Try running `testCalendarAccess()` first

**Events not syncing:**
- Check the execution log for errors
- Verify the trigger is set up and running
- Ensure you're within the sync date range

**Duplicate events:**
- Use the `resetSync()` function to clean up and start fresh
- Check that you don't have multiple triggers running

**Recurring events not syncing:**
- Check that the personal calendar is properly shared with your work account
- Verify recurring events appear in the debug output from `debugRecurringEvents()`
- Each recurring instance should have a unique tracking key

**Drive time events not appearing:**
- Ensure your personal events have location information filled in
- Check that `DRIVE_TIME_MINUTES` is set to a reasonable value (default: 30)
- Verify the main event was created successfully first

**Colors not applied:**
- Run `testColorApplication()` to verify color functionality
- Colors may not appear immediately due to Google Calendar UI caching
- Try refreshing your calendar view

**Weekend events still appearing:**
- Run `resetSync()` to clean up previously synced weekend events
- Check that weekend filtering code is not commented out
- Verify the weekend cleanup logic is working in the logs

### Manual Functions

**Reset everything:**
```javascript
resetSync() // Removes all synced events and clears sync data
```

**Test setup:**
```javascript
testCalendarAccess() // Verifies you can access both calendars
```

**Debug recurring events:**
```javascript
debugRecurringEvents() // Shows how recurring events are fetched
listAllCalendars() // Lists all accessible calendars and their IDs
```

**Test color application:**
```javascript
testColorApplication() // Creates a test event to verify color functionality
```

**Clean up orphaned events:**
```javascript
cleanupOrphanedEvents() // Manually clean up untracked "Busy" events
```

### Viewing Logs
1. In Apps Script editor, click "Executions" in the left sidebar
2. Click on any execution to see detailed logs
3. Look for errors or confirmation messages

## Quota Limits

Google Apps Script has daily quotas:
- **Calendar events read**: 100,000 per day
- **Calendar events write**: 5,000 per day
- **Execution time**: 6 minutes per execution

For typical personal use, these limits should be sufficient.

## Security Considerations

- The script only accesses your Google Calendars
- No data is sent to external services
- Sync data is stored in Google's PropertiesService (encrypted)
- You can revoke access anytime via Google Account settings

## Support

If you encounter issues:
1. Check the execution logs in Apps Script
2. Verify your calendar IDs and permissions
3. Try the `testCalendarAccess()` function
4. Use `resetSync()` if you need to start over

The script includes comprehensive error handling and logging to help diagnose issues.
