/**
 * Google Calendar Sync Script
 * Syncs personal calendar events to work calendar as "Busy" blocks
 * 
 * Setup Instructions:
 * 1. Replace PERSONAL_CALENDAR_ID with your personal calendar ID
 * 2. Replace WORK_CALENDAR_ID with your work calendar ID (or use 'primary')
 * 3. Set up a time-based trigger to run syncCalendars() regularly
 * 4. Grant necessary permissions when prompted
 */

// Configuration - Update these with your actual calendar IDs
const PERSONAL_CALENDAR_ID = 'sheagcraig@gmail.com'; // Your personal calendar ID
const WORK_CALENDAR_ID = 'primary'; // Use 'primary' for main work calendar or specific calendar ID
const SYNC_DAYS_AHEAD = 30; // How many days ahead to sync
const SYNC_DAYS_BEHIND = 7; // How many days behind to sync (for updates)

// Event title for blocked time
const BUSY_EVENT_TITLE = 'Busy';
const DRIVE_TIME_MINUTES = 30; // Minutes to add before and after events with locations

// Color for synced events (see Google Calendar color options below)
// Available colors: PALE_BLUE, PALE_GREEN, MAUVE, PALE_RED, YELLOW, ORANGE, CYAN, GRAY, BLUE, GREEN, RED
const SYNC_EVENT_COLOR = CalendarApp.EventColor.YELLOW;

// Orphaned event cleanup - set to true to automatically clean up orphaned "Busy" events
// that match our sync pattern but aren't tracked in our sync data
const CLEANUP_ORPHANED_EVENTS = true;

// Availability filtering - set to true to only sync events marked as "Busy" (skip "Free" events)
// When true: only syncs events that show as "Busy" on your personal calendar
// When false: syncs all events regardless of availability setting
const ONLY_SYNC_BUSY_EVENTS = true;

// Property key for storing sync metadata
const SYNC_PROPERTY_KEY = 'CALENDAR_SYNC_DATA';

/**
 * Main function to sync calendars
 * This should be called by a time-based trigger
 */
function syncCalendars() {
  try {
    console.log('Starting calendar sync...');
    
    // Get calendars
    const personalCalendar = CalendarApp.getCalendarById(PERSONAL_CALENDAR_ID);
    const workCalendar = CalendarApp.getCalendarById(WORK_CALENDAR_ID);
    
    if (!personalCalendar) {
      throw new Error(`Cannot access personal calendar: ${PERSONAL_CALENDAR_ID}`);
    }
    
    if (!workCalendar) {
      throw new Error(`Cannot access work calendar: ${WORK_CALENDAR_ID}`);
    }
    
    // Calculate date range
    const now = new Date();
    const startDate = new Date(now.getTime() - (SYNC_DAYS_BEHIND * 24 * 60 * 60 * 1000));
    const endDate = new Date(now.getTime() + (SYNC_DAYS_AHEAD * 24 * 60 * 60 * 1000));
    
    console.log(`Syncing events from ${startDate.toISOString()} to ${endDate.toISOString()}`);
    
    // Get existing sync data
    const syncData = getSyncData();
    
    // Get personal calendar events
    const personalEvents = personalCalendar.getEvents(startDate, endDate);
    console.log(`Found ${personalEvents.length} personal calendar events`);
    
    // Get existing work calendar events that we created
    const workEvents = workCalendar.getEvents(startDate, endDate)
      .filter(event => event.getTitle() === BUSY_EVENT_TITLE);
    
    // Process personal events
    const currentPersonalEventKeys = new Set();
    const weekendEventsToCleanup = new Set();
    const freeEventsToCleanup = new Set();
    
    for (const personalEvent of personalEvents) {
      // Skip all-day events if desired (uncomment next line to skip)
      // if (personalEvent.isAllDayEvent()) continue;
      
      const personalEventId = personalEvent.getId();
      
      // Create a unique key for this specific event instance
      // For recurring events, we need to include the start time to make each instance unique
      const eventInstanceKey = personalEvent.isRecurringEvent() 
        ? `${personalEventId}_${personalEvent.getStartTime().getTime()}`
        : personalEventId;
      
      // Check availability/transparency (Free vs Busy)
      const eventTransparency = personalEvent.getTransparency();
      const isFreeEvent = (eventTransparency === CalendarApp.EventTransparency.TRANSPARENT);
      
      // Skip weekend events (Saturday = 6, Sunday = 0)
      // Comment out the next 6 lines if you want to sync weekend events
      const eventDay = personalEvent.getStartTime().getDay();
      if (eventDay === 0 || eventDay === 6) {
        console.log(`Skipping weekend event: "${personalEvent.getTitle()}" on ${personalEvent.getStartTime().toDateString()}`);
        // Track weekend events for cleanup - we want to remove any previously synced weekend events
        weekendEventsToCleanup.add(eventInstanceKey);
        continue;
      }
      
      // Skip "Free" events if filtering is enabled
      // Comment out the next 6 lines if you want to sync "Free" events
      if (ONLY_SYNC_BUSY_EVENTS && isFreeEvent) {
        console.log(`Skipping "Free" event: "${personalEvent.getTitle()}" (availability set to Free)`);
        // Track free events for cleanup - we want to remove any previously synced free events
        freeEventsToCleanup.add(eventInstanceKey);
        continue;
      }
      
      currentPersonalEventKeys.add(eventInstanceKey);
      
      const eventData = {
        id: personalEventId,
        instanceKey: eventInstanceKey,
        title: personalEvent.getTitle(),
        startTime: personalEvent.getStartTime(),
        endTime: personalEvent.getEndTime(),
        lastModified: personalEvent.getLastUpdated(),
        isRecurring: personalEvent.isRecurringEvent(),
        location: personalEvent.getLocation() || ''
      };
      
      console.log(`Processing event: "${eventData.title}" (${eventData.isRecurring ? 'Recurring' : 'Single'}) - Key: ${eventInstanceKey}`);
      
      // Check if we need to create or update the work event
      const existingSyncRecord = syncData.events[eventInstanceKey];
      
      if (!existingSyncRecord) {
        // New event - create it
        console.log(`Creating new work event for: ${eventData.title}`);
        createWorkEvent(workCalendar, eventData, syncData);
      } else if (hasEventChanged(eventData, existingSyncRecord)) {
        // Event changed - update it
        console.log(`Updating work event for: ${eventData.title}`);
        updateWorkEvent(workCalendar, eventData, existingSyncRecord, syncData);
      } else {
        console.log(`No changes needed for: ${eventData.title}`);
      }
    }
    
    // Clean up work events for personal events that no longer exist OR are now filtered out (like weekends or free events)
    cleanupRemovedEvents(workCalendar, syncData, currentPersonalEventKeys, weekendEventsToCleanup, freeEventsToCleanup);
    
    // Clean up orphaned events if enabled
    if (CLEANUP_ORPHANED_EVENTS) {
      cleanupOrphanedEvents(workCalendar, syncData, startDate, endDate);
    }
    
    // Save updated sync data
    saveSyncData(syncData);
    
    console.log('Calendar sync completed successfully');
    
  } catch (error) {
    console.error('Error during calendar sync:', error);
    // Optionally send email notification about the error
    // MailApp.sendEmail('your-email@company.com', 'Calendar Sync Error', error.toString());
  }
}

/**
 * Creates a new "Busy" event (and drive time events if needed) on the work calendar
 */
function createWorkEvent(workCalendar, personalEventData, syncData) {
  try {
    const hasLocation = personalEventData.location && personalEventData.location.trim() !== '';
    const createdEvents = [];
    
    // Create drive time events if the event has a location
    if (hasLocation) {
      console.log(`Event "${personalEventData.title}" has location: "${personalEventData.location}" - adding drive time`);
      
      // Create "Drive to" event (30 minutes before)
      const driveToStart = new Date(personalEventData.startTime.getTime() - (DRIVE_TIME_MINUTES * 60 * 1000));
      const driveToEnd = personalEventData.startTime;
      
      const driveToEvent = workCalendar.createEvent(
        BUSY_EVENT_TITLE,
        driveToStart,
        driveToEnd,
        {
          description: `Auto-synced drive time to: ${personalEventData.title}`
        }
      );
      driveToEvent.setColor(SYNC_EVENT_COLOR);
      createdEvents.push({
        type: 'drive_to',
        eventId: driveToEvent.getId(),
        startTime: driveToStart.getTime(),
        endTime: driveToEnd.getTime()
      });
      
      console.log(`Created drive-to event: ${driveToStart.toISOString()} to ${driveToEnd.toISOString()}`);
    }
    
    // Create the main event
    const mainEvent = workCalendar.createEvent(
      BUSY_EVENT_TITLE,
      personalEventData.startTime,
      personalEventData.endTime,
      {
        description: hasLocation 
          ? `Auto-synced from personal calendar (at: ${personalEventData.location})`
          : 'Auto-synced from personal calendar'
      }
    );
    mainEvent.setColor(SYNC_EVENT_COLOR);
    createdEvents.push({
      type: 'main',
      eventId: mainEvent.getId(),
      startTime: personalEventData.startTime.getTime(),
      endTime: personalEventData.endTime.getTime()
    });
    
    console.log(`Created main event: ${personalEventData.startTime.toISOString()} to ${personalEventData.endTime.toISOString()}`);
    
    // Create drive time events if the event has a location
    if (hasLocation) {
      // Create "Drive from" event (30 minutes after)
      const driveFromStart = personalEventData.endTime;
      const driveFromEnd = new Date(personalEventData.endTime.getTime() + (DRIVE_TIME_MINUTES * 60 * 1000));
      
      const driveFromEvent = workCalendar.createEvent(
        BUSY_EVENT_TITLE,
        driveFromStart,
        driveFromEnd,
        {
          description: `Auto-synced drive time from: ${personalEventData.title}`
        }
      );
      driveFromEvent.setColor(SYNC_EVENT_COLOR);
      createdEvents.push({
        type: 'drive_from',
        eventId: driveFromEvent.getId(),
        startTime: driveFromStart.getTime(),
        endTime: driveFromEnd.getTime()
      });
      
      console.log(`Created drive-from event: ${driveFromStart.toISOString()} to ${driveFromEnd.toISOString()}`);
    }
    
    // Store sync record using the instance key for proper tracking
    syncData.events[personalEventData.instanceKey] = {
      personalEventId: personalEventData.id,
      instanceKey: personalEventData.instanceKey,
      lastSyncedTime: personalEventData.startTime.getTime(),
      lastSyncedEndTime: personalEventData.endTime.getTime(),
      lastModified: personalEventData.lastModified.getTime(),
      hasLocation: hasLocation,
      location: personalEventData.location,
      createdEvents: createdEvents,
      created: new Date().getTime()
    };
    
    const eventCount = createdEvents.length;
    const eventTypes = createdEvents.map(e => e.type).join(', ');
    console.log(`Created ${eventCount} work events for "${personalEventData.title}": ${eventTypes}`);
    
  } catch (error) {
    console.error(`Failed to create work event for ${personalEventData.title}:`, error);
  }
}

/**
 * Updates existing "Busy" events (including drive time events) on the work calendar
 */
function updateWorkEvent(workCalendar, personalEventData, syncRecord, syncData) {
  try {
    const hasLocation = personalEventData.location && personalEventData.location.trim() !== '';
    const hadLocation = syncRecord.hasLocation || false;
    const locationChanged = personalEventData.location !== (syncRecord.location || '');
    
    // Check if we need to recreate events due to location changes
    if (hasLocation !== hadLocation || locationChanged) {
      console.log(`Location changed for "${personalEventData.title}" - recreating events`);
      console.log(`  Had location: ${hadLocation}, Has location: ${hasLocation}`);
      console.log(`  Old location: "${syncRecord.location || ''}", New location: "${personalEventData.location}"`);
      
      // Delete all existing events and recreate
      deleteAllEventsForRecord(workCalendar, syncRecord);
      delete syncData.events[personalEventData.instanceKey];
      createWorkEvent(workCalendar, personalEventData, syncData);
      return;
    }
    
    // Update existing events
    let updatedCount = 0;
    const eventsToUpdate = syncRecord.createdEvents || [];
    
    for (const eventRecord of eventsToUpdate) {
      try {
        const workEvent = workCalendar.getEventById(eventRecord.eventId);
        if (workEvent) {
          let newStartTime, newEndTime;
          
          if (eventRecord.type === 'drive_to') {
            newStartTime = new Date(personalEventData.startTime.getTime() - (DRIVE_TIME_MINUTES * 60 * 1000));
            newEndTime = personalEventData.startTime;
          } else if (eventRecord.type === 'main') {
            newStartTime = personalEventData.startTime;
            newEndTime = personalEventData.endTime;
          } else if (eventRecord.type === 'drive_from') {
            newStartTime = personalEventData.endTime;
            newEndTime = new Date(personalEventData.endTime.getTime() + (DRIVE_TIME_MINUTES * 60 * 1000));
          }
          
          workEvent.setTime(newStartTime, newEndTime);
          
          // Ensure color is applied (in case it wasn't applied during creation)
          try {
            workEvent.setColor(SYNC_EVENT_COLOR);
          } catch (colorError) {
            console.log(`Note: Could not set color on updated event: ${colorError}`);
          }
          
          // Update the event record
          eventRecord.startTime = newStartTime.getTime();
          eventRecord.endTime = newEndTime.getTime();
          
          updatedCount++;
          console.log(`Updated ${eventRecord.type} event: ${newStartTime.toISOString()} to ${newEndTime.toISOString()}`);
        } else {
          console.log(`Work event ${eventRecord.eventId} (${eventRecord.type}) not found - will recreate`);
        }
      } catch (error) {
        console.error(`Error updating ${eventRecord.type} event ${eventRecord.eventId}:`, error);
      }
    }
    
    // Update sync record
    syncRecord.lastSyncedTime = personalEventData.startTime.getTime();
    syncRecord.lastSyncedEndTime = personalEventData.endTime.getTime();
    syncRecord.lastModified = personalEventData.lastModified.getTime();
    syncRecord.hasLocation = hasLocation;
    syncRecord.location = personalEventData.location;
    syncRecord.lastUpdated = new Date().getTime();
    
    console.log(`Updated ${updatedCount} work events for: ${personalEventData.title}`);
    
  } catch (error) {
    console.error(`Failed to update work events for ${personalEventData.title}:`, error);
    // If update fails, try to recreate
    console.log(`Attempting to recreate events for: ${personalEventData.title}`);
    try {
      deleteAllEventsForRecord(workCalendar, syncRecord);
      delete syncData.events[personalEventData.instanceKey];
      createWorkEvent(workCalendar, personalEventData, syncData);
    } catch (recreateError) {
      console.error(`Failed to recreate events for ${personalEventData.title}:`, recreateError);
    }
  }
}

/**
 * Helper function to delete all events for a sync record
 */
function deleteAllEventsForRecord(workCalendar, syncRecord) {
  const eventsToDelete = syncRecord.createdEvents || [];
  let deletedCount = 0;
  
  for (const eventRecord of eventsToDelete) {
    try {
      const workEvent = workCalendar.getEventById(eventRecord.eventId);
      if (workEvent) {
        workEvent.deleteEvent();
        deletedCount++;
        console.log(`Deleted ${eventRecord.type} event: ${eventRecord.eventId}`);
      }
    } catch (error) {
      console.error(`Failed to delete ${eventRecord.type} event ${eventRecord.eventId}:`, error);
    }
  }
  
  return deletedCount;
}

/**
 * Removes work calendar events for personal events that no longer exist or are now filtered out
 */
function cleanupRemovedEvents(workCalendar, syncData, currentPersonalEventKeys, weekendEventsToCleanup, freeEventsToCleanup) {
  const eventsToRemove = [];
  
  // Find sync records for events that no longer exist in personal calendar or are now filtered out
  for (const [eventKey, syncRecord] of Object.entries(syncData.events)) {
    if (!currentPersonalEventKeys.has(eventKey) || 
        weekendEventsToCleanup.has(eventKey) || 
        freeEventsToCleanup.has(eventKey)) {
      
      let reason = 'personal event removed';
      if (weekendEventsToCleanup.has(eventKey)) {
        reason = 'weekend event now filtered';
      } else if (freeEventsToCleanup.has(eventKey)) {
        reason = 'free event now filtered';
      }
      
      eventsToRemove.push({ eventKey, syncRecord, reason });
    }
  }
  
  // Remove the work events and sync records
  let totalDeletedEvents = 0;
  for (const { eventKey, syncRecord, reason } of eventsToRemove) {
    try {
      // Delete all events associated with this sync record (main + drive time events)
      const deletedCount = deleteAllEventsForRecord(workCalendar, syncRecord);
      totalDeletedEvents += deletedCount;
      console.log(`Deleted ${deletedCount} work events for ${reason}: ${eventKey}`);
    } catch (error) {
      console.error(`Failed to delete work events for ${eventKey}:`, error);
    }
    
    // Remove sync record
    delete syncData.events[eventKey];
  }
  
  if (eventsToRemove.length > 0) {
    console.log(`Cleaned up ${eventsToRemove.length} removed/filtered events (${totalDeletedEvents} total calendar events deleted)`);
  }
}

/**
 * Cleans up orphaned "Busy" events that match our sync pattern but aren't tracked
 */
function cleanupOrphanedEvents(workCalendar, syncData, startDate, endDate) {
  try {
    console.log('=== CLEANING UP ORPHANED EVENTS ===');
    
    // Get all "Busy" events in the date range
    const allBusyEvents = workCalendar.getEvents(startDate, endDate)
      .filter(event => event.getTitle() === BUSY_EVENT_TITLE);
    
    console.log(`Found ${allBusyEvents.length} total "Busy" events in work calendar`);
    
    // Get all tracked event IDs from our sync data
    const trackedEventIds = new Set();
    for (const syncRecord of Object.values(syncData.events)) {
      if (syncRecord.createdEvents) {
        for (const eventRecord of syncRecord.createdEvents) {
          trackedEventIds.add(eventRecord.eventId);
        }
      }
    }
    
    console.log(`Found ${trackedEventIds.size} tracked event IDs in sync data`);
    
    // Find orphaned events
    const orphanedEvents = [];
    for (const busyEvent of allBusyEvents) {
      const eventId = busyEvent.getId();
      if (!trackedEventIds.has(eventId)) {
        const description = busyEvent.getDescription();
        // Check if this looks like one of our auto-synced events
        if (description && (
          description.includes('Auto-synced from personal calendar') ||
          description.includes('Auto-synced drive time')
        )) {
          orphanedEvents.push({
            event: busyEvent,
            id: eventId,
            description: description,
            startTime: busyEvent.getStartTime(),
            endTime: busyEvent.getEndTime()
          });
        }
      }
    }
    
    console.log(`Found ${orphanedEvents.length} orphaned sync events`);
    
    // Delete orphaned events
    let deletedOrphanedCount = 0;
    for (const orphan of orphanedEvents) {
      try {
        orphan.event.deleteEvent();
        deletedOrphanedCount++;
        console.log(`Deleted orphaned event: ${orphan.id} (${orphan.startTime.toISOString()} - ${orphan.endTime.toISOString()})`);
      } catch (error) {
        console.error(`Failed to delete orphaned event ${orphan.id}:`, error);
      }
    }
    
    if (deletedOrphanedCount > 0) {
      console.log(`✅ Cleaned up ${deletedOrphanedCount} orphaned events`);
    } else {
      console.log('✅ No orphaned events found');
    }
    
    return deletedOrphanedCount;
    
  } catch (error) {
    console.error('Error during orphaned event cleanup:', error);
    return 0;
  }
}

/**
 * Checks if a personal event has changed since last sync
 */
function hasEventChanged(eventData, syncRecord) {
  return (
    eventData.startTime.getTime() !== syncRecord.lastSyncedTime ||
    eventData.endTime.getTime() !== syncRecord.lastSyncedEndTime ||
    eventData.lastModified.getTime() > syncRecord.lastModified
  );
}

/**
 * Gets stored sync data from PropertiesService
 */
function getSyncData() {
  try {
    const stored = PropertiesService.getScriptProperties().getProperty(SYNC_PROPERTY_KEY);
    if (stored) {
      return JSON.parse(stored);
    }
  } catch (error) {
    console.error('Error loading sync data:', error);
  }
  
  // Return default structure
  return {
    events: {},
    lastSync: null
  };
}

/**
 * Saves sync data to PropertiesService
 */
function saveSyncData(syncData) {
  try {
    syncData.lastSync = new Date().getTime();
    PropertiesService.getScriptProperties().setProperty(
      SYNC_PROPERTY_KEY, 
      JSON.stringify(syncData)
    );
  } catch (error) {
    console.error('Error saving sync data:', error);
  }
}

/**
 * Manual cleanup function - removes all synced events and sync data
 * Use this if you need to reset everything
 * 
 * IMPORTANT: This will delete ALL "Busy" events from your work calendar
 * and clear all sync tracking data. Use with caution!
 */
function resetSync() {
  try {
    console.log('=== RESET SYNC STARTING ===');
    console.log('WARNING: This will delete all synced "Busy" events and reset sync data');
    
    const workCalendar = CalendarApp.getCalendarById(WORK_CALENDAR_ID);
    if (!workCalendar) {
      console.error('Cannot access work calendar for reset');
      return;
    }
    
    const syncData = getSyncData();
    console.log(`Found ${Object.keys(syncData.events).length} tracked events to clean up`);
    
    let deletedCount = 0;
    let errorCount = 0;
    
    // Delete all tracked work events
    for (const [eventKey, syncRecord] of Object.entries(syncData.events)) {
      try {
        const eventDeletedCount = deleteAllEventsForRecord(workCalendar, syncRecord);
        deletedCount += eventDeletedCount;
        console.log(`Deleted ${eventDeletedCount} work events for: ${eventKey}`);
      } catch (error) {
        errorCount++;
        console.error(`Error deleting events for ${eventKey}:`, error);
      }
    }
    
    // Also clean up orphaned events during reset if enabled
    if (CLEANUP_ORPHANED_EVENTS) {
      console.log('=== CLEANING UP ORPHANED EVENTS DURING RESET ===');
      const now = new Date();
      const resetStartDate = new Date(now.getTime() - (365 * 24 * 60 * 60 * 1000)); // Look back 1 year
      const resetEndDate = new Date(now.getTime() + (365 * 24 * 60 * 60 * 1000)); // Look ahead 1 year
      
      const orphanedCount = cleanupOrphanedEvents(workCalendar, { events: {} }, resetStartDate, resetEndDate);
      deletedCount += orphanedCount;
    }
    
    // Clear sync data
    PropertiesService.getScriptProperties().deleteProperty(SYNC_PROPERTY_KEY);
    
    console.log('=== RESET SYNC COMPLETED ===');
    console.log(`✅ Deleted ${deletedCount} work calendar events`);
    console.log(`❌ Failed to delete ${errorCount} events`);
    console.log('✅ Cleared all sync tracking data');
    console.log('You can now run syncCalendars() to start fresh');
    
    return {
      success: true,
      deletedCount,
      errorCount,
      message: `Reset complete. Deleted ${deletedCount} events, ${errorCount} errors.`
    };
    
  } catch (error) {
    console.error('❌ Error during reset:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Test function to verify calendar access
 */
function testCalendarAccess() {
  try {
    console.log('=== TESTING CALENDAR ACCESS ===');
    console.log('');
    
    console.log('Configuration:');
    console.log(`Personal Calendar ID: ${PERSONAL_CALENDAR_ID}`);
    console.log(`Work Calendar ID: ${WORK_CALENDAR_ID}`);
    console.log('');
    
    // Test personal calendar access
    console.log('Testing personal calendar access...');
    let personalCalendar = null;
    try {
      personalCalendar = CalendarApp.getCalendarById(PERSONAL_CALENDAR_ID);
      if (personalCalendar) {
        console.log(`✅ Personal calendar found: "${personalCalendar.getName()}"`);
        console.log(`   Calendar ID: ${personalCalendar.getId()}`);
        console.log(`   Description: ${personalCalendar.getDescription() || 'No description'}`);
        console.log(`   Owned by me: ${personalCalendar.isOwnedByMe()}`);
      } else {
        console.log('❌ Personal calendar: NOT FOUND (null returned)');
      }
    } catch (error) {
      console.log(`❌ Personal calendar error: ${error.toString()}`);
    }
    
    console.log('');
    
    // Test work calendar access
    console.log('Testing work calendar access...');
    let workCalendar = null;
    try {
      workCalendar = CalendarApp.getCalendarById(WORK_CALENDAR_ID);
      if (workCalendar) {
        console.log(`✅ Work calendar found: "${workCalendar.getName()}"`);
        console.log(`   Calendar ID: ${workCalendar.getId()}`);
        console.log(`   Description: ${workCalendar.getDescription() || 'No description'}`);
        console.log(`   Owned by me: ${workCalendar.isOwnedByMe()}`);
      } else {
        console.log('❌ Work calendar: NOT FOUND (null returned)');
      }
    } catch (error) {
      console.log(`❌ Work calendar error: ${error.toString()}`);
    }
    
    console.log('');
    console.log('=== SUMMARY ===');
    
    if (personalCalendar && workCalendar) {
      console.log('✅ SUCCESS: Both calendars are accessible!');
      console.log('You can now run syncCalendars() to start syncing.');
    } else {
      console.log('❌ FAILED: Cannot access one or both calendars');
      console.log('');
      console.log('Next steps:');
      if (!personalCalendar) {
        console.log('1. Run listAllCalendars() to see available calendars');
        console.log('2. Share your personal calendar with this work account');
        console.log('3. Update PERSONAL_CALENDAR_ID with the correct ID');
      }
      if (!workCalendar) {
        console.log('1. Verify WORK_CALENDAR_ID is correct');
        console.log('2. Try using "primary" for your main work calendar');
      }
    }
    
    return { personalCalendar: !!personalCalendar, workCalendar: !!workCalendar };
    
  } catch (error) {
    console.error('❌ Calendar access test failed:', error);
    return { error: error.toString() };
  }
}

/**
 * Test function to verify color application works
 */
function testColorApplication() {
  try {
    console.log('=== TESTING COLOR APPLICATION ===');
    
    const workCalendar = CalendarApp.getCalendarById(WORK_CALENDAR_ID);
    if (!workCalendar) {
      console.error('Cannot access work calendar for color test');
      return;
    }
    
    // Create a test event
    const now = new Date();
    const testStart = new Date(now.getTime() + (5 * 60 * 1000)); // 5 minutes from now
    const testEnd = new Date(testStart.getTime() + (30 * 60 * 1000)); // 30 minutes duration
    
    console.log('Creating test event...');
    const testEvent = workCalendar.createEvent(
      'COLOR TEST - DELETE ME',
      testStart,
      testEnd,
      {
        description: 'This is a test event to verify color application. You can delete this.'
      }
    );
    
    console.log(`Test event created: ${testEvent.getId()}`);
    
    // Apply color
    console.log('Applying color...');
    testEvent.setColor(SYNC_EVENT_COLOR);
    
    console.log('✅ Color applied successfully!');
    console.log(`Test event: "${testEvent.getTitle()}" from ${testStart.toISOString()} to ${testEnd.toISOString()}`);
    console.log('Check your work calendar - you should see a colored "COLOR TEST - DELETE ME" event');
    console.log('You can manually delete this test event when done.');
    
    return {
      success: true,
      eventId: testEvent.getId(),
      message: 'Color test event created successfully'
    };
    
  } catch (error) {
    console.error('❌ Color application test failed:', error);
    return { error: error.toString() };
  }
}
