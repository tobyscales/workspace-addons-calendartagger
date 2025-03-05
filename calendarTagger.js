/**
 * @OnlyCurrentDoc
 */

// --- CONFIGURABLE VARIABLES ---
const DEBUG_MODE = false; // Set to true to enable debug logging
const CACHE_TIME = 120; //default is two minutes; set to lower value to troubleshoot
const DEFAULT_USER_TAGS = ['#Work', '#Personal', '#Internal_Meeting', '#External_Meeting']; // Define your default tags here
// --- END OF CONFIGURABLE VARIABLES ---

const userCache = CacheService.getUserCache();
const userProperties = PropertiesService.getUserProperties();

/**
 * Helper function for logging. Only logs if DEBUG_MODE is enabled.
 * @param {string} message - The message to log.
 * @param {...any} params - Optional additional parameters to log.
 */
function log(message, ...params) {
  if (DEBUG_MODE) {
    console.log(`${message}`, ...params);
  }
}

/**
 * Runs when a new event is created or updated in the target calendar.
 * This function initializes or loads tag data, automatically tags events
 * based on attendee emails and the event title, and displays a card for tag selection.
 */
function onCalendarEventOpen(e) {
  log('onCalendarEventOpen called', e);

  const calendarId = e.calendar.calendarId;
  let eventId = e.calendar.id;
  let eventTitle = null;
  let attendees = [];

  // Try to fetch the event to get the title and attendees, handle new events
  if (eventId) {
    try {
      const event = Calendar.Events.get(calendarId, eventId);
      eventTitle = event.summary;
      attendees = event.attendees || [];
      log('Fetched event title:', eventTitle);
      log('Fetched attendees:', attendees);
    } catch (error) {
      log(`Error fetching event: ${error.message}`);
      eventId = null; // Treat it as a new event if fetching fails
    }
  } else {
    log('New event detected.');
  }

  log(`Calendar ID: ${calendarId}, Event ID: ${eventId}, Event Title: ${eventTitle}`);

  if (!calendarId) {
    log('Error: Calendar ID is not available. This may be a compose-mode trigger.');
    return buildErrorCard("Calendar ID is not available. Please try again later.");
  }

  // Store calendarId in cache for later use by saveTagsFromCache
  userCache.put('calendarId', calendarId, 21600); // Cache for 6 hours

  const cacheKey = eventId ? `selectedTags_${eventId}` : `selectedTags_new_${Utilities.getUuid()}`;

  let selectedTags = JSON.parse(userCache.get(cacheKey));

  if (!selectedTags) {
    log('No cached tags found. Loading from event or initializing.');
    selectedTags = loadSelectedTags(calendarId, eventId);

    // Extract tag from title if it's a new event or tags were not loaded
    // Update: Extract all matching tags
    if (!eventId || selectedTags.size === 0) {
      const extractedTags = extractTagFromTitle(eventTitle);
      log(`Extracted Tags from Title: ${extractedTags}`);
      extractedTags.forEach(tag => selectedTags.add(tag));
    }

    // Auto Tag based on attendees
    const autoTag = getAutoTagFromAttendees(attendees);
    log('Auto-tag based on attendees:', autoTag);
    if (autoTag) {
      selectedTags.add(autoTag);
      // Update the title if it's a new event
      if (!eventId) {
        eventTitle = `${autoTag} ${eventTitle || ''}`.trim();
        log('Updated title with auto-tag:', eventTitle);
      }
    }

    userCache.put(cacheKey, JSON.stringify(Array.from(selectedTags)), CACHE_TIME);
  } else {
    log('Loading tags from cacheKey:', cacheKey);
  }

  return showTagDialog(new Set(selectedTags), cacheKey, eventTitle);
}



/**
 * Loads selected tags from event extended properties or initializes an empty set.
 * @param {string} calendarId - The ID of the calendar.
 * @param {string} eventId - The ID of the event.
 * @returns {Set<string>} - The set of selected tags.
 */
function loadSelectedTags(calendarId, eventId) {
  log('loadSelectedTags called', calendarId, eventId);

  if (!eventId) {
    log("New event, initializing selected tags as empty");
    return new Set(); // Early exit for new events
  }

  try {
    const event = Calendar.Events.get(calendarId, eventId, { fields: 'id,extendedProperties' }); // Fetch only necessary fields
    if (event.extendedProperties?.private?.selectedTags) {
      const selectedTags = new Set(JSON.parse(event.extendedProperties.private.selectedTags));
      log("Loaded saved tags:", selectedTags);
      return selectedTags;
    } else {
      log("No tags saved, initializing as empty");
      return new Set();
    }
  } catch (error) {
    log(`Error retrieving event: ${error.message}`);
    return new Set();
  }
}

/**
 * Displays a dialog with tag buttons and event title.
 * @param {Set<string>} selectedTags - The set of currently selected tags.
 * @param {string} cacheKey - The cache key for the current event.
 * @param {string} title - The updated event title.
 */
function showTagDialog(selectedTags, cacheKey, title) {
  log('showTagDialog called', selectedTags);

  const card = rebuildCard(selectedTags, cacheKey, title);

  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().pushCard(card))
    .setStateChanged(true)
    .build();
}

/**
 * Handles the click event for tag buttons.
 * @param {Object} e - The event object.
 */
function handleTagClick(e) {
  log('handleTagClick called', e);

  const clickedTag = e.parameters.tag;
  log('Clicked tag:', clickedTag);

  const cacheKey = e.parameters.cacheKey;
  log('Cache Key from event:', cacheKey);
  if (!cacheKey) {
    log('Error: Cache Key is missing.');
    return buildErrorCard('Cache Key is missing.');
  }

  let selectedTags;
  try {
    selectedTags = new Set(JSON.parse(userCache.get(cacheKey)));
  } catch (error) {
    log(`Error parsing cached tags: ${error.message}`);
    selectedTags = new Set();
  }

  selectedTags.has(clickedTag) ? selectedTags.delete(clickedTag) : selectedTags.add(clickedTag);
  log('Toggled tags:', selectedTags);

  // Update Cache
  userCache.put(cacheKey, JSON.stringify(Array.from(selectedTags)), CACHE_TIME); // Refresh expiration

  // Add event to modified events stack
  pushModifiedEvent(cacheKey);

  const updatedCard = rebuildCard(selectedTags, cacheKey);

  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(updatedCard))
    .setStateChanged(true)
    .build();
}

/**
 * Builds the card with tag buttons based on the selectedTags.
 * @param {Set<string>} selectedTags - The set of currently selected tags.
 * @param {string} cacheKey - The cache key for the current event.
 * @param {string} title - The updated event title.
 */
function rebuildCard(selectedTags, cacheKey, title) {
  log('rebuildCard called', selectedTags);

  const USER_TAGS = new Set(getUserTags());

  const tagButtons = [];

  // Create buttons for selected tags, even if not in USER_TAGS
  for (const tag of selectedTags) {
    const isSelected = true;
    const isUserTag = USER_TAGS.has(tag);

    let button = CardService.newTextButton()
      .setText(tag)
      .setTextButtonStyle(isUserTag ? CardService.TextButtonStyle.FILLED : CardService.TextButtonStyle.TEXT)
      .setOnClickAction(CardService.newAction()
        .setFunctionName('handleTagClick')
        .setParameters({ tag: tag, cacheKey: cacheKey }));

    // Only set background color if it's NOT a user tag
    if (!isUserTag) {
      try {
        button = button.setBackgroundColor("#d3d3d3");
      } catch (error) {
        log(`Error setting background color for tag ${tag}: ${error.message}`);
        // Optional: Handle the error, e.g., add the tag to a list of problem tags to report later
      }
    }

    tagButtons.push(button);
  }

  // Create buttons for USER_TAGS that are not already selected
  for (const tag of USER_TAGS) {
    if (!selectedTags.has(tag)) {
      tagButtons.push(CardService.newTextButton()
        .setText(tag)
        .setTextButtonStyle(CardService.TextButtonStyle.TEXT)
        .setOnClickAction(CardService.newAction()
          .setFunctionName('handleTagClick')
          .setParameters({ tag: tag, cacheKey: cacheKey })));
    }
  }

  const buttonsSet = CardService.newButtonSet();
  tagButtons.forEach(button => buttonsSet.addButton(button));

  const buttonSection = CardService.newCardSection()
    .setHeader("Select Tags")
    .addWidget(buttonsSet);

  const titleSection = CardService.newCardSection()
    .setHeader("Event Title")
    .addWidget(CardService.newTextParagraph()
      .setText(title || ""));

  return CardService.newCardBuilder()
    .addSection(titleSection)
    .addSection(buttonSection)
    .setName('tagCard')
    .build();
}


/**
 * Pushes a cache key onto the ModifiedEvents stack stored in User Cache.
 *
 * @param {string} cacheKey - The cache key to push onto the stack.
 */
function pushModifiedEvent(cacheKey) {
  log('pushModifiedEvent called', cacheKey);

  let stack = JSON.parse(userCache.get('ModifiedEvents')) || {};
  log('Current stack:', stack);

  // Add the new cache key to the stack
  //stack.push(cacheKey);
  stack[cacheKey] = "Dirty"; //Mark the status

  // Store the updated stack in the cache
  userCache.put('ModifiedEvents', JSON.stringify(stack), CACHE_TIME);
  log('Updated stack:', stack);
}

/**
 * Saves the selected tags for all modified events in the cache to their respective events.
 */
function saveTagsFromCache() {
  log('saveTagsFromCache called');

  const stack = JSON.parse(userCache.get('ModifiedEvents'));
  log('Modified events stack:', stack);

  if (!stack) {
    log('No modified events found.');
    return;
  }

  const calendarId = userCache.get('calendarId');
  log('Calendar ID:', calendarId);

  if (!calendarId) {
    log('Error: Calendar ID not found in cache.');
    return;
  }

  // Iterate through the modified events stack
  //for (let cacheKey of stack) {
  for (let cacheKey in stack) {
    if (stack[cacheKey] === "Dirty") {
      const eventId = cacheKey.startsWith('selectedTags_new_') ? null : cacheKey.substring('selectedTags_'.length);
      log(`Processing cache key: ${cacheKey}, Event ID: ${eventId}`);

      const selectedTags = JSON.parse(userCache.get(cacheKey));
      log('Selected tags from cache:', selectedTags);

      if (!selectedTags) {
        log(`No tags found for cache key: ${cacheKey}. Skipping.`);
        continue;
      }

      try {
           if (eventId) {
            // Existing event
            let event = Calendar.Events.get(calendarId, eventId); 
            log('Event retrieved:', event);

            event.extendedProperties = event.extendedProperties || { private: {} };
            event.extendedProperties.private.selectedTags = JSON.stringify(Array.from(selectedTags));

            Calendar.Events.update(event, calendarId, eventId);
            log(`Tags saved to event: ${eventId}`);

            // Invalidate Cache after successful update
            userCache.remove(cacheKey);
          } else {
          // New event - We should still keep it in modifiedEvents stack for later updates
          log('New event. Tags are not saved yet.');
          continue;
        }

        // Remove the processed event from the stack (if it was an existing event)
        //const index = stack.indexOf(cacheKey);
        //if (index > -1) {
        //  stack.splice(index, 1);
        //}

        // Mark as Clean
        stack[cacheKey] = "Clean";

      } catch (error) {
        log(`Error saving tags for event ${eventId}: ${error}`);
      }
    }
  }

  // Update the stack in the cache
  userCache.put('ModifiedEvents', JSON.stringify(stack), 21);
  log('Modified events stack updated:', stack);
}

/**
 * Builds an error card with a given message.
 * @param {string} message - The error message to display.
 * @returns {CardService.Card} The error card.
 */
function buildErrorCard(message) {
  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('Error'))
    .addSection(
      CardService.newCardSection().addWidget(
        CardService.newTextParagraph().setText(message)
      )
    )
    .build();
  return card;
}

/**
 * Triggers every 1 minute
 */
function triggerEveryMinute() {
  ScriptApp.newTrigger("saveTagsFromCache")
    .timeBased()
    .everyMinutes(1)
    .create();
}


/**
 * Extracts all tags from the event title that match valid tags in USER_TAGS (case-insensitive).
 * Returns an array of the matching tags from USER_TAGS.
 * @param {string} title - The event title.
 * @returns {string[]} - An array of matching tags from USER_TAGS.
 */
function extractTagFromTitle(title) {
  if (!title) return [];

  const USER_TAGS = getUserTags();
  const matchingTags = [];

  // Use a regular expression to find all tags in the title (case-insensitive)
  const tagRegex = /#\w+/gi;
  let match;
  while ((match = tagRegex.exec(title)) !== null) {
    const extractedTag = match[0];

    // Find the matching tag in USER_TAGS (case-insensitive)
    const matchingTag = USER_TAGS.find(tag => tag.toLowerCase() === extractedTag.toLowerCase());

    if (matchingTag) {
      matchingTags.push(matchingTag); // Add the matching tag to the array
    }
  }

  return matchingTags;
}

/**
 * Creates the onHomePageOpened trigger and card.
 */
function onHomePageOpen() {
  handleRefreshTags();
  return createHomePageCard();
}

/**
 * Creates the homepage card with spreadsheet configuration, tag display, and refresh button.
 *
 * @returns {CardService.Card} The homepage card.
 */
function createHomePageCard() {
  log('createHomePageCard called');
  const card = CardService.newCardBuilder();

  // Spreadsheet ID input section
  const spreadsheetIdInput = CardService.newTextInput()
    .setFieldName("spreadsheet_id")
    .setTitle("Enter Spreadsheet ID")
    .setValue(userProperties.getProperty("spreadsheetId") || "");

  const sheetNameInput = CardService.newTextInput()
    .setFieldName("sheet_name")
    .setTitle("Enter Sheet Name")
    .setValue(userProperties.getProperty("sheetName") || "");

  const columnInput = CardService.newTextInput()
    .setFieldName("column")
    .setTitle("Enter Tag Column (e.g., A, B, C)")
    .setValue(userProperties.getProperty("column") || "");

  const emailDomainColumnInput = CardService.newTextInput()
    .setFieldName("email_domain_column")
    .setTitle("Enter Email Domain Column (e.g., A, B, C)")
    .setValue(userProperties.getProperty("emailDomainColumn") || "");

  const saveConfigAction = CardService.newAction()
    .setFunctionName("handleSaveConfig");

  const saveConfigButton = CardService.newTextButton()
    .setText("Save Configuration")
    .setOnClickAction(saveConfigAction);

  const configSection = CardService.newCardSection()
    .setHeader("Spreadsheet Configuration")
    .addWidget(spreadsheetIdInput)
    .addWidget(sheetNameInput)
    .addWidget(columnInput)
    .addWidget(emailDomainColumnInput)
    .addWidget(saveConfigButton);

  card.addSection(configSection);

  // Refresh button section
  const refreshAction = CardService.newAction()
    .setFunctionName("handleRefreshTags");

  const refreshButton = CardService.newTextButton()
    .setText("Refresh Tags")
    .setOnClickAction(refreshAction);

  const refreshSection = CardService.newCardSection()
    .setHeader("Refresh Tags")
    .addWidget(refreshButton);

  card.addSection(refreshSection);

  // Current tags section
  const currentTags = getUserTags();
  const currentTagsSection = CardService.newCardSection()
    .setHeader("Current Tags");

  if (currentTags.length > 0) {
    // Create buttons for each tag
    const buttonsSet = CardService.newButtonSet();
    currentTags.forEach(tag => {
      const button = CardService.newTextButton()
        .setText(tag)
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setBackgroundColor("#d3d3d3")
        .setOnClickAction(CardService.newAction()
          .setFunctionName("handleTagClickFromHomepage")
          .setParameters({ tag: tag }));
      buttonsSet.addButton(button);
    });
    currentTagsSection.addWidget(buttonsSet);
  } else {
    currentTagsSection.addWidget(CardService.newTextParagraph()
      .setText("Error: Unable to retrieve tags. Please check your spreadsheet configuration and try refreshing."));
  }

  card.addSection(currentTagsSection);

  log('createHomePageCard finished');
  return card.build();
}

/**
 * Handles refreshing the user tags from the spreadsheet.
 *
 * @returns {CardService.ActionResponse} The action response.
 */
function handleRefreshTags() {
  log('handleRefreshTags called');

  try {
    // Invalidate cached tags
    userProperties.deleteProperty("userTags");

    // Force tag reload
    const tags = getUserTags();
    log('Tags after refresh:', tags);

    log('handleRefreshTags finished');
    // Create a new card with updated tag information
    const updatedCard = createHomePageCard();

    // Update the current card
    return CardService.newActionResponseBuilder()
        .setNavigation(CardService.newNavigation().updateCard(updatedCard))
        .setStateChanged(true)
        .build();

  } catch (error) {
    log(`Error in handleRefreshTags: ${error.message}`, error);
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
      .setText("Error refreshing tags. Check logs."))
      .setStateChanged(true)
      .build();
  }
}
/**
 * Handles saving the spreadsheet configuration from the homepage card.
 *
 * @param {Object} e - The event object.
 * @returns {CardService.ActionResponse} The action response.
 */
function handleSaveConfig(e) {
  log('handleSaveConfig called', e);
  const spreadsheetId = e.formInput.spreadsheet_id;
  const sheetName = e.formInput.sheet_name;
  const column = e.formInput.column;
  const emailDomainColumn = e.formInput.email_domain_column;
  log(`Saving Spreadsheet ID: ${spreadsheetId}, Sheet Name: ${sheetName}, Tag Column: ${column}, Email Domain Column: ${emailDomainColumn}`);

  try {
    userProperties.setProperty("spreadsheetId", spreadsheetId);
    userProperties.setProperty("sheetName", sheetName);
    userProperties.setProperty("column", column);
    userProperties.setProperty("emailDomainColumn", emailDomainColumn);

    // Invalidate cached tags
    userProperties.deleteProperty("userTags");

    log('handleSaveConfig finished');
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText("Configuration saved."))
      .setStateChanged(true)
      .build();
  } catch (error) {
    log(`Error in handleSaveConfig: ${error.message}`, error);
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText("Error saving configuration. Check logs."))
      .setStateChanged(true)
      .build();
  }
}

function getAutoTagFromAttendees(attendees) {
  log('getAutoTagFromAttendees called', attendees);
  const emailDomainColumn = userProperties.getProperty("emailDomainColumn");

  if (!emailDomainColumn) {
    log('Email domain column not configured.');
    return null;
  }

  const spreadsheetId = userProperties.getProperty("spreadsheetId");
  const sheetName = userProperties.getProperty("sheetName");
  const tagColumn = userProperties.getProperty("column");

  if (!spreadsheetId || !sheetName || !tagColumn) {
    log('Spreadsheet ID, sheet name, or tag column not configured.');
    return null;
  }

  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      log(`Sheet "${sheetName}" not found in spreadsheet.`);
      return null;
    }

    const lastRow = sheet.getLastRow();
    if (lastRow === 0) {
      log('Sheet is empty.');
      return null;
    }

    const tagData = sheet.getRange(`${tagColumn}1:${tagColumn}${lastRow}`).getValues();
    const emailDomainData = sheet.getRange(`${emailDomainColumn}1:${emailDomainColumn}${lastRow}`).getValues();

    for (const attendee of attendees) {
      const email = attendee.email;
      if (email) {
        const domain = email.split('@')[1];
        log(`Checking attendee email: ${email}, domain: ${domain}`);

        for (let i = 0; i < emailDomainData.length; i++) {
          if (emailDomainData[i][0] === domain) {
            const tag = tagData[i][0] ? (tagData[i][0].startsWith('#') ? tagData[i][0] : `#${tagData[i][0]}`) : null;
            log(`Found match for domain ${domain}, tag: ${tag}`);
            return tag;
          }
        }
      }
    }
  } catch (error) {
    log(`Error in getAutoTagFromAttendees: ${error.message}`, error);
  }

  log('No matching domain found for attendees.');
  return null;
}

/**
 * Gets the user tags, combining default tags and tags from the spreadsheet.
 *
 * @returns {string[]} The user tags.
 */
function getUserTags() {
  log('getUserTags called');
  let userTags = JSON.parse(userProperties.getProperty("userTags"));

  if (!userTags) {
    log('User tags not found in cache. Fetching from spreadsheet.');
    const spreadsheetTags = fetchTagsFromSpreadsheet();
    log('Fetched tags from spreadsheet:', spreadsheetTags);

    // Combine default tags and spreadsheet tags
    userTags = [...DEFAULT_USER_TAGS, ...spreadsheetTags];

    //Ensure Uniqueness
    userTags = Array.from(new Set(userTags));

    log('Combined tags:', userTags);

    userProperties.setProperty("userTags", JSON.stringify(userTags));
  }

  log('getUserTags finished', userTags);
  return userTags;
}

/**
 * Fetches the user tags from the spreadsheet.
 *
 * @returns {string[]} The user tags from the spreadsheet.
 */
function fetchTagsFromSpreadsheet() {
  log('fetchTagsFromSpreadsheet called');
  const spreadsheetId = userProperties.getProperty("spreadsheetId");
  const sheetName = userProperties.getProperty("sheetName");
  const column = userProperties.getProperty("column");

  if (!spreadsheetId || !sheetName || !column) {
    log('Spreadsheet ID, sheet name, or column not configured.');
    return []; // Return an empty array if not configured
  }

  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    log(`Successfully opened spreadsheet with ID: ${spreadsheetId}`);

    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      log(`Sheet "${sheetName}" not found in spreadsheet.`);
      return []; // Return an empty array if sheet not found
    }
    log(`Successfully got sheet: ${sheetName}`);

    const lastRow = sheet.getLastRow();
    log(`Last row in sheet: ${lastRow}`);
    if (lastRow === 0) {
      log('Sheet is empty.');
      return []; // Return an empty array if sheet is empty
    }
    const values = sheet.getRange(`${column}1:${column}${lastRow}`).getValues();
    log(`Values from sheet:`, values);

    // Get unique tags, prepending # if necessary
    const uniqueTags = new Set();
    values.forEach(row => {
      if (row[0]) {
        const tag = row[0].startsWith('#') ? row[0] : `#${row[0]}`;
        uniqueTags.add(tag);
      }
    });

    const tags = Array.from(uniqueTags);
    log(`Fetched tags from spreadsheet: ${tags}`);
    return tags;
  } catch (error) {
    log(`Error fetching tags from spreadsheet: ${error.message}`, error);
    return []; // Return an empty array on error
  }
}
