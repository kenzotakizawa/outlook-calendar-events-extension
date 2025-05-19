// content.js: Logic for extracting event information from Outlook Web.

console.log("Content script loaded.");

// Function to check if the current view is a supported calendar view
function isCalendarView() {
  const supportedViews = [
    'Surface_WorkWeek', // 稼働日ビュー
    'Surface_Week',     // 週ビュー
    'Surface_Month'     // 月ビュー
  ];
  // Check for elements that identify the calendar view
  return supportedViews.some(view => document.querySelector(`[data-app-section="${view}"]`));
}

// Function to extract event information
function extractEvents() {
  if (!isCalendarView()) {
    console.log("Not on a supported Outlook calendar view.");
    return [];
  }

  console.log("On a supported calendar view. Extracting events...");
  const events = [];
  // Find all elements that are likely event buttons
  const eventElements = document.querySelectorAll('[role="button"]');

  eventElements.forEach(element => {
    const ariaLabel = element.getAttribute('aria-label');
    if (ariaLabel) {
      // Further process elements with aria-label to confirm they are events
      // Basic check for time format in aria-label
      // Check if the aria-label matches the expected event format
      const eventMatch = ariaLabel.match(/(.+?)、(\d{1,2}:\d{2}) から (\d{1,2}:\d{2})、(\d{4} 年 \d{1,2} 月 \d{1,2} 日) \(.*?\)/);

      if (eventMatch) {
        console.log("Found event element:", ariaLabel);
        const title = eventMatch[1].trim();
        const startTimeStr = eventMatch[2];
        const endTimeStr = eventMatch[3];
        const dateStr = eventMatch[4];

        // Convert date and time strings to Date objects
        // Robust method: Construct date string in YYYY/MM/DD HH:MM format
        const dateParts = dateStr.match(/(\d{4}) 年 (\d{1,2}) 月 (\d{1,2}) 日/);
        if (dateParts) {
          const year = dateParts[1];
          const month = dateParts[2]; // Month is 1-based from regex
          const day = dateParts[3];

          const startDateTimeStr = `${year}/${month}/${day} ${startTimeStr}`;
          const endDateTimeStr = `${year}/${month}/${day} ${endTimeStr}`;

          const startDate = new Date(startDateTimeStr);
          const endDate = new Date(endDateTimeStr);

          // Calculate duration in minutes
          const durationMinutes = (endDate.getTime() - startDate.getTime()) / (1000 * 60);

          // Format duration as "○時間△分"
          const hours = Math.floor(durationMinutes / 60);
          const minutes = durationMinutes % 60;
          const durationFormatted = `${hours}時間${minutes}分`;

          events.push({
            title: title,
            startTime: startTimeStr,
            endTime: endTimeStr,
            date: dateStr,
            duration: durationFormatted,
            durationMinutes: durationMinutes // Keep minutes for potential total calculation
          });
        } else {
          console.warn("Could not parse date from aria-label:", ariaLabel);
        }
      }
    }
  });

  return events;
}

// Function to be called by chrome.scripting.executeScript
function getOutlookEvents() {
  return extractEvents();
}

// Listen for messages from the popup script (keeping for now, but will likely remove)
// chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
//   if (request.action === "extractEvents") {
//     const extractedEvents = extractEvents();
//     sendResponse({ events: extractedEvents });
//   }
// });
