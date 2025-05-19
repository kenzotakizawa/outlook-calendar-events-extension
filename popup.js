// popup.js: Logic for the extension popup UI and communication.

// This function will be injected and executed in the Outlook content script's context
function getOutlookEventsInContentScript() {
  // Call the function defined in content.js
  return getOutlookEvents();
}

document.addEventListener('DOMContentLoaded', () => {
  // Get references to the new overlay elements
  const popupTourOverlay = document.getElementById('popup-tour-overlay');
  const proceedToAiSiteButton = document.getElementById('proceed-to-ai-site');
  const dontShowCheckboxPopup = document.getElementById('dont-show-checkbox-popup');
  const analyzeAiButton = document.getElementById('analyze-ai'); // Get reference to the analyze button
  const eventListElement = document.getElementById('event-list'); // Get reference to event list container
  const eventsTableBody = document.getElementById('events-table-body'); // Get reference to table body
  const totalDurationElement = document.getElementById('total-duration'); // Get reference to total duration element
  const copyCsvButton = document.getElementById('copy-csv'); // Get reference to copy CSV button
  const analysisTextArea = document.getElementById('analysis-prompt'); // Get reference to analysis text area


  // --- Setup Functions ---

  function setupOutlookEventExtraction() {
    // Execute script in the active tab (Outlook) to get event data
    chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
      chrome.scripting.executeScript({
        target: { tabId: tabs[0].id },
        function: getOutlookEventsInContentScript // Function to inject and execute
      }, (results) => {
        if (results && results[0] && results[0].result) {
          displayEvents(results[0].result);
        } else {
          eventListElement.innerHTML = '<p>対象のOutlookカレンダービューではありません、または予定が見つかりませんでした。</p>';
        }
      });
    });
  }

  function setupCopyCsv() {
    copyCsvButton.addEventListener('click', () => {
      // Pass the current events data to the copy function
      // We need to store the events data accessible here, or re-extract/regenerate
      // For simplicity, let's assume events data is available in displayEvents scope or regenerate
      // Regenerating is safer if displayEvents is not always called before copy
      chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
         chrome.scripting.executeScript({
            target: { tabId: tabs[0].id },
            function: getOutlookEventsInContentScript // Re-extract events
          }, (results) => {
            if (results && results[0] && results[0].result) {
               copyEventsAsCsv(results[0].result);
            } else {
               console.warn("Could not re-extract events for CSV copy.");
            }
          });
      });
    });
  }

  function setupAnalyzeAi() {
     // Add event listener for the main analyze button
    analyzeAiButton.addEventListener('click', () => {
       // Re-extract events before copying and showing instructions/opening AI site
       chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
         chrome.scripting.executeScript({
            target: { tabId: tabs[0].id },
            function: getOutlookEventsInContentScript // Re-extract events
          }, (results) => {
            if (results && results[0] && results[0].result) {
               copyAnalysisPromptAndOpenAI(results[0].result);
            } else {
               console.warn("Could not re-extract events for AI analysis.");
            }
          });
      });
    });

    // Add event listener for the proceed button within the overlay
    proceedToAiSiteButton.addEventListener('click', () => {
      // Hide the overlay
      popupTourOverlay.style.display = 'none';
      // Open the AI site
      chrome.tabs.create({ url: 'https://chat.candyhouse.co/' });
    });

    // Load initial checkbox state for the popup overlay
    chrome.storage.sync.get('hideInstructions', (data) => {
      if (data.hideInstructions) {
        dontShowCheckboxPopup.checked = true;
      }
    });

    // Save user preference when popup checkbox state changes
    dontShowCheckboxPopup.addEventListener('change', (event) => {
      chrome.storage.sync.set({ hideInstructions: event.target.checked });
    });
  }


  // --- Core Logic Functions ---

  // Function to display events in the popup
  function displayEvents(events) {
    eventsTableBody.innerHTML = ''; // Clear existing rows

    if (events.length === 0) {
      eventListElement.innerHTML = '<p>画面に表示されている期間内に予定は見つかりませんでした。</p>';
      totalDurationElement.textContent = ''; // Clear total duration
      analysisTextArea.value = "分析する予定データがありません。"; // Clear analysis text area
      return;
    }

    let totalDurationMinutes = 0;

    events.forEach(event => {
      const row = eventsTableBody.insertRow();
      row.insertCell(0).textContent = event.title;
      row.insertCell(1).textContent = event.startTime;
      row.insertCell(2).textContent = event.endTime;
      row.insertCell(3).textContent = event.duration;

      totalDurationMinutes += event.durationMinutes;
    });

    // Display total duration
    const totalHours = Math.floor(totalDurationMinutes / 60);
    const totalMinutes = totalDurationMinutes % 60;
    totalDurationElement.textContent = `合計所要時間: ${totalHours}時間${totalMinutes}分`;

    // Populate analysis prompt text area after events are displayed
    populateAnalysisPrompt(events);
  }

  // Function to copy analysis prompt and show popup overlay or open AI site directly
  function copyAnalysisPromptAndOpenAI(events) {
    if (!analysisTextArea || !analysisTextArea.value) {
      console.warn("No analysis prompt content to copy.");
      return;
    }

    navigator.clipboard.writeText(analysisTextArea.value).then(() => {
      console.log("Analysis prompt and CSV copied to clipboard.");
      // Optionally, provide user feedback
      analyzeAiButton.textContent = 'コピーしました！';
      analyzeAiButton.disabled = true; // Disable button temporarily

      // Check user preference for showing instructions
      chrome.storage.sync.get('hideInstructions', (data) => {
        const hideInstructions = data.hideInstructions === true;

        if (hideInstructions) {
          // If hideInstructions is true, open AI site directly
          chrome.tabs.create({ url: 'https://chat.candyhouse.co/' }, () => {
             // Re-enable button after tab is created
            analyzeAiButton.textContent = '生成AIで分析する';
            analyzeAiButton.disabled = false;
          });
        } else {
          // If hideInstructions is false or not set, show the popup overlay
          popupTourOverlay.style.display = 'flex'; // Show the overlay
           // Re-enable button after showing overlay
          analyzeAiButton.textContent = '生成AIで分析する';
          analyzeAiButton.disabled = false;
        }
      });

    }).catch(err => {
      console.error("Failed to copy analysis prompt:", err);
      // Optionally, provide user feedback about the error
      analyzeAiButton.textContent = 'コピー失敗';
       setTimeout(() => {
        analyzeAiButton.textContent = '生成AIで分析する';
      }, 2000);
    });
  }

  // Function to generate CSV content
  function generateCsvContent(events) {
     if (events.length === 0) {
      return "";
    }

    const header = ["タイトル", "開始時間", "終了時間", "所要時間"];
    const rows = events.map(event => [
      `"${event.title.replace(/"/g, '""')}"`, // Escape double quotes in title
      event.startTime,
      event.endTime,
      event.duration
    ]);

    return [
      header.join(','),
      ...rows.map(row => row.join(','))
    ].join('\n');
  }

   // Function to copy events as CSV to clipboard
  function copyEventsAsCsv(events) {
    const csvContent = generateCsvContent(events);
     if (!csvContent) {
      console.warn("No events to copy.");
      return;
    }

    navigator.clipboard.writeText(csvContent).then(() => {
      console.log("Events copied to clipboard as CSV.");
      // Optionally, provide user feedback (e.g., change button text temporarily)
      copyCsvButton.textContent = 'コピーしました！';
      setTimeout(() => {
        copyCsvButton.textContent = 'CSVとしてコピー';
      }, 2000);
    }).catch(err => {
      console.error("Failed to copy events as CSV:", err);
      // Optionally, provide user feedback about the error
      copyCsvButton.textContent = 'コピー失敗';
       setTimeout(() => {
        copyCsvButton.textContent = 'CSVとしてコピー';
      }, 2000);
    });
  }


  // Function to populate the analysis prompt text area
  function populateAnalysisPrompt(events) {
    if (!analysisTextArea) return;

    if (events.length === 0) {
      analysisTextArea.value = "分析する予定データがありません。";
      return;
    }

    const analysisPrompt = `以下のCSVデータは、Outlookカレンダーから抽出した予定のリストです。\n\nこのデータを使って、業務種別や会社別に費やした時間の工数を分析してください。\n\nCSVデータ:\n`;
    const csvContent = generateCsvContent(events);

    analysisTextArea.value = analysisPrompt + csvContent;
  }

  // --- Initial Setup ---
  setupOutlookEventExtraction();
  setupCopyCsv();
  setupAnalyzeAi();

});
