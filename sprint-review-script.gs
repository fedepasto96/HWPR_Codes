// *** Defines *** 

const BODY_TEXT_PATTERN = "\\{\\{body_text\\}\\}";

const STORY_KEY_PATTERN = "\\{\\{story_key\\}\\}";

const TEMPLATE_PRESENTATION_ID = "1idVd8G-Ec1L2yMF3_fz-eVmpeEYtB1m1I2eWhwW-AGk";

// Column indices (0-based: A=0, B=1, C=2, D=3, E=4, F=5, G=6, etc.)
const COL_ISSUE_TYPE = 0;      // Column A
const COL_STORY_KEY = 2;        // Column C
const COL_STORY_SUMMARY = 3;    // Column D
const COL_STORY_DESCRIPTION = 4; // Column E - adjust if description is in a different column
const COL_STORY_STATUS = 5;     // Column F
const COL_OWNER = 6;            // Column G
const COL_EPIC_LINK = 17;       // Column R

// *** Presentation Information *** 

const teamName = "HW-PR";
const SLACK_WEBHOOK_PROPERTY_KEY = "SLACK_WEBHOOK_URL";

// Calculate sprint information dynamically from spreadsheet
function getSprintInfo() {
  var sheet = SpreadsheetApp.getActive().getSheets()[0];
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();
  
  var maxSprintEndDate = null;
  var maxSprintNumber = null;
  var fallbackMaxSprintNumber = null;
  var fallbackMaxSprintOrder = -1;
  
  // Column Z is index 25 (Sprint.endDate)
  // Column AA is index 26 (Sprint.name)
  var endDateCol = 25; // Column Z
  var sprintNameCol = 26; // Column AA

  function parseDateValue(value) {
    if (!value) return null;
    if (value instanceof Date) return value;
    var dateObj = new Date(String(value));
    if (isNaN(dateObj.getTime())) return null;
    return dateObj;
  }

  function extractSprintNumber(value) {
    if (value === null || value === undefined) return null;
    var text = String(value);
    // Accept formats like "Sprint 26-04", "HW-PR Sprint 26-4", case-insensitive.
    var match = text.match(/sprint\s+(\d{2})-(\d{1,2})/i);
    if (!match) return null;
    var yy = match[1];
    var ww = ("0" + parseInt(match[2], 10)).slice(-2);
    return yy + "-" + ww;
  }

  function sprintOrder(sprintNum) {
    if (!sprintNum) return -1;
    var parts = sprintNum.split("-");
    if (parts.length !== 2) return -1;
    var yearPart = parseInt(parts[0], 10);
    var weekPart = parseInt(parts[1], 10);
    if (isNaN(yearPart) || isNaN(weekPart)) return -1;
    return (yearPart * 100) + weekPart;
  }
  
  // Skip header row (index 0)
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var dateObj = parseDateValue(row[endDateCol]);
    var sprintNum = extractSprintNumber(row[sprintNameCol]);

    // Track max sprint number seen anywhere as fallback.
    if (sprintNum) {
      var order = sprintOrder(sprintNum);
      if (order > fallbackMaxSprintOrder) {
        fallbackMaxSprintOrder = order;
        fallbackMaxSprintNumber = sprintNum;
      }
    }

    // Prefer latest sprint end date, taking sprint number from the same row.
    if (dateObj && (!maxSprintEndDate || dateObj > maxSprintEndDate)) {
      maxSprintEndDate = dateObj;
      if (sprintNum) {
        maxSprintNumber = sprintNum;
      }
    }
  }

  // If latest end-date row had no sprint name, use best fallback sprint number.
  if (!maxSprintNumber && fallbackMaxSprintNumber) {
    maxSprintNumber = fallbackMaxSprintNumber;
  }
  
  // If no data found, fall back to current date calculation
  if (!maxSprintEndDate || !maxSprintNumber) {
    var today = new Date();
    var year = today.getFullYear().toString().substring(2);
    var weekNumber = ("0" + getWeekNumber(today)).slice(-2);
    maxSprintNumber = year + "-" + weekNumber;
    maxSprintEndDate = new Date(today);
    maxSprintEndDate.setDate(maxSprintEndDate.getDate() + 7); // Default to 7 days from now
  }
  
  // Find the last Friday before the sprint end date
  var reviewDate = getLastFridayBefore(maxSprintEndDate);
  
  return {
    sprintNumber: maxSprintNumber,
    reviewDate: formatDate(reviewDate),
    sprintEndDate: maxSprintEndDate
  };
}

function getLastFridayBefore(endDate) {
  var date = new Date(endDate);
  // Go back up to 7 days to find the last Friday
  for (var i = 0; i < 7; i++) {
    if (date.getDay() === 5) { // 5 = Friday
      return date;
    }
    date.setDate(date.getDate() - 1);
  }
  // If no Friday found (shouldn't happen), return the end date
  return endDate;
}

function getWeekNumber(date) {
  var d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  var dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  var yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  return Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
}

function formatDate(date) {
  var year = date.getFullYear();
  var month = date.getMonth() + 1;
  var day = date.getDate();
  var monthStr = month < 10 ? "0" + month : month.toString();
  var dayStr = day < 10 ? "0" + day : day.toString();
  return year + "-" + monthStr + "-" + dayStr;
}

// *** main *** 

function doMagic() {

  console.log("*** Sorting the stories in the active spreadsheet ***");

  sortStoriesSheet();

  console.log("*** Creating slides from the sheet ***");

  fillTemplate();

}

function buttonPressed() {

  console.log(">> Sheet Button Pressed <<");

  doMagic();

}

function fillTemplate() {

  try {
    console.log("=== fillTemplate() called at: " + new Date().toISOString() + " ===");

    // Get current sprint information
    var sprintInfo = getSprintInfo();
    var sprintNumber = sprintInfo.sprintNumber;
    var reviewDate = sprintInfo.reviewDate;
    var slidesTitle = "Sprint " + sprintNumber + " Review " + reviewDate;
    var slidesSubtitle = "HW - Propulsion";

    var slidesFileName = reviewDate + "_" + teamName + "_Sprint-" + sprintNumber + "_Review";

    console.log("Sprint Number: " + sprintNumber);
    console.log("Review Date: " + reviewDate);
    console.log("Slides title: " + slidesTitle);

    console.log("File name: " + slidesFileName);

    // Create a copy of the presentation using DriveApp

    var template = DriveApp.getFileById(TEMPLATE_PRESENTATION_ID);

    var fileName = template.getName();

    console.log("Copy slide deck from template");

    var copy = template.makeCopy();

    copy.setName(slidesFileName);

    var PRESENTATION_ID = copy.getId();
    
    // Get the folder where the template is located (or root if in root)
    var templateParents = template.getParents();
    var outputFolder = templateParents.hasNext() ? templateParents.next() : DriveApp.getRootFolder();
    var outputFolderName = outputFolder.getName();
    var outputFolderUrl = outputFolder.getUrl();

    var presentationUrl = "https://docs.google.com/presentation/d/" + PRESENTATION_ID;

    console.log("PRESENTATION_ID: " + PRESENTATION_ID);
    console.log("OUTPUT LOCATION: " + outputFolderName + " (" + outputFolderUrl + ")");
    console.log("Presentation URL: " + presentationUrl);

    // Open the presentation

    var presentation = SlidesApp.openById(PRESENTATION_ID);

    // extact key slides from new presentation

    var slides = presentation.getSlides();

    if (slides.length === 0) {
      throw new Error("Template presentation has no slides!");
    }

    var templateSlide = slides[slides.length - 1];

    var titleSlide = slides[0];

    // Complete the title slide

    titleSlide.replaceAllText("{{deckTitle}}", slidesTitle);

    titleSlide.replaceAllText("{{deckSubtitle}}", slidesSubtitle);

    // Read data from the spreadsheet - try getRichTextValues first, fallback to getValues

    var sheet = SpreadsheetApp.getActive().getSheets()[0];
    var dataRange = sheet.getDataRange();
    var values = null;
    var useRichText = true;

    try {
      values = dataRange.getRichTextValues();
      console.log("Using getRichTextValues()");
    } catch (e) {
      console.log("getRichTextValues() failed, using getValues() instead. Error: " + e.toString());
      values = dataRange.getValues();
      useRichText = false;
    }

    var storyCount = values.length;

    console.log("Total rows in spreadsheet: " + storyCount);
    console.log("Will process stories (excluding header and duplicates)...");

    // Replace template variables in the presentation with values

    var count = 0;
    var processedStoryKeys = {}; // Track processed story keys to avoid duplicates
    var errorCount = 0;

    for (var rowIndex = 0; rowIndex < values.length; rowIndex++) {
      try {
        var row = values[rowIndex];

        // Helper function to safely get text from a cell
        function safeGetText(cell, defaultValue) {
          if (!cell) return defaultValue || "";
          if (useRichText && cell.getText && typeof cell.getText === 'function') {
            try {
              return cell.getText();
            } catch (e) {
              return defaultValue || "";
            }
          } else if (typeof cell === 'string') {
            return cell;
          } else if (cell !== null && cell !== undefined) {
            return String(cell);
          }
          return defaultValue || "";
        }

        function safeGetLinkUrl(cell) {
          if (!cell) return null;
          if (useRichText && cell.getLinkUrl && typeof cell.getLinkUrl === 'function') {
            try {
              return cell.getLinkUrl();
            } catch (e) {
              return null;
            }
          }
          return null;
        }

        // get values from row for the next slide - with null checks

        var issueType = safeGetText(row[COL_ISSUE_TYPE], "");
        var epicLink = safeGetText(row[COL_EPIC_LINK], "");
        var storyKey = safeGetText(row[COL_STORY_KEY], "");
        var url = safeGetLinkUrl(row[COL_STORY_KEY]);
        var storySummary = safeGetText(row[COL_STORY_SUMMARY], "");

        // Get story description - handle empty cells properly
        var storyDescription = safeGetText(row[COL_STORY_DESCRIPTION], "");
        
        // Debug logging for first non-header row
        if (storyKey != "Key" && count === 0) {
          console.log("DEBUG - Story Key: " + storyKey);
          console.log("DEBUG - Story Summary: " + storySummary);
          var descPreview = storyDescription ? storyDescription.substring(0, Math.min(100, storyDescription.length)) : "(empty)";
          console.log("DEBUG - Story Description (col index " + COL_STORY_DESCRIPTION + "): " + descPreview);
        }

        var storyStatus = safeGetText(row[COL_STORY_STATUS], "").toUpperCase();
        var owner = safeGetText(row[COL_OWNER], "");
        var storyAcceptanceCriteria = "";

        // Cut the story description up to line 8 (only if description exists)
        if (storyDescription && storyDescription.trim() !== "") {
          storyDescription = getTextUpToLine(storyDescription, 8);
        } else {
          storyDescription = ""; // Set to empty string if no description
        }

        // Skip header row and check for duplicates

        if (storyKey != "Key" && storyKey && storyKey.trim() !== "") {
          
          // Check if we've already processed this story key
          if (processedStoryKeys[storyKey]) {
            console.log("Skipping duplicate story key: " + storyKey);
            continue; // Skip this row
          }
          
          // Mark this story key as processed
          processedStoryKeys[storyKey] = true;

          // add one more slide

          presentation.appendSlide(templateSlide);

          // update slides after appending a new slide

          slides = presentation.getSlides();

          var lastSlide = slides[slides.length - 1];

          setStatusColor(lastSlide, storyStatus);

          //setStorySummaryWithLink(lastSlide, storyKey, url, storySummary);

          setStorKeyWithLink(lastSlide, storyKey, url);

          lastSlide.replaceAllText("{{epic}}", epicLink);

          lastSlide.replaceAllText("{{story_title}}", storySummary);

          lastSlide.replaceAllText("{{story_summary}}", storyDescription);

          lastSlide.replaceAllText("{{story_ac}}", storyAcceptanceCriteria); // TODO remove "# Default checklist" and newline

          //lastSlide.replaceAllText("{{story_status}}", storyStatus);

          lastSlide.replaceAllText("{{owner}}", owner);

          // TODO add subtasks if any to body text
          
          count++;

        }
      } catch (rowError) {
        errorCount++;
        console.log("Error processing row " + (rowIndex + 1) + ": " + rowError.toString());
        // Continue processing other rows
      }
    }

    // delete the template slide

    //templateSlide.remove();

    console.log("Slide deck creation completed. Total slides created: " + count);
    if (errorCount > 0) {
      console.log("WARNING: " + errorCount + " rows had errors and were skipped.");
    }

    // Send Slack notification (if webhook is configured)
    sendSlackSlidesNotification(
      presentationUrl,
      "Sprint review slides for the team",
      slidesFileName
    );

    console.log("=== fillTemplate() finished ===");
    
  } catch (error) {
    console.log("FATAL ERROR in fillTemplate(): " + error.toString());
    console.log("Stack trace: " + error.stack);
    throw error;
  }

}

function sortStoriesSheet() {

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    var sheet = ss.getSheets()[0];

    var lastRow = sheet.getLastRow();
    
    if (lastRow < 2) {
      console.log("No data rows to sort (only header or empty sheet)");
      return;
    }

    // Use actual data range instead of hardcoded range
    var range = sheet.getRange("A2:Z" + lastRow);

    // Sorts by Status (column F, index 6)

    range.sort(6);

    // Sorts by Epic (column C, index 2)

    range.sort(2);
    
    console.log("Sheet sorted successfully");
  } catch (error) {
    console.log("Error sorting sheet: " + error.toString());
    // Don't throw - sorting is not critical, continue with slide generation
  }
}

// *** Helper functions *** 

function setStorKeyWithLink(slide, storyKey, storyUrl) {

  var sha = getShapeWithText(slide, STORY_KEY_PATTERN);

  if (sha) {
    setShapeUrl(sha, STORY_KEY_PATTERN, storyUrl, storyKey, "");
  } else {
    console.log("WARNING: Could not find shape with pattern " + STORY_KEY_PATTERN + " for story " + storyKey);
  }

}

function setStorySummaryWithLink(slide, storyKey, storyUrl, storySummary) {

  var sha = getShapeWithText(slide, BODY_TEXT_PATTERN);

  if (sha) {
    setShapeUrl(sha, BODY_TEXT_PATTERN, storyUrl, storyKey, " " + storySummary);
  } else {
    console.log("WARNING: Could not find shape with pattern " + BODY_TEXT_PATTERN + " for story " + storyKey);
  }

}

function getShapeWithText(slide, text) {

  var returnShape = null;

  slide.getShapes().forEach(shape => {

    shape.getText().find(text)

      .forEach((v) => {

        returnShape = shape;

      })

  })

  return returnShape;

}

function setShapeUrl(shape, pattern, url, urlText, postUrlText) {

  if (!shape) {
    console.log("WARNING: setShapeUrl called with null shape");
    return;
  }

  if (url) {
    try {
      shape.getText().find(pattern)
        .forEach((v) => {
          const style = v.setText(urlText).getTextStyle();
          style.setLinkUrl(url);
          v.appendText(postUrlText);
        });
    } catch (e) {
      console.log("WARNING: Error setting URL for pattern " + pattern + ": " + e.toString());
    }
  }

}

function setStatusColor(slide, statusText) {

  //Get all shapes in the current slide

  slide.getShapes().forEach(shape => {

    var text = shape.getText();

    //Search for the string "{{story_status}}"

    var str = text.find("\\{\\{story_status\\}\\}");

    str.forEach(s => {

      var color = getStatusColor(statusText);

      s.getTextStyle().setForegroundColor(color);

    });

  })

}

function getStatusColor(text) {

  if (text == "CLOSED" || text == "WON'T IMPLEMENT" || text == "REOPENED") {

    return "#34A853"

  } else if (text == "IN PROGRESS" || text == "IN REVIEW" || text == "BLOCKED") {

    return "#F1C232"

  } else if (text == "NEW") {

    return "#C00000"

  }

  return "#454545"

}

// text cut up to line 8

function getTextUpToLine(text, lineNumber) {

  var lines = text.split('\n'); 

  if (lines.length > lineNumber) {

    var slicedLines = lines.slice(0, lineNumber);

    return slicedLines.join('\n') + '...';

  } else {

    return text

  }

}

function sendSlackSlidesNotification(slidesUrl, messageText, deckName) {
  try {
    var webhookUrl = PropertiesService.getScriptProperties().getProperty(SLACK_WEBHOOK_PROPERTY_KEY);
    if (!webhookUrl) {
      console.log("Slack notification skipped: missing Script Property " + SLACK_WEBHOOK_PROPERTY_KEY);
      return;
    }

    var payload = {
      text: messageText + "\n" + slidesUrl,
      unfurl_links: true,
      attachments: [
        {
          color: "#2eb886",
          title: deckName,
          title_link: slidesUrl,
          text: "Sprint review slides deck is ready."
        }
      ]
    };

    var options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    var response = UrlFetchApp.fetch(webhookUrl, options);
    var statusCode = response.getResponseCode();
    if (statusCode >= 200 && statusCode < 300) {
      console.log("Slack notification sent successfully.");
    } else {
      console.log("Slack notification failed. HTTP " + statusCode + ": " + response.getContentText());
    }
  } catch (e) {
    console.log("Slack notification error: " + e.toString());
  }
}

// Run once manually to configure the incoming webhook URL.
function setSlackWebhookUrl(url) {
  if (!url || typeof url !== "string" || url.indexOf("https://hooks.slack.com/services/") !== 0) {
    throw new Error("Invalid Slack webhook URL.");
  }
  PropertiesService.getScriptProperties().setProperty(SLACK_WEBHOOK_PROPERTY_KEY, url);
  console.log("Slack webhook URL saved in Script Properties.");
}

// *** Automated Triggers ***

// Time-based trigger: Run every Monday at 11:00 AM, but only if it's the Monday before sprint end
function runOnMondayIfBeforeSprintEnd() {
  var today = new Date();
  var dayOfWeek = today.getDay(); // 0 = Sunday, 1 = Monday, etc.
  
  // Only run on Mondays
  if (dayOfWeek !== 1) {
    console.log("Not a Monday. Skipping.");
    return;
  }
  
  // Get sprint info to check if we're in the week before sprint ends
  var sprintInfo = getSprintInfo();
  var sprintEndDate = sprintInfo.sprintEndDate;
  
  // Check if today is within 7 days before the sprint end date
  var daysUntilSprintEnd = Math.floor((sprintEndDate - today) / (1000 * 60 * 60 * 24));
  
  if (daysUntilSprintEnd >= 0 && daysUntilSprintEnd <= 7) {
    console.log("Monday before sprint end detected. Days until sprint end: " + daysUntilSprintEnd);
    console.log("Running script...");
    doMagic();
  } else {
    console.log("Not the Monday before sprint end. Days until sprint end: " + daysUntilSprintEnd + ". Skipping.");
  }
}

// Function to set up the time-based trigger (run this once manually)
function setupMondayTrigger() {
  // Delete any existing triggers with the same function name
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'runOnMondayIfBeforeSprintEnd') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Create a new trigger for every Monday at 11:00 AM
  ScriptApp.newTrigger('runOnMondayIfBeforeSprintEnd')
    .timeBased()
    .everyWeeks(1)
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(11) // 11:00 AM
    .create();
  
  console.log("Monday trigger set up successfully! Script will run every Monday at 11:00 AM.");
}

// *** end of script *** 

