// *** Defines *** 

const BODY_TEXT_PATTERN = "\\{\\{body_text\\}\\}";

const STORY_KEY_PATTERN = "\\{\\{story_key\\}\\}";

const TEMPLATE_PRESENTATION_ID = "1idVd8G-Ec1L2yMF3_fz-eVmpeEYtB1m1I2eWhwW-AGk";

// *** Presentation Information *** 

const teamName = "HW-PR";

// Calculate sprint information dynamically from spreadsheet
function getSprintInfo() {
  var sheet = SpreadsheetApp.getActive().getSheets()[0];
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();
  
  var maxSprintEndDate = null;
  var maxSprintNumber = null;
  
  // Column Z is index 25 (Sprint.endDate)
  // Column AA is index 26 (Sprint.name)
  var endDateCol = 25; // Column Z
  var sprintNameCol = 26; // Column AA
  
  // Skip header row (index 0)
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    
    // Get sprint end date from column Z
    var endDateValue = row[endDateCol];
    var dateObj = null;
    
    if (endDateValue) {
      if (endDateValue instanceof Date) {
        dateObj = endDateValue;
      } else if (typeof endDateValue === 'string') {
        // Try to parse ISO date string (e.g., "2025-12-15T07:29:18.000Z")
        dateObj = new Date(endDateValue);
        if (isNaN(dateObj.getTime())) {
          dateObj = null; // Invalid date
        }
      }
      
      if (dateObj && (!maxSprintEndDate || dateObj > maxSprintEndDate)) {
        maxSprintEndDate = dateObj;
      }
    }
    
    // Get sprint name from column AA and extract sprint number
    var sprintName = row[sprintNameCol];
    if (sprintName && typeof sprintName === 'string') {
      // Extract sprint number from format "HW-PR Sprint 25-22"
      var match = sprintName.match(/Sprint\s+(\d{2}-\d{2})/);
      if (match && match[1]) {
        var sprintNum = match[1];
        // Compare sprint numbers (e.g., "25-22" vs "25-21")
        if (!maxSprintNumber || sprintNum > maxSprintNumber) {
          maxSprintNumber = sprintNum;
        }
      }
    }
  }
  
  // If no data found, fall back to current date calculation
  if (!maxSprintEndDate || !maxSprintNumber) {
    var today = new Date();
    var year = today.getFullYear().toString().substring(2);
    var weekNumber = getWeekNumber(today);
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

  console.log("PRESENTATION_ID: " + PRESENTATION_ID);

  // Open the presentation

  var presentation = SlidesApp.openById(PRESENTATION_ID);

  // extact key slides from new presentation

  var slides = presentation.getSlides();

  var templateSlide = slides[slides.length - 1];

  var titleSlide = slides[0];

  // Complete the title slide

  titleSlide.replaceAllText("{{deckTitle}}", slidesTitle);

  titleSlide.replaceAllText("{{deckSubtitle}}", slidesSubtitle);

  // Read data from the spreadsheet

  var values = SpreadsheetApp.getActive().getSheets()[0].getDataRange().getRichTextValues();

  var storyCount = values.length;

  console.log("Story count: " + storyCount);

  // Replace template variables in the presentation with values

  var count = 0;

  values.forEach(function (row) {

    // get values from row for the next slide

    var issueType = row[0].getText();

    var epicLink = row[17].getText();

    var storyKey = row[2].getText();

    var url = row[2].getLinkUrl();

    var storySummary = row[3].getText();

    var storyDescription = row[4].getText();

    var storyStatus = row[5].getText().toString().toUpperCase();

    var owner = row[6].getText();

    var storyAcceptanceCriteria = "";

    // Cut the story description up to line 8

    storyDescription = getTextUpToLine(storyDescription, 8);

    // Skip header row

    if (storyKey != "Key") {

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

    }

  });

  // delete the template slide

  //templateSlide.remove();

  console.log("Slide deck creation completed");

}

function sortStoriesSheet() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var sheet = ss.getSheets()[0];

  var range = sheet.getRange("A2:Z300");

  // Sorts by Status

  range.sort(6);

  // Sorts by Epic

  range.sort(2);

}

// *** Helper functions *** 

function setStorKeyWithLink(slide, storyKey, storyUrl) {

  sha = getShapeWithText(slide, STORY_KEY_PATTERN);

  setShapeUrl(sha, STORY_KEY_PATTERN, storyUrl, storyKey, "")

}

function setStorySummaryWithLink(slide, storyKey, storyUrl, storySummary) {

  sha = getShapeWithText(slide, BODY_TEXT_PATTERN);

  setShapeUrl(sha, BODY_TEXT_PATTERN, storyUrl, storyKey, " " + storySummary)

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

  if (url) {

    shape.getText().find(pattern)

      .forEach((v) => {

        const style = v.setText(urlText).getTextStyle();

        style.setLinkUrl(url);

        v.appendText(postUrlText);

      })

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

