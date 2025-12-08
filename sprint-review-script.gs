// *** Defines *** 



const BODY_TEXT_PATTERN = "\\{\\{body_text\\}\\}";

const STORY_KEY_PATTERN = "\\{\\{story_key\\}\\}";

const TEMPLATE_PRESENTATION_ID = "1idVd8G-Ec1L2yMF3_fz-eVmpeEYtB1m1I2eWhwW-AGk";

// *** Presentation Information *** 

const sprintNumber = "25-21";

const reviewDate = "2025-11-28";

const teamName = "HW-PR";

const slidesTitle = "Sprint " + sprintNumber + " Review " + reviewDate;

const slidesSubtitle = "HW - Propulsion";

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

  var slidesFileName = reviewDate + "_" + teamName + "_Sprint-" + sprintNumber + "_Review";

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

// *** end of script *** 

