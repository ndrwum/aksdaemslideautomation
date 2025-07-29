/**
 * Google Apps Script for automating hymn slides creation
 * This script reads from a spreadsheet and creates a presentation with hymn lyrics
 */

// Configuration constants
const CONFIG = {
  TEMPLATE_ID: '<hidden>',
  SPREADSHEET_ID: '<hidden>',
  MIN_FONT_SIZE: 50,
  DEFAULT_FONT_SIZE: 60,
  LINE_SPACING: 2
};

// Column names in the spreadsheet
const COLUMNS = {
  OPENING_HYMN: 'Opening Hymn',
  CLOSING_HYMN: 'Closing Hymn',
  SCRIPTURE_READING: 'Scripture Reading',     // For the Bible passage
  SCRIPTURE_READER: 'Scripture Reader',       // Changed this line - For the person reading
  SERMON_TITLE: 'Sermon Title',
  SPEAKER: 'Speaker',
  SPECIAL_MUSIC: 'Special Music',
  INTERCESSORY_PRAYER: 'Intercessory Prayer',
  CHILDREN_STORY: "Children's Story"
};


// Add the new placeholders
const PLACEHOLDERS = {
  OPENING: '{{opening}}',
  CLOSING: '{{closing}}',
  OPENING_LYRICS: '{{opening_lyrics}}',
  CLOSING_LYRICS: '{{closing_lyrics}}',
  PASSAGE: '{{passage}}',
  VERSE: '{{verse}}',
  SERMON: '{{sermon}}',
  SPEAKER: '{{speaker}}',
  MUSIC: '{{music}}',
  PRAYER: '{{prayer}}',
  READING: '{{reading}}',
  STORY: '{{story}}'
};

/**
 * Main function to create hymn slides
 */
function createHymnsSlides() {
  const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  if (!spreadsheet) {
    Logger.log('Could not find spreadsheet');
    return;
  }

  const targetSheet = findTargetSheet(spreadsheet);
  if (!targetSheet) {
    Logger.log('Could not find target sheet');
    return;
  }

  const upcomingSaturday = getUpcomingSaturday();
  const upcomingSaturdayString = getDateFormatted(upcomingSaturday);
  Logger.log('Looking for date: ' + upcomingSaturdayString);

  const hymnsData = extractHymnsData(targetSheet, upcomingSaturdayString);
  Logger.log('Extracted data:', hymnsData); // Debug log

  if (!hymnsData.openingHymnNumber || !hymnsData.closingHymnNumber) {
    Logger.log('Missing hymn numbers');
    return;
  }

  const hymnDetails = fetchHymnDetails(hymnsData);
  if (!hymnDetails) {
    Logger.log('Could not fetch hymn details');
    return;
  }

  const scriptureContent = fetchScriptureContent(hymnsData.scriptureReading);
  createPresentation(hymnsData, hymnDetails, scriptureContent, upcomingSaturdayString);
}

/**
 * Finds the target sheet containing "Sabbath Schedule 2024"
 */
function findTargetSheet(spreadsheet) {
  return spreadsheet.getSheets().find(sheet => 
    sheet.getName().includes("Sabbath Schedule 2024")
  );
}

/**
 * Gets the date of the upcoming Saturday
 */
function getUpcomingSaturday() {
  const today = new Date();
  const upcomingSaturday = new Date(today);
  upcomingSaturday.setDate(today.getDate() + (6 - today.getDay()) % 7);
  return upcomingSaturday;
}

/**
 * Formats date as MM/dd/yyyy
 */
function getDateFormatted(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "MM/dd/yyyy");
}

/**
 * Extracts hymn numbers and other data from spreadsheet
 */
function extractHymnsData(sheet, targetDate) {
  const dataRange = sheet.getDataRange().getValues();
  const headerRow = dataRange[1]; // Assuming headers are in row 2
  
  const columnIndices = getColumnIndices(headerRow);

  for (let i = 2; i < dataRange.length; i++) {
    const dateCell = dataRange[i][0];
    Logger.log(`Checking row ${dataRange[i]}, date: ${columnIndices.scriptureReading}`); // Debug log
    
    if (dateCell instanceof Date && getDateFormatted(dateCell) === targetDate) {
      const rowData = {
        openingHymnNumber: extractHymnNumber(dataRange[i][columnIndices.openingHymn]),
        closingHymnNumber: extractHymnNumber(dataRange[i][columnIndices.closingHymn]),
        scriptureReading: dataRange[i][columnIndices.scriptureReading],
        sermonTitle: dataRange[i][columnIndices.sermonTitle] || '',
        speaker: dataRange[i][columnIndices.speaker] || '',
        specialMusic: dataRange[i][columnIndices.specialMusic] || '',
        prayer: dataRange[i][columnIndices.prayer] || '',
        reader: dataRange[i][columnIndices.reader] || '',
        story: dataRange[i][columnIndices.story] || ''
      };
      
      Logger.log('Found row data:', rowData); // Debug log
      return rowData;
    }
  }
  Logger.log('No matching date found'); // Debug log
  return {};
}

/**
 * Gets indices of relevant columns
 */
function getColumnIndices(headerRow) {
  const indices = {};
  
  headerRow.forEach((header, index) => {
    const trimmedHeader = header.toString().trim();
    Logger.log(`Checking header: "${trimmedHeader}" at index ${index}`); // Debug log
    
    switch(trimmedHeader) {
      case 'Opening Hymn':
        indices.openingHymn = index;
        break;
      case 'Closing Hymn':
        indices.closingHymn = index;
        break;
      case 'Scripture Reading':
        indices.scriptureReading = index;
        break;
      case 'Sermon Title':
        indices.sermonTitle = index;
        break;
      case 'Speaker':
        indices.speaker = index;
        break;
      case 'Special Music':
        indices.specialMusic = index;
        break;
      case 'Intercessory Prayer':
        indices.prayer = index;
        break;
      case 'Scripture Reader':
        indices.reader = index;
        break;
      case "Children's Story":
        indices.story = index;
        break;
    }
  });

  Logger.log('Found column indices:', indices); // Debug log
  return indices;
}
function updateParticipantsSlides(slides, hymnsData) {
  Logger.log('Updating participants with data:', hymnsData); // Debug log
  
  slides.forEach((slide, index) => {
    slide.getShapes().forEach(shape => {
      try {
        const textRange = shape.getText();
        if (!textRange) return;
        
        const text = textRange.asString();
        Logger.log(`Checking slide ${index}, found text: ${text}`); // Debug log
        
        // Using a map of placeholders to their values for cleaner replacement
        const replacements = {
          '{{speaker}}': hymnsData.speaker,
          '{{music}}': hymnsData.specialMusic,
          '{{prayer}}': hymnsData.prayer,
          '{{reading}}': hymnsData.reader,
          '{{story}}': hymnsData.story
        };

        // Perform all replacements
        Object.entries(replacements).forEach(([placeholder, value]) => {
          if (text.includes(placeholder)) {
            Logger.log(`Replacing ${placeholder} with ${value}`); // Debug log
            textRange.replaceAllText(placeholder, value || '');
          }
        });
      } catch (error) {
        Logger.log(`Error processing shape in slide ${index}: ${error}`); // Debug log
      }
    });
  });
}
/**
 * Extracts hymn number from cell value
 */
function extractHymnNumber(cellValue) {
  if (!cellValue) return null;
  const match = cellValue.match(/(\d+)/);
  return match && match[1] ? match[1].padStart(3, '0') : null;
}

/**
 * Fetches hymn details from the website
 */
function fetchHymnDetails(hymnsData) {
  const { openingHymnNumber, closingHymnNumber } = hymnsData;
  
  try {
    const openingLyricsUrl = `https://sdahymnals.com/Hymnal/${openingHymnNumber}`;
    const closingLyricsUrl = `https://sdahymnals.com/Hymnal/${closingHymnNumber}`;
    
    const openingLyricsHtml = UrlFetchApp.fetch(openingLyricsUrl, { muteHttpExceptions: true }).getContentText();
    const closingLyricsHtml = UrlFetchApp.fetch(closingLyricsUrl, { muteHttpExceptions: true }).getContentText();
    
    if (!openingLyricsHtml || !closingLyricsHtml || 
        openingLyricsHtml.includes("404 Not Found") || 
        closingLyricsHtml.includes("404 Not Found")) {
      return null;
    }

    return {
      opening: {
        title: extractHymnTitle(openingLyricsHtml),
        ...extractHymnVerses(openingLyricsHtml)
      },
      closing: {
        title: extractHymnTitle(closingLyricsHtml),
        ...extractHymnVerses(closingLyricsHtml)
      }
    };
  } catch (error) {
    return null;
  }
}

/**
 * Creates the presentation with all slides
 */
function createPresentation(hymnsData, hymnDetails, scriptureContent, presentationName) {
  const presentation = SlidesApp.openById(
    DriveApp.getFileById(CONFIG.TEMPLATE_ID)
      .makeCopy(presentationName)
      .getId()
  );

  const slides = presentation.getSlides();
  Logger.log(`Created presentation with ${slides.length} slides`); // Debug log
  
  const templateSlides = findTemplateSlides(slides);
  
  if (!areAllTemplateSlidesFound(templateSlides)) {
    Logger.log('Missing template slides'); // Debug log
    return;
  }

  // Update the slides
  updateTitleSlides(templateSlides, hymnDetails);
  createVersesSlides(templateSlides, hymnDetails);
  updateScriptureSlides(slides, scriptureContent);
  updateSermonSlides(slides, hymnsData.sermonTitle);
  updateParticipantsSlides(slides, hymnsData);

  // Clean up template slides
  templateSlides.openingLyrics.remove();
  templateSlides.closingLyrics.remove();

  presentation.saveAndClose();
  Logger.log('Presentation created successfully'); // Debug log
}

/**
 * Finds template slides in the presentation
 */
function findTemplateSlides(slides) {
  const templates = {
    title: null,
    openingLyrics: null,
    closingLyrics: null,
    openingTitle: null,
    closingTitle: null
  };

  slides.forEach(slide => {
    const shapes = slide.getShapes();
    shapes.forEach(shape => {
      try {
        const text = shape.getText()?.asString() || '';
        if (text.includes(PLACEHOLDERS.OPENING)) {
          templates.openingTitle = slide;
          templates.title = templates.title || slide;
        }
        if (text.includes(PLACEHOLDERS.OPENING_LYRICS)) {
          templates.openingLyrics = slide;
        }
        if (text.includes(PLACEHOLDERS.CLOSING_LYRICS)) {
          templates.closingLyrics = slide;
        }
        if (text.includes(PLACEHOLDERS.CLOSING)) {
          templates.closingTitle = slide;
        }
      } catch (error) {
        // Skip shapes that don't have text
      }
    });
  });

  return templates;
}

/**
 * Checks if all required template slides are found
 */
function areAllTemplateSlidesFound(templates) {
  return templates.title && 
         templates.openingLyrics && 
         templates.closingLyrics && 
         templates.openingTitle && 
         templates.closingTitle;
}

/**
 * Updates title slides with hymn information
 */
function updateTitleSlides(templates, hymnDetails) {
  templates.openingTitle.replaceAllText(PLACEHOLDERS.OPENING, hymnDetails.opening.title);
  templates.closingTitle.replaceAllText(PLACEHOLDERS.CLOSING, hymnDetails.closing.title);
}

/**
 * Creates verse slides for both opening and closing hymns
 */
function createVersesSlides(templates, hymnDetails) {
  const createSlidesForHymn = (verses, refrain, templateSlide) => {
    const slidesToCreate = [];

    verses.forEach((verse, index) => {
      slidesToCreate.push({ type: 'verse', text: verse });
      
      if (refrain && index < verses.length - 1) {
        const formattedRefrain = refrain.replace(/Refrain/g, '[Refrain]');
        slidesToCreate.push({ type: 'refrain', text: formattedRefrain });
      }
    });

    // Remove numeric-only last slide if present
    if (slidesToCreate.length > 0) {
      const lastSlideText = slidesToCreate[slidesToCreate.length - 1].text;
      if (/^\d+$/.test(lastSlideText)) {
        slidesToCreate.pop();
      }
    }

    slidesToCreate.reverse().forEach(slideData => {
      const newSlide = templateSlide.duplicate();
      const textShape = findMainTextShape(newSlide);
      if (textShape) {
        adjustFontSizeToFitShape(textShape, slideData.text);
      }
    });
  };

  createSlidesForHymn(hymnDetails.opening.verses, hymnDetails.opening.refrain, templates.openingLyrics);
  createSlidesForHymn(hymnDetails.closing.verses, hymnDetails.closing.refrain, templates.closingLyrics);
}

/**
 * Finds the main text shape in a slide
 */
function findMainTextShape(slide) {
  return slide.getShapes().find(shape => {
    try {
      return shape.getText()?.asString().trim() !== "";
    } catch (error) {
      return false;
    }
  });
}

/**
 * Adjusts font size to fit text within shape
 */
function adjustFontSizeToFitShape(shape, text) {
  const textRange = shape.getText();
  textRange.setText(text);
  
  let fontSize = CONFIG.DEFAULT_FONT_SIZE;
  const shapeHeight = shape.getHeight();
  
  while (calculateTextHeight(text, fontSize) > shapeHeight && fontSize > CONFIG.MIN_FONT_SIZE) {
    fontSize--;
  }
  
  textRange.getTextStyle().setFontSize(fontSize);
}

/**
 * Calculates text height based on font size and content
 */
function calculateTextHeight(text, fontSize) {
  const numLines = text.split('\n').length || 1;
  return fontSize * CONFIG.LINE_SPACING * numLines;
}

/**
 * Fetches scripture content from Bible API
 */
function fetchScriptureContent(scriptureReading) {
  if (!scriptureReading) return { passage: '', verse: '' };

  const verses = scriptureReading.split(',').map(v => v.trim());
  const apiRequests = verses.map(v => ({
    url: `https://www.biblegateway.com/passage/?search=${v}&version=NIV`,
    muteHttpExceptions: true
  }));

  try {
    const responses = UrlFetchApp.fetchAll(apiRequests);
    const passages = responses
      .filter(response => response.getResponseCode() === 200)
      .map(response => {
        const htmlContent = response.getContentText();
        const startIndex = htmlContent.indexOf('std-text');
        if (startIndex === -1) return ''; // Return empty if class is not found

        const snippet = htmlContent.substring(startIndex);
        const startSpan = snippet.indexOf('>') + 1;
        const endSpan = snippet.indexOf('</span>');
        const spanContent = snippet.substring(startSpan, endSpan).trim();

        // Remove all HTML tags using a regular expression
        let textContent = spanContent.replace(/<[^>]+>/g, '').trim();

        // Remove (A), (B), (C) references
        textContent = textContent.replace(/\(\s*[A-Z]\s*\)/g, '').trim();

        return textContent; // Return only the text content
      });

    return {
      passage: passages.join(' '),
      verse: verses.join(', ')
    };
  } catch (error) {
    return { passage: '', verse: '' };
  }
}

/**
 * Updates scripture slides with fetched content
 */
function updateScriptureSlides(slides, scriptureContent) {
  slides.forEach(slide => {
    slide.getShapes().forEach(shape => {
      try {
        const text = shape?.getText()?.asString() || '';
        if (text.includes(PLACEHOLDERS.PASSAGE)) {
          const textRange = shape?.getText();
          textRange?.replaceAllText(PLACEHOLDERS.PASSAGE, scriptureContent.passage);
          adjustFontSizeToFitShape(shape, scriptureContent.passage);
        }
        if (text?.includes(PLACEHOLDERS.VERSE)) {
          shape.getText().replaceAllText(PLACEHOLDERS.VERSE, scriptureContent.verse);
        }
      } catch (error) {
        // Skip shapes that don't have text
      }
    });
  });
}

/**
 * Updates sermon title slides
 */
function updateSermonSlides(slides, sermonTitle) {
  slides.forEach(slide => {
    slide.getShapes().forEach(shape => {
      try {
        const text = shape.getText()?.asString() || '';
        if (text.includes(PLACEHOLDERS.SERMON)) {
          shape.getText().replaceAllText(PLACEHOLDERS.SERMON, sermonTitle || '');
        }
      } catch (error) {
        // Skip shapes that don't have text
      }
    });
  });
}

/**
 * Extracts hymn title from HTML content
 */
function extractHymnTitle(html) {
  const titleMatch = html.match(/<h1[^>]*class\s*=\s*["']?title\s+single-title\s+entry-title["']?[^>]*>(.*?)<\/h1>/);
  return titleMatch ? decodeHtmlEntities(titleMatch[1].trim()) : "Untitled Hymn";
}

/**
 * Decodes HTML entities in text
 */
function decodeHtmlEntities(text) {
  text = text.replace(/&#(\d+);/g, (match, dec) => String.fromCharCode(dec))
             .replace(/&#x([a-fA-F0-9]+);/g, (match, hex) => String.fromCharCode(parseInt(hex, 16)));

  const entities = {
    amp: '&',
    lt: '<',
    gt: '>',
    quot: '"',
    apos: "'",
    nbsp: ' '
  };

  return text.replace(/&([a-zA-Z]+);/g, (match, entity) => entities[entity] || match);
}

/**
 * Extracts hymn verses from HTML content
 */
function extractHymnVerses(html) {
  const tableMatches = html.match(/<table[^>]*>([\s\S]*?)<\/table>/g);
  if (!tableMatches) return { verses: [], refrain: "" };

  const contentBoxHtml = tableMatches[0];
  const verses = [];
  let refrain = "";

  const pTags = contentBoxHtml.match(/<p>([\s\S]*?)<\/p>/g) || [];
  
  pTags.forEach(pTag => {
    const verseHtml = pTag.replace(/<a[^>]*>.*?<\/a>/g, '')
                         .replace(/<br\s*\/?>/g, ' ')
                         .replace(/<\/?[^>]+(>|$)/g, "")
                         .trim();
    
    const decodedVerse = decodeHtmlEntities(verseHtml);
    
    if (decodedVerse) {
      if (decodedVerse.toLowerCase().includes("refrain")) {
        refrain = decodedVerse.replace(/<br\s*\/?>/g, '\n').trim();
      } else {
        verses.push(decodedVerse);
      }
    }
  });

  return { verses, refrain };
}

/**
 * Creates a trigger to run the script automatically
 */
function createTrigger() {
  // Delete existing triggers
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
  
  // Create a new trigger to run every Friday at 6:00 AM
  ScriptApp.newTrigger('createHymnsSlides')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.FRIDAY)
    .atHour(6)
    .create();
}

/**
 * Setup function to initialize the script
 */
function setup() {
  createTrigger();
  // Add any additional setup steps here
}

/**
 * Helper function to validate URLs
 */
function isValidUrl(url) {
  try {
    const response = UrlFetchApp.fetch(url, { 
      muteHttpExceptions: true,
      validateHttpsCertificates: true 
    });
    return response.getResponseCode() === 200;
  } catch (error) {
    return false;
  }
}

/**
 * Helper function to clean up old presentations
 * Optional: Can be used to delete presentations older than X days
 */
function cleanupOldPresentations() {
  const cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - 30); // Keep last 30 days
  
  const files = DriveApp.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    if (file.getMimeType() === MimeType.GOOGLE_SLIDES && 
        file.getDateCreated() < cutoffDate) {
      file.setTrashed(true);
    }
  }
}

/**
 * Error handling wrapper function
 */
function handleError(error) {
  console.error('Error in hymn slides script:', error);
  
  // Optional: Send email notification
  if (error && error.toString().includes('Rate limit exceeded')) {
    MailApp.sendEmail({
      to: Session.getEffectiveUser().getEmail(),
      subject: 'Hymn Slides Script - Rate Limit Error',
      body: 'The script hit API rate limits. Please try again later.'
    });
  }
}

/**
 * Main entry point with error handling
 */
function main() {
  try {
    createHymnsSlides();
  } catch (error) {
    handleError(error);
  }
}
