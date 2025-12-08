/**
 * Google Apps Script for automating hymn slides creation
 * This script reads from a spreadsheet and creates a presentation with hymn lyrics
 * Now includes Gmail integration for praise/worship songs
 * Updated to include cleaning announcements
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
  SCRIPTURE_READING: 'Scripture Reading',
  SCRIPTURE_READER: 'Scripture Reader',
  SERMON_TITLE: 'Sermon Title',
  SPEAKER: 'Speaker',
  SPECIAL_MUSIC: 'Special Music',
  INTERCESSORY_PRAYER: 'Intercessory Prayer',
  CHILDREN_STORY: "Children's Story",
  CLEANING: 'Cleaning'
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
  STORY: '{{story}}',
  PRAISE_SONG: '{{praise_song}}',
  PRAISE_LYRICS: '{{praise_lyrics}}',
  TODAY_ACCOUNCEMENT: '{{today_accouncement}}',
  UPCOMING_ACCOUNCEMENT: '{{upcoming_accouncement}}'
};

/**
 * Main function to create hymn slides (updated to include praise songs and cleaning announcements)
 */
function createHymnsSlides() {
  try {
    Logger.log('Starting createHymnsSlides function');
    
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
    Logger.log('Extracted data:', hymnsData);

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
    
    // Search for praise/worship lyrics in Gmail
    const praiseData = searchGmailForPraiseLyrics();
    
    // Get cleaning announcements
    const cleaningAnnouncements = getCleaningAnnouncements(targetSheet, upcomingSaturday);
    
    createPresentation(hymnsData, hymnDetails, scriptureContent, upcomingSaturdayString, praiseData, cleaningAnnouncements);
    
  } catch (error) {
    Logger.log('Error in createHymnsSlides: ' + error.toString());
    Logger.log('Error stack: ' + error.stack);
  }
}

/**
 * Gets cleaning announcements for current and next Saturday
 */
function getCleaningAnnouncements(sheet, upcomingSaturday) {
  try {
    const dataRange = sheet.getDataRange().getValues();
    if (dataRange.length < 2) {
      Logger.log('Sheet has insufficient data for cleaning announcements');
      return { today: '', upcoming: '' };
    }
    
    const headerRow = dataRange[1]; // Assuming headers are in row 2
    const columnIndices = getColumnIndices(headerRow);
    
    if (columnIndices.cleaning === undefined) {
      Logger.log('Cleaning column not found');
      return { today: '', upcoming: '' };
    }
    
    const upcomingSaturdayString = getDateFormatted(upcomingSaturday);
    
    // Calculate next Saturday (7 days after upcoming Saturday)
    const nextSaturday = new Date(upcomingSaturday);
    nextSaturday.setDate(nextSaturday.getDate() + 7);
    const nextSaturdayString = getDateFormatted(nextSaturday);
    
    let todayAnnouncement = '';
    let upcomingAnnouncement = '';
    
    // Find the rows for both dates
    for (let i = 2; i < dataRange.length; i++) {
      const dateCell = dataRange[i][0];
      
      if (dateCell instanceof Date) {
        const dateString = getDateFormatted(dateCell);
        
        if (dateString === upcomingSaturdayString) {
          const cleaningContent = dataRange[i][columnIndices.cleaning] || '';
          if (cleaningContent) {
            todayAnnouncement = 'Dishwashers+table cleaners: ' + cleaningContent;
          }
          Logger.log('Found cleaning for upcoming Saturday: ' + cleaningContent);
        }
        
        if (dateString === nextSaturdayString) {
          const cleaningContent = dataRange[i][columnIndices.cleaning] || '';
          if (cleaningContent) {
            upcomingAnnouncement = 'Dishwashers+table cleaners: ' + cleaningContent;
          }
          Logger.log('Found cleaning for next Saturday: ' + cleaningContent);
        }
      }
    }
    
    return {
      today: todayAnnouncement,
      upcoming: upcomingAnnouncement
    };
    
  } catch (error) {
    Logger.log('Error getting cleaning announcements: ' + error.toString());
    return { today: '', upcoming: '' };
  }
}

/**
 * Creates the presentation with all slides (updated to include praise songs and cleaning announcements)
 */
function createPresentation(hymnsData, hymnDetails, scriptureContent, presentationName, praiseData, cleaningAnnouncements) {
  try {
    Logger.log('Creating presentation with name: ' + presentationName);
    
    const presentation = SlidesApp.openById(
      DriveApp.getFileById(CONFIG.TEMPLATE_ID)
        .makeCopy(presentationName)
        .getId()
    );

    const slides = presentation.getSlides();
    Logger.log(`Created presentation with ${slides.length} slides`);
    
    const templateSlides = findTemplateSlides(slides);
    
    if (!areAllTemplateSlidesFound(templateSlides)) {
      Logger.log('Missing template slides');
      return;
    }

    // Update the slides in proper order
    updateTitleSlides(templateSlides, hymnDetails);
    createVersesSlides(templateSlides, hymnDetails);
    updateScriptureSlides(slides, scriptureContent);
    updateSermonSlides(slides, hymnsData.sermonTitle);
    updateParticipantsSlides(slides, hymnsData);
    
    // Update cleaning announcements
    updateCleaningAnnouncementSlides(slides, cleaningAnnouncements);
    
    // Add praise song slides if email was found
    if (praiseData) {
      updatePraiseSongSlides(presentation, praiseData);
    }

    // Clean up template slides AFTER all processing
    if (templateSlides.openingLyrics) {
      templateSlides.openingLyrics.remove();
    }
    if (templateSlides.closingLyrics) {
      templateSlides.closingLyrics.remove();
    }

    presentation.saveAndClose();
    Logger.log('Presentation created successfully');
    
  } catch (error) {
    Logger.log('Error in createPresentation: ' + error.toString());
    Logger.log('Error stack: ' + error.stack);
  }
}

/**
 * Updates slides with cleaning announcement placeholders
 */
function updateCleaningAnnouncementSlides(slides, cleaningAnnouncements) {
  try {
    Logger.log('Updating cleaning announcements:', cleaningAnnouncements);
    
    slides.forEach((slide, index) => {
      const shapes = slide.getShapes();
      shapes.forEach(shape => {
        try {
          const textRange = shape.getText();
          if (!textRange) return;
          
          const text = textRange.asString();
          
          // Replace today's announcement (with the typo as in user's template)
          if (text.includes(PLACEHOLDERS.TODAY_ACCOUNCEMENT)) {
            textRange.replaceAllText(PLACEHOLDERS.TODAY_ACCOUNCEMENT, cleaningAnnouncements.today);
            Logger.log(`Replaced today's announcement on slide ${index + 1}`);
          }
          
          // Replace upcoming announcement (with the typo as in user's template)
          if (text.includes(PLACEHOLDERS.UPCOMING_ACCOUNCEMENT)) {
            textRange.replaceAllText(PLACEHOLDERS.UPCOMING_ACCOUNCEMENT, cleaningAnnouncements.upcoming);
            Logger.log(`Replaced upcoming announcement on slide ${index + 1}`);
          }
        } catch (error) {
          Logger.log(`Error processing shape in slide ${index}: ${error}`);
        }
      });
    });
    
  } catch (error) {
    Logger.log('Error updating cleaning announcement slides: ' + error.toString());
  }
}

/**
 * Finds the target sheet containing "Sabbath Schedule"
 */
function findTargetSheet(spreadsheet) {
  try {
    const sheets = spreadsheet.getSheets();
    for (let sheet of sheets) {
      if (sheet.getName().includes("Sabbath Schedule")) {
        Logger.log('Found target sheet: ' + sheet.getName());
        return sheet;
      }
    }
    Logger.log('No sheet found containing "Sabbath Schedule"');
    return null;
  } catch (error) {
    Logger.log('Error finding target sheet: ' + error.toString());
    return null;
  }
}

/**
 * Gets the date of the upcoming Saturday
 */
function getUpcomingSaturday() {
  const today = new Date();
  const upcomingSaturday = new Date(today);
  const daysUntilSaturday = (6 - today.getDay()) % 7;
  upcomingSaturday.setDate(today.getDate() + (daysUntilSaturday === 0 ? 7 : daysUntilSaturday));
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
  try {
    const dataRange = sheet.getDataRange().getValues();
    if (dataRange.length < 2) {
      Logger.log('Sheet has insufficient data');
      return {};
    }
    
    const headerRow = dataRange[1]; // Assuming headers are in row 2
    const columnIndices = getColumnIndices(headerRow);

    for (let i = 2; i < dataRange.length; i++) {
      const dateCell = dataRange[i][0];
      Logger.log(`Checking row ${i}, date: ${dateCell}`);
      
      if (dateCell instanceof Date && getDateFormatted(dateCell) === targetDate) {
        const rowData = {
          openingHymnNumber: extractHymnNumber(dataRange[i][columnIndices.openingHymn]),
          closingHymnNumber: extractHymnNumber(dataRange[i][columnIndices.closingHymn]),
          scriptureReading: dataRange[i][columnIndices.scriptureReading] || '',
          sermonTitle: dataRange[i][columnIndices.sermonTitle] || '',
          speaker: dataRange[i][columnIndices.speaker] || '',
          specialMusic: dataRange[i][columnIndices.specialMusic] || '',
          prayer: dataRange[i][columnIndices.prayer] || '',
          reader: dataRange[i][columnIndices.reader] || '',
          story: dataRange[i][columnIndices.story] || ''
        };
        
        Logger.log('Found row data:', rowData);
        return rowData;
      }
    }
    Logger.log('No matching date found');
    return {};
    
  } catch (error) {
    Logger.log('Error extracting hymns data: ' + error.toString());
    return {};
  }
}

/**
 * Gets indices of relevant columns
 */
function getColumnIndices(headerRow) {
  const indices = {};
  
  headerRow.forEach((header, index) => {
    const trimmedHeader = header.toString().trim();
    Logger.log(`Checking header: "${trimmedHeader}" at index ${index}`);
    
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
      case 'Cleaning':
        indices.cleaning = index;
        break;
    }
  });

  Logger.log('Found column indices:', indices);
  return indices;
}

function updateParticipantsSlides(slides, hymnsData) {
  try {
    Logger.log('Updating participants with data:', hymnsData);
    
    slides.forEach((slide, index) => {
      const shapes = slide.getShapes();
      shapes.forEach(shape => {
        try {
          const textRange = shape.getText();
          if (!textRange) return;
          
          const text = textRange.asString();
          
          // Using a map of placeholders to their values for cleaner replacement
          const replacements = {
            '{{speaker}}': hymnsData.speaker || '',
            '{{music}}': hymnsData.specialMusic || '',
            '{{prayer}}': hymnsData.prayer || '',
            '{{reading}}': hymnsData.reader || '',
            '{{story}}': hymnsData.story || ''
          };

          // Perform all replacements
          Object.entries(replacements).forEach(([placeholder, value]) => {
            if (text.includes(placeholder)) {
              Logger.log(`Replacing ${placeholder} with ${value}`);
              textRange.replaceAllText(placeholder, value);
            }
          });
        } catch (error) {
          Logger.log(`Error processing shape in slide ${index}: ${error}`);
        }
      });
    });
    
  } catch (error) {
    Logger.log('Error updating participants slides: ' + error.toString());
  }
}

/**
 * Extracts hymn number from cell value
 */
function extractHymnNumber(cellValue) {
  if (!cellValue) return null;
  const match = cellValue.toString().match(/(\d+)/);
  return match && match[1] ? match[1].padStart(3, '0') : null;
}

/**
 * Fetches hymn details from the website
 */
function fetchHymnDetails(hymnsData) {
  const { openingHymnNumber, closingHymnNumber } = hymnsData;
  
  try {
    Logger.log('Fetching hymn details for opening: ' + openingHymnNumber + ', closing: ' + closingHymnNumber);
    
    const openingLyricsUrl = `https://sdahymnals.com/Hymnal/${openingHymnNumber}`;
    const closingLyricsUrl = `https://sdahymnals.com/Hymnal/${closingHymnNumber}`;
    
    const openingResponse = UrlFetchApp.fetch(openingLyricsUrl, { muteHttpExceptions: true });
    const closingResponse = UrlFetchApp.fetch(closingLyricsUrl, { muteHttpExceptions: true });
    
    if (openingResponse.getResponseCode() !== 200 || closingResponse.getResponseCode() !== 200) {
      Logger.log('HTTP error fetching hymns');
      return null;
    }
    
    const openingLyricsHtml = openingResponse.getContentText();
    const closingLyricsHtml = closingResponse.getContentText();
    
    if (!openingLyricsHtml || !closingLyricsHtml || 
        openingLyricsHtml.includes("404 Not Found") || 
        closingLyricsHtml.includes("404 Not Found")) {
      Logger.log('404 error or empty content');
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
    Logger.log('Error fetching hymn details: ' + error.toString());
    return null;
  }
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
          if (!templates.title) templates.title = slide;
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

  Logger.log('Found template slides:', {
    title: templates.title ? 'found' : 'missing',
    openingLyrics: templates.openingLyrics ? 'found' : 'missing',
    closingLyrics: templates.closingLyrics ? 'found' : 'missing',
    openingTitle: templates.openingTitle ? 'found' : 'missing',
    closingTitle: templates.closingTitle ? 'found' : 'missing'
  });

  return templates;
}

/**
 * Checks if all required template slides are found
 */
function areAllTemplateSlidesFound(templates) {
  const required = templates.title && 
         templates.openingLyrics && 
         templates.closingLyrics && 
         templates.openingTitle && 
         templates.closingTitle;
  
  if (!required) {
    Logger.log('Missing required template slides');
  }
  
  return required;
}

/**
 * Updates title slides with hymn information
 */
function updateTitleSlides(templates, hymnDetails) {
  try {
    if (templates.openingTitle && hymnDetails.opening) {
      templates.openingTitle.replaceAllText(PLACEHOLDERS.OPENING, hymnDetails.opening.title);
      Logger.log('Updated opening title: ' + hymnDetails.opening.title);
    }
    if (templates.closingTitle && hymnDetails.closing) {
      templates.closingTitle.replaceAllText(PLACEHOLDERS.CLOSING, hymnDetails.closing.title);
      Logger.log('Updated closing title: ' + hymnDetails.closing.title);
    }
  } catch (error) {
    Logger.log('Error updating title slides: ' + error.toString());
  }
}

/**
 * Creates verse slides for both opening and closing hymns
 */
function createVersesSlides(templates, hymnDetails) {
  try {
    const createSlidesForHymn = (verses, refrain, templateSlide, hymnType) => {
      if (!verses || !Array.isArray(verses) || !templateSlide) {
        Logger.log(`Invalid data for ${hymnType} hymn verses`);
        return;
      }
      
      const slidesToCreate = [];

      verses.forEach((verse, index) => {
        if (verse && verse.trim()) {
          slidesToCreate.push({ type: 'verse', text: verse });
          
          if (refrain && index < verses.length - 1) {
            const formattedRefrain = refrain.replace(/Refrain/g, '[Refrain]');
            slidesToCreate.push({ type: 'refrain', text: formattedRefrain });
          }
        }
      });

      // Remove numeric-only last slide if present
      if (slidesToCreate.length > 0) {
        const lastSlideText = slidesToCreate[slidesToCreate.length - 1].text;
        if (/^\d+$/.test(lastSlideText.trim())) {
          slidesToCreate.pop();
        }
      }

      Logger.log(`Creating ${slidesToCreate.length} slides for ${hymnType} hymn`);

      slidesToCreate.reverse().forEach((slideData, index) => {
        try {
          const newSlide = templateSlide.duplicate();
          const textShape = findMainTextShape(newSlide);
          if (textShape) {
            adjustFontSizeToFitShape(textShape, slideData.text);
          }
        } catch (error) {
          Logger.log(`Error creating slide ${index} for ${hymnType}: ${error}`);
        }
      });
    };

    if (hymnDetails.opening) {
      createSlidesForHymn(hymnDetails.opening.verses, hymnDetails.opening.refrain, templates.openingLyrics, 'opening');
    }
    if (hymnDetails.closing) {
      createSlidesForHymn(hymnDetails.closing.verses, hymnDetails.closing.refrain, templates.closingLyrics, 'closing');
    }
    
  } catch (error) {
    Logger.log('Error creating verses slides: ' + error.toString());
  }
}

/**
 * Finds the main text shape in a slide
 */
function findMainTextShape(slide) {
  try {
    const shapes = slide.getShapes();
    for (let shape of shapes) {
      try {
        const textRange = shape.getText();
        if (textRange && textRange.asString().trim() !== "") {
          return shape;
        }
      } catch (error) {
        // Continue to next shape
      }
    }
    return shapes.length > 0 ? shapes[0] : null;
  } catch (error) {
    Logger.log('Error finding main text shape: ' + error.toString());
    return null;
  }
}

/**
 * Adjusts font size to fit text within shape
 */
function adjustFontSizeToFitShape(shape, text) {
  try {
    const textRange = shape.getText();
    if (!textRange) return;
    
    // Clean up the text before setting it
    const cleanedText = text.replace(/\n\s*\n/g, '\n') // Remove double line breaks
                           .replace(/^\s+|\s+$/g, '') // Trim whitespace
                           .trim();
    
    textRange.setText(cleanedText);
    
    let fontSize = CONFIG.DEFAULT_FONT_SIZE;
    const shapeHeight = shape.getHeight();
    
    while (calculateTextHeight(cleanedText, fontSize) > shapeHeight && fontSize > CONFIG.MIN_FONT_SIZE) {
      fontSize--;
    }
    
    textRange.getTextStyle().setFontSize(fontSize);
    Logger.log(`Set font size to ${fontSize} for text: ${cleanedText.substring(0, 50)}...`);
    
  } catch (error) {
    Logger.log('Error adjusting font size: ' + error.toString());
  }
}

/**
 * Calculates text height based on font size and content
 */
function calculateTextHeight(text, fontSize) {
  const numLines = (text.match(/\n/g) || []).length + 1;
  return fontSize * CONFIG.LINE_SPACING * numLines;
}

/**
 * Fetches scripture content from Bible API
 */
function fetchScriptureContent(scriptureReading) {
  if (!scriptureReading) return { passage: '', verse: '' };

  try {
    Logger.log('Fetching scripture for: ' + scriptureReading);
    
    const verses = scriptureReading.split(',').map(v => v.trim());
    const responses = [];
    
    // Fetch one at a time to avoid rate limits
    for (let verse of verses) {
      const url = `https://www.biblegateway.com/passage/?search=${encodeURIComponent(verse)}&version=NIV`;
      try {
        const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
        if (response.getResponseCode() === 200) {
          responses.push(response);
        }
        Utilities.sleep(1000); // Wait 1 second between requests
      } catch (error) {
        Logger.log('Error fetching verse ' + verse + ': ' + error.toString());
      }
    }
    
    const passages = responses
      .map(response => {
        const htmlContent = response.getContentText();
        return extractScriptureText(htmlContent);
      })
      .filter(text => text && text.trim() !== '');

    return {
      passage: passages.join(' '),
      verse: verses.join(', ')
    };
  } catch (error) {
    Logger.log('Error fetching scripture: ' + error.toString());
    return { passage: '', verse: '' };
  }
}

/**
 * Extracts scripture text from HTML
 */
function extractScriptureText(htmlContent) {
  try {
    // Find the std-text class element and extract content
    const stdTextRegex = /<[^>]*class\s*=\s*["']?[^"']*std-text[^"']*["']?[^>]*>([\s\S]*)/i;
    const stdTextMatch = htmlContent.match(stdTextRegex);
    
    if (!stdTextMatch || !stdTextMatch[1]) {
      Logger.log('Could not find std-text class content');
      return '';
    }
    
    let textContent = stdTextMatch[1];
    
    // Find the closing tag by counting div tags
    let divCount = 1;
    let endIndex = 0;
    
    for (let i = 0; i < textContent.length; i++) {
      if (textContent.substring(i, i + 4) === '<div') {
        let tagEnd = textContent.indexOf('>', i);
        if (tagEnd !== -1) {
          divCount++;
          i = tagEnd;
        }
      } else if (textContent.substring(i, i + 6) === '</div>') {
        divCount--;
        if (divCount === 0) {
          endIndex = i;
          break;
        }
        i += 5;
      }
    }
    
    textContent = endIndex > 0 ? textContent.substring(0, endIndex) : textContent;
    
    // Extract and preserve verse numbers
    const verseNumbers = [];
    textContent = textContent.replace(/<[^>]*class\s*=\s*["']?[^"']*versenum[^"']*["']?[^>]*>([\s\S]*?)<\/[^>]+>/gi, (match, verseNum) => {
      const cleanVerseNum = verseNum.replace(/<[^>]+>/g, '').trim();
      verseNumbers.push(cleanVerseNum);
      return `{{VERSE_${verseNumbers.length - 1}}}`;
    });
    
    // Remove other HTML elements
    textContent = textContent.replace(/<sup[^>]*>[\s\S]*?<\/sup>/gi, '');
    textContent = textContent.replace(/<[^>]+>/g, ' ');
    textContent = textContent.replace(/\(\s*[A-Z]\s*\)/g, '');
    textContent = textContent.replace(/\s+/g, ' ').trim();
    
    // Clean up HTML entities
    textContent = textContent.replace(/&nbsp;/g, ' ')
                             .replace(/&amp;/g, '&')
                             .replace(/&lt;/g, '<')
                             .replace(/&gt;/g, '>')
                             .replace(/&quot;/g, '"')
                             .replace(/&#39;/g, "'")
                             .replace(/&#\d+;/g, '');
    
    // Restore verse numbers
    verseNumbers.forEach((verseNum, index) => {
      textContent = textContent.replace(`{{VERSE_${index}}}`, verseNum + ' ');
    });
    
    return textContent;
    
  } catch (error) {
    Logger.log('Error extracting scripture text: ' + error.toString());
    return '';
  }
}

/**
 * Updates scripture slides with fetched content
 */
function updateScriptureSlides(slides, scriptureContent) {
  try {
    slides.forEach((slide, index) => {
      const shapes = slide.getShapes();
      shapes.forEach(shape => {
        try {
          const textRange = shape.getText();
          if (!textRange) return;
          
          const text = textRange.asString();
          if (text.includes(PLACEHOLDERS.PASSAGE)) {
            textRange.replaceAllText(PLACEHOLDERS.PASSAGE, scriptureContent.passage);
            adjustFontSizeToFitShape(shape, scriptureContent.passage);
            Logger.log('Updated scripture passage on slide ' + (index + 1));
          }
          if (text.includes(PLACEHOLDERS.VERSE)) {
            textRange.replaceAllText(PLACEHOLDERS.VERSE, scriptureContent.verse);
            Logger.log('Updated scripture verse reference on slide ' + (index + 1));
          }
        } catch (error) {
          // Skip shapes that don't have text
        }
      });
    });
  } catch (error) {
    Logger.log('Error updating scripture slides: ' + error.toString());
  }
}

/**
 * Updates sermon title slides
 */
function updateSermonSlides(slides, sermonTitle) {
  try {
    slides.forEach((slide, index) => {
      const shapes = slide.getShapes();
      shapes.forEach(shape => {
        try {
          const textRange = shape.getText();
          if (!textRange) return;
          
          const text = textRange.asString();
          if (text.includes(PLACEHOLDERS.SERMON)) {
            textRange.replaceAllText(PLACEHOLDERS.SERMON, sermonTitle || '');
            Logger.log('Updated sermon title on slide ' + (index + 1) + ': ' + sermonTitle);
          }
        } catch (error) {
          // Skip shapes that don't have text
        }
      });
    });
  } catch (error) {
    Logger.log('Error updating sermon slides: ' + error.toString());
  }
}

/**
 * Extracts hymn title from HTML content
 */
function extractHymnTitle(html) {
  try {
    const titleMatch = html.match(/<h1[^>]*class\s*=\s*["']?title\s+single-title\s+entry-title["']?[^>]*>(.*?)<\/h1>/);
    return titleMatch ? decodeHtmlEntities(titleMatch[1].trim()) : "Untitled Hymn";
  } catch (error) {
    Logger.log('Error extracting hymn title: ' + error.toString());
    return "Untitled Hymn";
  }
}

/**
 * Decodes HTML entities in text
 */
function decodeHtmlEntities(text) {
  try {
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
  } catch (error) {
    Logger.log('Error decoding HTML entities: ' + error.toString());
    return text;
  }
}

/**
 * Extracts hymn verses from HTML content
 */
function extractHymnVerses(html) {
  try {
    const tableMatches = html.match(/<table[^>]*>([\s\S]*?)<\/table>/g);
    if (!tableMatches) return { verses: [], refrain: "" };

    const contentBoxHtml = tableMatches[0];
    const verses = [];
    let refrain = "";

    const pTags = contentBoxHtml.match(/<p>([\s\S]*?)<\/p>/g) || [];
    
    pTags.forEach(pTag => {
      // First replace <br> tags with a special marker
      let verseHtml = pTag.replace(/<a[^>]*>.*?<\/a>/g, '')
                         .replace(/<br\s*\/?>/gi, '||LINEBREAK||')
                         .replace(/<\/?[^>]+(>|$)/g, "")
                         .trim();
      
      const decodedVerse = decodeHtmlEntities(verseHtml);
      
      if (decodedVerse && decodedVerse.trim()) {
        // Replace the markers with actual line breaks and clean up extra whitespace
        const cleanedVerse = decodedVerse.replace(/\|\|LINEBREAK\|\|/g, '\n')
                                        .replace(/\n\s*\n/g, '\n') // Remove double line breaks
                                        .replace(/^\s+|\s+$/g, '') // Trim start/end
                                        .replace(/[ \t]+/g, ' '); // Normalize spaces
        
        if (cleanedVerse.toLowerCase().includes("refrain")) {
          refrain = cleanedVerse;
        } else if (cleanedVerse.length > 0) {
          verses.push(cleanedVerse);
        }
      }
    });

    return { verses, refrain };
  } catch (error) {
    Logger.log('Error extracting hymn verses: ' + error.toString());
    return { verses: [], refrain: "" };
  }
}

/**
 * Modified searchGmailForPraiseLyrics to return data instead of updating slides directly
 */
function searchGmailForPraiseLyrics() {
  try {
    // Calculate date 6 days ago
    const tenDaysAgo = new Date();
    tenDaysAgo.setDate(tenDaysAgo.getDate() - 5);
    const dateString = Utilities.formatDate(tenDaysAgo, Session.getScriptTimeZone(), 'yyyy/MM/dd');
    
    // Search for emails with (praise AND lyrics) OR (worship AND lyrics) in subject within past 6 days
    const searchQuery = `after:${dateString} (subject:(praise lyrics) OR subject:(worship lyrics))`;
    
    const threads = GmailApp.search(searchQuery, 0, 5);
    
    if (threads.length === 0) {
      Logger.log('No praise/worship email threads found');
      return null;
    }
    
    // Get the most recent email (first in results)
    const mostRecentThread = threads[0];
    const message = mostRecentThread.getMessages()[0];
    
    // Get the HTML body
    let emailBody = message.getBody();
    
    // If HTML is empty, try plain text
    if (!emailBody || emailBody.trim() === '') {
      emailBody = message.getPlainBody();
      
      // For plain text, split by double newlines
      const lines = emailBody.split('\n').map(line => line.trim()).filter(line => line !== '');
      
      if (lines.length < 2) {
        Logger.log('Not enough content in plain text email');
        return null;
      }
      
      // First line is title
      const songTitle = lines[0];
      
      // Group remaining lines into paragraphs based on empty lines
      const lyrics = [];
      let currentParagraph = [];
      
      for (let i = 1; i < lines.length; i++) {
        if (lines[i] === '') {
          if (currentParagraph.length > 0) {
            lyrics.push(currentParagraph.join('\n'));
            currentParagraph = [];
          }
        } else {
          currentParagraph.push(lines[i]);
        }
      }
      
      if (currentParagraph.length > 0) {
        lyrics.push(currentParagraph.join('\n'));
      }
      
      return {
        title: songTitle,
        lyrics: lyrics,
        subject: message.getSubject(),
        date: message.getDate()
      };
    }
    
    // Process HTML email
    Logger.log('Processing HTML email');
    
    // Extract all div contents in order
    const divPattern = /<div[^>]*>(.*?)<\/div>/gi;
    const divContents = [];
    let match;
    
    while ((match = divPattern.exec(emailBody)) !== null) {
      let content = match[1];
      
      // Check if this is just a <br> (paragraph separator)
      if (content.match(/^\s*<br\s*\/?>\s*$/)) {
        divContents.push('||BREAK||');
      } else {
        // Clean the content
        content = content.replace(/<br\s*\/?>/gi, ' ')
                        .replace(/<[^>]*>/g, '')
                        .trim();
        
        if (content !== '') {
          // Decode HTML entities and quoted-printable
          content = content.replace(/&nbsp;/g, ' ')
                          .replace(/&amp;/g, '&')
                          .replace(/&lt;/g, '<')
                          .replace(/&gt;/g, '>')
                          .replace(/&quot;/g, '"')
                          .replace(/&#39;/g, "'")
                          .replace(/&#\d+;/g, '')
                          .replace(/=E2=80=99/g, "'")
                          .replace(/=\r?\n/g, '')
                          .replace(/=[0-9A-F]{2}/gi, '');
          
          divContents.push(content);
        }
      }
    }
    
    Logger.log('Found ' + divContents.length + ' div elements');
    
    // Now group the content based on ||BREAK|| markers
    const sections = [];
    let currentSection = [];
    
    for (let item of divContents) {
      if (item === '||BREAK||') {
        if (currentSection.length > 0) {
          sections.push(currentSection.join('\n'));
          currentSection = [];
        }
      } else {
        currentSection.push(item);
      }
    }
    
    // Don't forget the last section
    if (currentSection.length > 0) {
      sections.push(currentSection.join('\n'));
    }
    
    Logger.log('Grouped into ' + sections.length + ' sections');
    
    // If we have sections, first is title, rest are lyrics
    if (sections.length >= 2) {
      const songTitle = sections[0];
      const lyrics = sections.slice(1);
      
      Logger.log('Title: ' + songTitle);
      Logger.log('Number of lyric sections: ' + lyrics.length);
      
      return {
        title: songTitle,
        lyrics: lyrics,
        subject: message.getSubject(),
        date: message.getDate()
      };
    }
    
    // Fallback: If we don't have clear sections but have divContents
    // First non-break item is title, group rest by breaks
    if (divContents.length > 0) {
      // Filter out breaks and get real content
      const contentOnly = divContents.filter(item => item !== '||BREAK||');
      
      if (contentOnly.length < 2) {
        Logger.log('Not enough content found');
        return null;
      }
      
      const songTitle = contentOnly[0];
      
      // Group remaining content - look for natural paragraph breaks
      // Based on your screenshot: 2 lines, 4 lines, 4 lines, 4 lines
      const remainingLines = contentOnly.slice(1);
      const lyrics = [];
      
      // Your song structure appears to be:
      // First verse: 2 lines
      // Chorus: 4 lines  
      // Second verse: 4 lines
      // Chorus repeat: 4 lines
      
      if (remainingLines.length >= 14) {
        // We have enough lines for the full structure
        lyrics.push(remainingLines.slice(0, 2).join('\n'));  // First verse (2 lines)
        lyrics.push(remainingLines.slice(2, 6).join('\n'));  // First chorus (4 lines)
        lyrics.push(remainingLines.slice(6, 10).join('\n')); // Second verse (4 lines)
        lyrics.push(remainingLines.slice(10, 14).join('\n')); // Second chorus (4 lines)
      } else {
        // Try to intelligently group based on content
        let currentVerse = [];
        
        for (let i = 0; i < remainingLines.length; i++) {
          const line = remainingLines[i];
          currentVerse.push(line);
          
          // Check if this line suggests end of a verse/chorus
          // Look for lines ending with "silver", "gold", "will", "within", "sin"
          if (line.match(/(silver|gold|will|within|sin)$/i) || 
              (currentVerse.length === 4) || 
              (currentVerse.length === 2 && i < 4)) {
            lyrics.push(currentVerse.join('\n'));
            currentVerse = [];
          }
        }
        
        // Add any remaining lines
        if (currentVerse.length > 0) {
          lyrics.push(currentVerse.join('\n'));
        }
      }
      
      Logger.log('Parsed title: ' + songTitle);
      Logger.log('Created ' + lyrics.length + ' lyric sections');
      
      return {
        title: songTitle,
        lyrics: lyrics,
        subject: message.getSubject(),
        date: message.getDate()
      };
    }
    
    Logger.log('Could not parse email properly');
    return null;
    
  } catch (error) {
    Logger.log('Error searching Gmail for praise lyrics: ' + error.toString());
    return null;
  }
}

/**
 * Updated function to add praise song slides to existing presentation
 */
function updatePraiseSongSlides(presentation, praiseData) {
  if (!praiseData) {
    Logger.log('No praise data to process');
    return;
  }
  
  try {
    Logger.log('Starting updatePraiseSongSlides with data:');
    Logger.log('Title: ' + praiseData.title);
    Logger.log('Number of lyric sections: ' + praiseData.lyrics.length);
    
    const slides = presentation.getSlides();
    
    // First, replace {{praise_song}} with the title in all slides
    let praiseSongFound = false;
    slides.forEach((slide, index) => {
      const shapes = slide.getShapes();
      shapes.forEach(shape => {
        try {
          const textRange = shape.getText();
          if (!textRange) return;
          
          const text = textRange.asString();
          if (text.includes(PLACEHOLDERS.PRAISE_SONG)) {
            textRange.replaceAllText(PLACEHOLDERS.PRAISE_SONG, praiseData.title);
            praiseSongFound = true;
            Logger.log('Replaced {{praise_song}} with title on slide ' + (index + 1));
          }
        } catch (error) {
          // Skip shapes that don't have text
        }
      });
    });
    
    if (!praiseSongFound) {
      Logger.log('Warning: {{praise_song}} placeholder not found');
    }
    
    // Find the template slide containing {{praise_lyrics}}
    let templateSlide = null;
    let templateSlideIndex = -1;
    
    for (let i = 0; i < slides.length; i++) {
      const slide = slides[i];
      const shapes = slide.getShapes();
      
      for (let shape of shapes) {
        try {
          const text = shape.getText().asString();
          if (text.includes(PLACEHOLDERS.PRAISE_LYRICS)) {
            templateSlide = slide;
            templateSlideIndex = i;
            Logger.log('Found {{praise_lyrics}} placeholder on slide ' + (i + 1));
            break;
          }
        } catch (error) {
          // Skip shapes that don't have text
        }
      }
      
      if (templateSlide) break;
    }
    
    if (!templateSlide) {
      Logger.log('Warning: {{praise_lyrics}} placeholder not found in any slide');
      
      // Alternative: If we can't find {{praise_lyrics}}, check if {{praise_song}} was replaced
      // and use that slide to add the lyrics
      for (let i = 0; i < slides.length; i++) {
        const slide = slides[i];
        const shapes = slide.getShapes();
        
        for (let shape of shapes) {
          try {
            const text = shape.getText().asString();
            // Check if this slide contains the praise song title we just added
            if (text.includes(praiseData.title)) {
              // This might be our target slide
              // Replace the entire content with first verse, then duplicate for others
              templateSlide = slide;
              templateSlideIndex = i;
              Logger.log('Using slide with praise song title as template (slide ' + (i + 1) + ')');
              
              // Replace the title with the first verse
              const firstVerse = praiseData.lyrics[0];
              if (firstVerse) {
                shape.getText().setText(firstVerse);
                adjustFontSizeToFitShape(shape, firstVerse);
                Logger.log('Replaced title with first verse');
                
                // Now create additional slides for remaining verses
                for (let j = 1; j < praiseData.lyrics.length; j++) {
                  const newSlide = slide.duplicate();
                  const newShapes = newSlide.getShapes();
                  
                  // Find the text shape in the new slide
                  for (let newShape of newShapes) {
                    try {
                      const textRange = newShape.getText();
                      if (textRange && textRange.asString().includes(firstVerse)) {
                        textRange.setText(praiseData.lyrics[j]);
                        adjustFontSizeToFitShape(newShape, praiseData.lyrics[j]);
                        Logger.log('Created slide ' + (j + 1) + ' for verse: ' + praiseData.lyrics[j].substring(0, 30) + '...');
                        break;
                      }
                    } catch (error) {
                      // Continue
                    }
                  }
                }
                
                return; // We're done
              }
              break;
            }
          } catch (error) {
            // Skip shapes that don't have text
          }
        }
        
        if (templateSlide) break;
      }
      
      if (!templateSlide) {
        Logger.log('Could not find any slide to use for praise lyrics');
        return;
      }
    }
    
    // Filter out any empty paragraphs before processing
    const validParagraphs = praiseData.lyrics.filter(para => para && para.trim() !== '');
    
    Logger.log('Valid paragraphs to create: ' + validParagraphs.length);
    
    if (validParagraphs.length === 0) {
      Logger.log('No valid paragraphs to process');
      return;
    }
    
    // Log each paragraph for debugging
    validParagraphs.forEach((para, index) => {
      Logger.log('Paragraph ' + (index + 1) + ': ' + para.substring(0, 50) + '...');
    });
    
    // Create slides for each paragraph
    // First, update the template slide with the first paragraph
    const shapes = templateSlide.getShapes();
    let lyricsShapeFound = false;
    
    shapes.forEach(shape => {
      try {
        const textRange = shape.getText();
        if (!textRange) return;
        
        const text = textRange.asString();
        if (text.includes(PLACEHOLDERS.PRAISE_LYRICS)) {
          textRange.replaceAllText(PLACEHOLDERS.PRAISE_LYRICS, validParagraphs[0].trim());
          adjustFontSizeToFitShape(shape, validParagraphs[0].trim());
          lyricsShapeFound = true;
          Logger.log('Replaced {{praise_lyrics}} with first paragraph');
        }
      } catch (error) {
        // Skip shapes that don't have text
      }
    });
    
    if (!lyricsShapeFound) {
      Logger.log('Warning: Could not find {{praise_lyrics}} in the template slide shapes');
    }
    
    // Create duplicates for remaining paragraphs
    for (let i = 1; i < validParagraphs.length; i++) {
      const newSlide = templateSlide.duplicate();
      const newShapes = newSlide.getShapes();
      
      // Update the text in the duplicated slide
      newShapes.forEach(shape => {
        try {
          const textRange = shape.getText();
          if (!textRange) return;
          
          const text = textRange.asString();
          // The duplicated slide will have the first paragraph, replace it with the current paragraph
          if (text.includes(validParagraphs[0].trim())) {
            textRange.setText(validParagraphs[i].trim());
            adjustFontSizeToFitShape(shape, validParagraphs[i].trim());
            Logger.log('Created slide for paragraph ' + (i + 1));
          }
        } catch (error) {
          // Skip shapes that don't have text
        }
      });
    }
    
    Logger.log('Praise song slides creation completed');
    
  } catch (error) {
    Logger.log('Error in updatePraiseSongSlides: ' + error.toString());
  }
}

/**
 * Test function to verify the script works
 */
function testScript() {
  try {
    Logger.log('Testing script functionality...');
    
    // Test upcoming Saturday calculation
    const upcomingSaturday = getUpcomingSaturday();
    Logger.log('Upcoming Saturday: ' + getDateFormatted(upcomingSaturday));
    
    // Test spreadsheet connection
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    if (spreadsheet) {
      Logger.log('Successfully connected to spreadsheet');
      const targetSheet = findTargetSheet(spreadsheet);
      if (targetSheet) {
        Logger.log('Found target sheet: ' + targetSheet.getName());
        
        // Test cleaning announcements
        const cleaningAnnouncements = getCleaningAnnouncements(targetSheet, upcomingSaturday);
        Logger.log('Today\'s cleaning announcement: ' + cleaningAnnouncements.today);
        Logger.log('Upcoming cleaning announcement: ' + cleaningAnnouncements.upcoming);
      } else {
        Logger.log('Could not find target sheet');
      }
    } else {
      Logger.log('Could not connect to spreadsheet');
    }
    
    // Test Gmail search (if permission granted)
    try {
      const praiseData = searchGmailForPraiseLyrics();
      if (praiseData) {
        Logger.log('Found praise song: ' + praiseData.title);
      } else {
        Logger.log('No praise songs found in Gmail');
      }
    } catch (error) {
      Logger.log('Gmail access not available or error: ' + error.toString());
    }
    
    Logger.log('Test completed');
    
  } catch (error) {
    Logger.log('Test error: ' + error.toString());
  }
}
