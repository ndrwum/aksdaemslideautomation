/**
 * Google Apps Script for automating hymn slides creation
 * This script reads from a spreadsheet and creates a presentation with hymn lyrics
 * Now includes Gmail integration for praise/worship songs
 */

// Configuration constants
const CONFIG = {
  TEMPLATE_ID: 'HIDDEN',
  SPREADSHEET_ID: 'HIDDEN',
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
  STORY: '{{story}}',
  PRAISE_SONG: '{{praise_song}}',
  PRAISE_LYRICS: '{{praise_lyrics}}'
};

/**
 * Main function to create hymn slides (updated to include praise songs)
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
    
    createPresentation(hymnsData, hymnDetails, scriptureContent, upcomingSaturdayString, praiseData);
    
  } catch (error) {
    Logger.log('Error in createHymnsSlides: ' + error.toString());
    Logger.log('Error stack: ' + error.stack);
  }
}

/**
 * Creates the presentation with all slides (updated to include praise songs)
 */
function createPresentation(hymnsData, hymnDetails, scriptureContent, presentationName, praiseData) {
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
 * Finds the target sheet containing "Sabbath Schedule 2024"
 */
function findTargetSheet(spreadsheet) {
  try {
    const sheets = spreadsheet.getSheets();
    for (let sheet of sheets) {
      if (sheet.getName().includes("Sabbath Schedule 2024")) {
        Logger.log('Found target sheet: ' + sheet.getName());
        return sheet;
      }
    }
    Logger.log('No sheet found containing "Sabbath Schedule 2024"');
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
    // Calculate date 10 days ago
    const tenDaysAgo = new Date();
    tenDaysAgo.setDate(tenDaysAgo.getDate() - 10);
    const dateString = Utilities.formatDate(tenDaysAgo, Session.getScriptTimeZone(), 'yyyy/MM/dd');
    
    // Search for emails with (praise AND lyrics) OR (worship AND lyrics) in subject within past 10 days
    const searchQuery = `after:${dateString} (subject:(praise lyrics) OR subject:(worship lyrics))`;
    
    Logger.log('Search query: ' + searchQuery);
    
    const threads = GmailApp.search(searchQuery, 0, 10);
    Logger.log('Found ' + threads.length + ' matching emails');
    
    if (threads.length === 0) {
      Logger.log('No emails found matching the criteria');
      return null;
    }
    
    // Get the most recent email (first in results)
    const mostRecentThread = threads[0];
    const message = mostRecentThread.getMessages()[0];
    
    Logger.log('Processing email with subject: ' + message.getSubject());
    Logger.log('Email date: ' + message.getDate());
    
    // Try HTML version first to preserve formatting
    let emailBody = message.getBody();
    let isHtml = true;
    
    // If HTML is empty, try plain text
    if (!emailBody || emailBody.trim() === '') {
      emailBody = message.getPlainBody();
      isHtml = false;
      Logger.log('Using plain text version');
    } else {
      Logger.log('Using HTML version');
    }
    
    Logger.log('Raw email body length: ' + emailBody.length);
    
    if (isHtml) {
      // Better HTML to text conversion - preserve paragraph breaks AND line breaks within paragraphs
      emailBody = emailBody.replace(/<div[^>]*>/gi, '')
                          .replace(/<\/div>/gi, '\n\n')
                          .replace(/<p[^>]*>/gi, '')
                          .replace(/<\/p>/gi, '\n\n')
                          .replace(/<br\s*\/?>/gi, '\n')
                          .replace(/<[^>]*>/g, ''); // Remove all other HTML tags
      
      // Clean up HTML entities
      emailBody = emailBody.replace(/&nbsp;/g, ' ')
                          .replace(/&amp;/g, '&')
                          .replace(/&lt;/g, '<')
                          .replace(/&gt;/g, '>')
                          .replace(/&quot;/g, '"')
                          .replace(/&#39;/g, "'")
                          .replace(/&#\d+;/g, '');
    }
    
    // Remove markdown formatting
    emailBody = emailBody.replace(/^\*\*(.*?)\*\*$/gm, '$1') // Remove **bold** formatting from lines
                        .replace(/^\*(.*?)\*$/gm, '$1');     // Remove *italic* formatting from lines
    
    // If we still don't have proper line breaks, try a different approach
    if (!emailBody.includes('\n\n')) {
      Logger.log('No double line breaks found, trying alternative parsing');
      
      // If the content appears to be on one line, try to intelligently split it
      // Look for patterns that indicate new verses/paragraphs
      emailBody = emailBody.replace(/(\s)(Just the time)/g, '\n\n$2')
                          .replace(/(\s)(Oh, I want to know You more)/g, '\n\n$2')
                          .replace(/(\s)(And when my daily)/g, '\n\n$2')
                          .replace(/(\s)(And I would give)/g, '\n\n$2')
                          .replace(/(\s)(To know You in)/g, '\n$2'); // This should stay with previous line
      
      // Also add line breaks within paragraphs for better readability
      emailBody = emailBody.replace(/(self)(\s+)(Just the time)/g, '$1\n$3')
                          .replace(/(wealth)(\s+)(That's when)/g, '$1\n$3')
                          .replace(/(call)(\s+)(And all my)/g, '$1\n$3')
                          .replace(/(you)(\s+)(Oh, I want to know You,)/g, '$1\n$3')
                          .replace(/(mind)(\s+)(Looking in)/g, '$1\n$3')
                          .replace(/(me)(\s+)(Cries that say)/g, '$1\n$3')
                          .replace(/(song)(\s+)(My heart begins)/g, '$1\n$3')
                          .replace(/(gone)(\s+)(I've run)/g, '$1\n$3')
                          .replace(/(soul)(\s+)(Now, the gentle)/g, '$1\n$3')
                          .replace(/(You)(\s+)(Oh, I want to know You to)/g, '$1\n$3')
                          .replace(/(mind)(\s+)(Looking in Your eyes)/g, '$1\n$3')
                          .replace(/(me)(\s+)(Cries that say I want)/g, '$1\n$3')
                          .replace(/(breath)(\s+)(To know You)/g, '$1\n$3')
                          .replace(/(resurrection)(\s+)(Oh, I want)/g, '$1\n$3')
                          .replace(/(more)(\s+)(Oh, I want to know You,)/g, '$1\n$3');
    }
    
    Logger.log('After processing: "' + emailBody.substring(0, 400) + '"');
    
    // Split by double line breaks to get paragraphs
    let sections = emailBody.split(/\n\s*\n/)
                           .map(section => section.trim())
                           .filter(section => section !== '');
    
    Logger.log('Found ' + sections.length + ' sections after splitting:');
    sections.forEach((section, index) => {
      Logger.log(`Section ${index + 1}: "${section.substring(0, 100)}..."`);
    });
    
    if (sections.length === 0) {
      Logger.log('No content found in email');
      return null;
    }
    
    // First section should be just the title
    let songTitle, lyricsParagraphs;
    
    if (sections.length === 1) {
      // Everything is in one section, need to extract title manually
      const fullContent = sections[0];
      // Title is likely the first few words
      const words = fullContent.split(' ');
      songTitle = words.slice(0, 7).join(' '); // "Oh, I want to know You more"
      
      // Rest is lyrics - split into logical paragraphs
      const remainingContent = words.slice(7).join(' ');
      lyricsParagraphs = [
        remainingContent.substring(0, remainingContent.indexOf('Oh, I want to know You more')),
        remainingContent.substring(remainingContent.indexOf('Oh, I want to know You more'), remainingContent.indexOf('And when my daily')),
        remainingContent.substring(remainingContent.indexOf('And when my daily'), remainingContent.indexOf('Oh, I want to know You more', remainingContent.indexOf('And when my daily'))),
        remainingContent.substring(remainingContent.lastIndexOf('Oh, I want to know You more'))
      ].map(p => p.trim()).filter(p => p !== '');
      
    } else {
      songTitle = sections[0];
      lyricsParagraphs = sections.slice(1);
    }
    
    // Clean up paragraphs - normalize whitespace but preserve intentional line breaks
    lyricsParagraphs = lyricsParagraphs.map(paragraph => {
      return paragraph.replace(/[ \t]+/g, ' ') // Replace multiple spaces/tabs with single space
                     .replace(/\n +/g, '\n')    // Remove spaces at beginning of lines
                     .replace(/ +\n/g, '\n')    // Remove spaces at end of lines
                     .trim();
    });
    
    Logger.log('Song title: "' + songTitle + '"');
    Logger.log('Number of lyric paragraphs: ' + lyricsParagraphs.length);
    
    Logger.log('Final lyric paragraphs:');
    lyricsParagraphs.forEach((para, index) => {
      Logger.log(`Paragraph ${index + 1}: "${para}"`);
    });
    
    // Return the data
    return {
      title: songTitle,
      lyrics: lyricsParagraphs,
      subject: message.getSubject(),
      date: message.getDate()
    };
    
  } catch (error) {
    Logger.log('Error searching Gmail: ' + error.toString());
    Logger.log('Error stack: ' + error.stack);
    return null;
  }
}

/**
 * Updated function to add praise song slides to existing presentation
 */
function updatePraiseSongSlides(presentation, praiseData) {
  if (!praiseData) {
    Logger.log('No praise song data to add');
    return;
  }
  
  try {
    const slides = presentation.getSlides();
    Logger.log('Adding praise song slides to presentation with ' + slides.length + ' slides');
    
    // Replace {{praise_song}} in all slides
    slides.forEach((slide, index) => {
      const shapes = slide.getShapes();
      shapes.forEach(shape => {
        try {
          const textRange = shape.getText();
          if (!textRange) return;
          
          const text = textRange.asString();
          if (text.includes('{{praise_song}}')) {
            textRange.replaceAllText('{{praise_song}}', praiseData.title);
            Logger.log('Replaced {{praise_song}} with "' + praiseData.title + '" on slide ' + (index + 1));
          }
        } catch (error) {
          // Skip shapes that don't have text
        }
      });
    });
    
    // Find the template slide containing {{praise_lyrics}}
    let templateSlide = null;
    let templateSlideIndex = -1;
    
    for (let i = 0; i < slides.length; i++) {
      const slide = slides[i];
      const shapes = slide.getShapes();
      
      for (let shape of shapes) {
        try {
          const text = shape.getText().asString();
          if (text.includes('{{praise_lyrics}}')) {
            templateSlide = slide;
            templateSlideIndex = i;
            break;
          }
        } catch (error) {
          // Skip shapes that don't have text
        }
      }
      
      if (templateSlide) break;
    }
    
    if (!templateSlide) {
      Logger.log('Could not find slide with {{praise_lyrics}} placeholder');
      return;
    }
    
    Logger.log('Found {{praise_lyrics}} template on slide ' + (templateSlideIndex + 1));
    
    // First, create duplicates for all paragraphs BEFORE modifying any text
    // This ensures we're always duplicating the original template
    const createdSlides = [templateSlide]; // First slide is the original template
    
    // Create duplicate slides for paragraphs 2 and onwards
    for (let i = 1; i < praiseData.lyrics.length; i++) {
      const duplicatedSlide = templateSlide.duplicate();
      createdSlides.push(duplicatedSlide);
      Logger.log('Created duplicate slide for paragraph ' + (i + 1));
    }
    
    // Now update the text in each slide
    praiseData.lyrics.forEach((paragraph, index) => {
      if (paragraph.trim() === '') return; // Skip empty paragraphs
      
      const targetSlide = createdSlides[index];
      if (!targetSlide) {
        Logger.log('Warning: No slide available for paragraph ' + (index + 1));
        return;
      }
      
      // Find and update the text shape containing {{praise_lyrics}}
      const shapes = targetSlide.getShapes();
      let textUpdated = false;
      
      shapes.forEach(shape => {
        try {
          const textRange = shape.getText();
          if (!textRange) return;
          
          const text = textRange.asString();
          if (text.includes('{{praise_lyrics}}')) {
            textRange.replaceAllText('{{praise_lyrics}}', paragraph.trim());
            // Apply font sizing for the paragraph
            adjustFontSizeToFitShape(shape, paragraph.trim());
            Logger.log('Updated paragraph ' + (index + 1) + ' on slide with text: "' + paragraph.substring(0, 50) + '..."');
            textUpdated = true;
          }
        } catch (error) {
          Logger.log('Error updating shape: ' + error.toString());
        }
      });
      
      if (!textUpdated) {
        Logger.log('Warning: Could not find {{praise_lyrics}} placeholder in slide shapes for paragraph ' + (index + 1));
      }
    });
    
    Logger.log('Successfully added praise song slides with ' + praiseData.lyrics.length + ' paragraphs');
    
  } catch (error) {
    Logger.log('Error adding praise song slides: ' + error.toString());
    Logger.log('Error stack: ' + error.stack);
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
