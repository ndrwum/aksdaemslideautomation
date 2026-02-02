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
  SCRIPTURE_READING: 'Scripture Reading',
  SCRIPTURE_READER: 'Scripture Reader',
  SERMON_TITLE: 'Sermon Title',
  SPEAKER: 'Speaker',
  SPECIAL_MUSIC: 'Special Music',
  INTERCESSORY_PRAYER: 'Intercessory Prayer',
  CHILDREN_STORY: "Children's Story"
};

// Placeholders
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
  THIS_WEEK_DATE: '{{this_week_date}}',
  THIS_WEEK_LEADERS: '{{this_week_leaders}}',
  NEXT_WEEK_DATE: '{{next_week_date}}',
  NEXT_WEEK_LEADERS: '{{next_week_leaders}}',
  WEEK_AFTER_DATE: '{{week_after_date}}',
  WEEK_AFTER_LEADERS: '{{week_after_leaders}}'
};

/**
 * Main function to create hymn slides
 */
function createHymnsSlides() {
  try {
    Logger.log('Starting createHymnsSlides');
    
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
    const praiseData = searchGmailForPraiseLyrics();
    const bulletinLeadersData = getBulletinLeadersData(spreadsheet, upcomingSaturday);
    
    createPresentation(hymnsData, hymnDetails, scriptureContent, upcomingSaturdayString, praiseData, bulletinLeadersData);
    
  } catch (error) {
    Logger.log('Error in createHymnsSlides: ' + error.toString());
  }
}

/**
 * Gets bulletin leaders data from "For Bulletin" sheet
 */
function getBulletinLeadersData(spreadsheet, upcomingSaturday) {
  try {
    const bulletinSheet = spreadsheet.getSheetByName('For Bulletin');
    if (!bulletinSheet) {
      Logger.log('Could not find "For Bulletin" sheet');
      return null;
    }
    
    const dataRange = bulletinSheet.getDataRange().getValues();
    if (dataRange.length < 2) {
      return null;
    }
    
    const headerRow = dataRange[0];
    
    const thisWeekSaturday = new Date(upcomingSaturday);
    const nextWeekSaturday = new Date(upcomingSaturday);
    nextWeekSaturday.setDate(nextWeekSaturday.getDate() + 7);
    const weekAfterSaturday = new Date(upcomingSaturday);
    weekAfterSaturday.setDate(weekAfterSaturday.getDate() + 14);
    
    const thisWeekString = getDateFormatted(thisWeekSaturday);
    const nextWeekString = getDateFormatted(nextWeekSaturday);
    const weekAfterString = getDateFormatted(weekAfterSaturday);
    
    let thisWeekData = null;
    let nextWeekData = null;
    let weekAfterData = null;
    
    for (let i = 1; i < dataRange.length; i++) {
      const dateCell = dataRange[i][0];
      
      if (dateCell instanceof Date) {
        const dateString = getDateFormatted(dateCell);
        
        if (dateString === thisWeekString) {
          thisWeekData = formatBulletinRow(headerRow, dataRange[i], thisWeekSaturday);
        }
        if (dateString === nextWeekString) {
          nextWeekData = formatBulletinRow(headerRow, dataRange[i], nextWeekSaturday);
        }
        if (dateString === weekAfterString) {
          weekAfterData = formatBulletinRow(headerRow, dataRange[i], weekAfterSaturday);
        }
      }
    }
    
    return {
      thisWeek: thisWeekData || { date: thisWeekString, leaders: '' },
      nextWeek: nextWeekData || { date: nextWeekString, leaders: '' },
      weekAfter: weekAfterData || { date: weekAfterString, leaders: '' }
    };
    
  } catch (error) {
    Logger.log('Error getting bulletin leaders data: ' + error.toString());
    return null;
  }
}

/**
 * Formats a bulletin row into the required format
 */
function formatBulletinRow(headerRow, rowData, date) {
  try {
    const leaderPairs = [];
    
    for (let i = 1; i < headerRow.length; i++) {
      const columnTitle = headerRow[i];
      const cellValue = rowData[i];
      
      if (!columnTitle || columnTitle.toString().trim() === '') {
        continue;
      }
      
      if (cellValue !== undefined && cellValue !== null && cellValue.toString().trim() !== '') {
        let formattedValue = '';
        
        if (columnTitle.toString().trim().toLowerCase() === 'cleaners') {
          formattedValue = formatCleanersValue(cellValue.toString());
        } else {
          formattedValue = cellValue.toString().trim();
        }
        
        leaderPairs.push(columnTitle.toString().trim() + ': ' + formattedValue);
      }
    }
    
    return {
      date: getDateFormatted(date),
      leaders: leaderPairs.join('\n')
    };
    
  } catch (error) {
    Logger.log('Error formatting bulletin row: ' + error.toString());
    return { date: getDateFormatted(date), leaders: '' };
  }
}

/**
 * Formats the Cleaners column value
 */
function formatCleanersValue(cleanersContent) {
  try {
    if (!cleanersContent || cleanersContent.trim() === '') {
      return '';
    }
    
    const names = cleanersContent.split(',').map(name => name.trim()).filter(name => name !== '');
    const dishwashers = [];
    const tableCleaners = [];
    
    names.forEach(name => {
      if (name.startsWith('*')) {
        tableCleaners.push(name.substring(1).trim());
      } else if (name.endsWith('*')) {
        tableCleaners.push(name.substring(0, name.length - 1).trim());
      } else {
        dishwashers.push(name);
      }
    });
    
    let formatted = '';
    
    if (dishwashers.length > 0) {
      formatted += 'Dishwashers: ' + dishwashers.join(', ');
    }
    
    if (tableCleaners.length > 0) {
      if (formatted !== '') {
        formatted += '\n';
      }
      formatted += 'Table cleaners: ' + tableCleaners.join(', ');
    }
    
    return formatted;
    
  } catch (error) {
    return cleanersContent;
  }
}

/**
 * Creates the presentation with all slides
 */
function createPresentation(hymnsData, hymnDetails, scriptureContent, presentationName, praiseData, bulletinLeadersData) {
  try {
    Logger.log('Creating presentation: ' + presentationName);
    
    const presentation = SlidesApp.openById(
      DriveApp.getFileById(CONFIG.TEMPLATE_ID)
        .makeCopy(presentationName)
        .getId()
    );

    const slides = presentation.getSlides();
    const templateSlides = findTemplateSlides(slides);
    
    if (!areAllTemplateSlidesFound(templateSlides)) {
      Logger.log('Missing template slides');
      return;
    }

    updateTitleSlides(templateSlides, hymnDetails);
    createVersesSlides(templateSlides, hymnDetails);
    updateScriptureSlides(slides, scriptureContent);
    updateSermonSlides(slides, hymnsData.sermonTitle);
    updateParticipantsSlides(slides, hymnsData);
    
    if (bulletinLeadersData) {
      updateBulletinLeadersSlides(slides, bulletinLeadersData);
    }
    
    if (praiseData) {
      updatePraiseSongSlides(presentation, praiseData);
    }

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
  }
}

/**
 * Updates slides with bulletin leaders placeholders
 */
function updateBulletinLeadersSlides(slides, bulletinLeadersData) {
  try {
    slides.forEach((slide) => {
      const shapes = slide.getShapes();
      shapes.forEach(shape => {
        try {
          const textRange = shape.getText();
          if (!textRange) return;
          
          const text = textRange.asString();
          
          if (text.includes(PLACEHOLDERS.THIS_WEEK_DATE)) {
            textRange.replaceAllText(PLACEHOLDERS.THIS_WEEK_DATE, bulletinLeadersData.thisWeek.date);
          }
          if (text.includes(PLACEHOLDERS.THIS_WEEK_LEADERS)) {
            textRange.replaceAllText(PLACEHOLDERS.THIS_WEEK_LEADERS, bulletinLeadersData.thisWeek.leaders);
          }
          if (text.includes(PLACEHOLDERS.NEXT_WEEK_DATE)) {
            textRange.replaceAllText(PLACEHOLDERS.NEXT_WEEK_DATE, bulletinLeadersData.nextWeek.date);
          }
          if (text.includes(PLACEHOLDERS.NEXT_WEEK_LEADERS)) {
            textRange.replaceAllText(PLACEHOLDERS.NEXT_WEEK_LEADERS, bulletinLeadersData.nextWeek.leaders);
          }
          if (text.includes(PLACEHOLDERS.WEEK_AFTER_DATE)) {
            textRange.replaceAllText(PLACEHOLDERS.WEEK_AFTER_DATE, bulletinLeadersData.weekAfter.date);
          }
          if (text.includes(PLACEHOLDERS.WEEK_AFTER_LEADERS)) {
            textRange.replaceAllText(PLACEHOLDERS.WEEK_AFTER_LEADERS, bulletinLeadersData.weekAfter.leaders);
          }
        } catch (error) {
          // Skip shapes that don't have text
        }
      });
    });
  } catch (error) {
    Logger.log('Error updating bulletin leaders slides: ' + error.toString());
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
        return sheet;
      }
    }
    return null;
  } catch (error) {
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
      return {};
    }
    
    const headerRow = dataRange[1];
    const columnIndices = getColumnIndices(headerRow);

    for (let i = 2; i < dataRange.length; i++) {
      const dateCell = dataRange[i][0];
      
      if (dateCell instanceof Date && getDateFormatted(dateCell) === targetDate) {
        return {
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
      }
    }
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
    
    switch(trimmedHeader) {
      case 'Opening Hymn': indices.openingHymn = index; break;
      case 'Closing Hymn': indices.closingHymn = index; break;
      case 'Scripture Reading': indices.scriptureReading = index; break;
      case 'Sermon Title': indices.sermonTitle = index; break;
      case 'Speaker': indices.speaker = index; break;
      case 'Special Music': indices.specialMusic = index; break;
      case 'Intercessory Prayer': indices.prayer = index; break;
      case 'Scripture Reader': indices.reader = index; break;
      case "Children's Story": indices.story = index; break;
    }
  });

  return indices;
}

function updateParticipantsSlides(slides, hymnsData) {
  try {
    slides.forEach((slide) => {
      const shapes = slide.getShapes();
      shapes.forEach(shape => {
        try {
          const textRange = shape.getText();
          if (!textRange) return;
          
          const text = textRange.asString();
          
          const replacements = {
            '{{speaker}}': hymnsData.speaker || '',
            '{{music}}': hymnsData.specialMusic || '',
            '{{prayer}}': hymnsData.prayer || '',
            '{{reading}}': hymnsData.reader || '',
            '{{story}}': hymnsData.story || ''
          };

          Object.entries(replacements).forEach(([placeholder, value]) => {
            if (text.includes(placeholder)) {
              textRange.replaceAllText(placeholder, value);
            }
          });
        } catch (error) {
          // Skip
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
    const openingLyricsUrl = `https://sdahymnals.com/Hymnal/${openingHymnNumber}`;
    const closingLyricsUrl = `https://sdahymnals.com/Hymnal/${closingHymnNumber}`;
    
    const openingResponse = UrlFetchApp.fetch(openingLyricsUrl, { muteHttpExceptions: true });
    const closingResponse = UrlFetchApp.fetch(closingLyricsUrl, { muteHttpExceptions: true });
    
    if (openingResponse.getResponseCode() !== 200 || closingResponse.getResponseCode() !== 200) {
      return null;
    }
    
    const openingLyricsHtml = openingResponse.getContentText();
    const closingLyricsHtml = closingResponse.getContentText();
    
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
        // Skip
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
  try {
    if (templates.openingTitle && hymnDetails.opening) {
      templates.openingTitle.replaceAllText(PLACEHOLDERS.OPENING, hymnDetails.opening.title);
    }
    if (templates.closingTitle && hymnDetails.closing) {
      templates.closingTitle.replaceAllText(PLACEHOLDERS.CLOSING, hymnDetails.closing.title);
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
    const createSlidesForHymn = (verses, refrain, templateSlide) => {
      if (!verses || !Array.isArray(verses) || !templateSlide) {
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

      if (slidesToCreate.length > 0) {
        const lastSlideText = slidesToCreate[slidesToCreate.length - 1].text;
        if (/^\d+$/.test(lastSlideText.trim())) {
          slidesToCreate.pop();
        }
      }

      slidesToCreate.reverse().forEach((slideData) => {
        try {
          const newSlide = templateSlide.duplicate();
          const textShape = findMainTextShape(newSlide);
          if (textShape) {
            adjustFontSizeToFitShape(textShape, slideData.text);
          }
        } catch (error) {
          // Skip
        }
      });
    };

    if (hymnDetails.opening) {
      createSlidesForHymn(hymnDetails.opening.verses, hymnDetails.opening.refrain, templates.openingLyrics);
    }
    if (hymnDetails.closing) {
      createSlidesForHymn(hymnDetails.closing.verses, hymnDetails.closing.refrain, templates.closingLyrics);
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
        // Continue
      }
    }
    return shapes.length > 0 ? shapes[0] : null;
  } catch (error) {
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
    
    const cleanedText = text.replace(/\n\s*\n/g, '\n').replace(/^\s+|\s+$/g, '').trim();
    
    textRange.setText(cleanedText);
    
    let fontSize = CONFIG.DEFAULT_FONT_SIZE;
    const shapeHeight = shape.getHeight();
    
    while (calculateTextHeight(cleanedText, fontSize) > shapeHeight && fontSize > CONFIG.MIN_FONT_SIZE) {
      fontSize--;
    }
    
    textRange.getTextStyle().setFontSize(fontSize);
  } catch (error) {
    // Skip
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
    const verses = scriptureReading.split(',').map(v => v.trim());
    const responses = [];
    
    for (let verse of verses) {
      const url = `https://www.biblegateway.com/passage/?search=${encodeURIComponent(verse)}&version=NIV`;
      try {
        const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
        if (response.getResponseCode() === 200) {
          responses.push(response);
        }
        Utilities.sleep(1000);
      } catch (error) {
        // Skip
      }
    }
    
    const passages = responses
      .map(response => extractScriptureText(response.getContentText()))
      .filter(text => text && text.trim() !== '');

    return {
      passage: passages.join(' '),
      verse: verses.join(', ')
    };
  } catch (error) {
    return { passage: '', verse: '' };
  }
}

/**
 * Extracts scripture text from HTML
 */
function extractScriptureText(htmlContent) {
  try {
    const stdTextRegex = /<[^>]*class\s*=\s*["']?[^"']*std-text[^"']*["']?[^>]*>([\s\S]*)/i;
    const stdTextMatch = htmlContent.match(stdTextRegex);
    
    if (!stdTextMatch || !stdTextMatch[1]) {
      return '';
    }
    
    let textContent = stdTextMatch[1];
    
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
    
    const verseNumbers = [];
    textContent = textContent.replace(/<[^>]*class\s*=\s*["']?[^"']*versenum[^"']*["']?[^>]*>([\s\S]*?)<\/[^>]+>/gi, (match, verseNum) => {
      const cleanVerseNum = verseNum.replace(/<[^>]+>/g, '').trim();
      verseNumbers.push(cleanVerseNum);
      return `{{VERSE_${verseNumbers.length - 1}}}`;
    });
    
    textContent = textContent.replace(/<sup[^>]*>[\s\S]*?<\/sup>/gi, '');
    textContent = textContent.replace(/<[^>]+>/g, ' ');
    textContent = textContent.replace(/\(\s*[A-Z]\s*\)/g, '');
    textContent = textContent.replace(/\s+/g, ' ').trim();
    
    textContent = textContent.replace(/&nbsp;/g, ' ')
                             .replace(/&amp;/g, '&')
                             .replace(/&lt;/g, '<')
                             .replace(/&gt;/g, '>')
                             .replace(/&quot;/g, '"')
                             .replace(/&#39;/g, "'")
                             .replace(/&#\d+;/g, '');
    
    verseNumbers.forEach((verseNum, index) => {
      textContent = textContent.replace(`{{VERSE_${index}}}`, verseNum + ' ');
    });
    
    return textContent;
  } catch (error) {
    return '';
  }
}

/**
 * Updates scripture slides with fetched content
 */
function updateScriptureSlides(slides, scriptureContent) {
  try {
    slides.forEach((slide) => {
      const shapes = slide.getShapes();
      shapes.forEach(shape => {
        try {
          const textRange = shape.getText();
          if (!textRange) return;
          
          const text = textRange.asString();
          if (text.includes(PLACEHOLDERS.PASSAGE)) {
            textRange.replaceAllText(PLACEHOLDERS.PASSAGE, scriptureContent.passage);
            adjustFontSizeToFitShape(shape, scriptureContent.passage);
          }
          if (text.includes(PLACEHOLDERS.VERSE)) {
            textRange.replaceAllText(PLACEHOLDERS.VERSE, scriptureContent.verse);
          }
        } catch (error) {
          // Skip
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
    slides.forEach((slide) => {
      const shapes = slide.getShapes();
      shapes.forEach(shape => {
        try {
          const textRange = shape.getText();
          if (!textRange) return;
          
          const text = textRange.asString();
          if (text.includes(PLACEHOLDERS.SERMON)) {
            textRange.replaceAllText(PLACEHOLDERS.SERMON, sermonTitle || '');
          }
        } catch (error) {
          // Skip
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

    const entities = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: ' ' };
    return text.replace(/&([a-zA-Z]+);/g, (match, entity) => entities[entity] || match);
  } catch (error) {
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
      let verseHtml = pTag.replace(/<a[^>]*>.*?<\/a>/g, '')
                         .replace(/<br\s*\/?>/gi, '||LINEBREAK||')
                         .replace(/<\/?[^>]+(>|$)/g, "")
                         .trim();
      
      const decodedVerse = decodeHtmlEntities(verseHtml);
      
      if (decodedVerse && decodedVerse.trim()) {
        const cleanedVerse = decodedVerse.replace(/\|\|LINEBREAK\|\|/g, '\n')
                                        .replace(/\n\s*\n/g, '\n')
                                        .replace(/^\s+|\s+$/g, '')
                                        .replace(/[ \t]+/g, ' ');
        
        if (cleanedVerse.toLowerCase().includes("refrain")) {
          refrain = cleanedVerse;
        } else if (cleanedVerse.length > 0) {
          verses.push(cleanedVerse);
        }
      }
    });

    return { verses, refrain };
  } catch (error) {
    return { verses: [], refrain: "" };
  }
}

/**
 * Searches Gmail for praise/worship lyrics
 */
function searchGmailForPraiseLyrics() {
  try {
    const tenDaysAgo = new Date();
    tenDaysAgo.setDate(tenDaysAgo.getDate() - 5);
    const dateString = Utilities.formatDate(tenDaysAgo, Session.getScriptTimeZone(), 'yyyy/MM/dd');
    
    const searchQuery = `after:${dateString} (subject:(praise lyrics) OR subject:(worship lyrics))`;
    const threads = GmailApp.search(searchQuery, 0, 5);
    
    if (threads.length === 0) {
      return null;
    }
    
    const mostRecentThread = threads[0];
    const message = mostRecentThread.getMessages()[0];
    
    let emailBody = message.getBody();
    
    if (!emailBody || emailBody.trim() === '') {
      emailBody = message.getPlainBody();
      
      const lines = emailBody.split('\n').map(line => line.trim()).filter(line => line !== '');
      
      if (lines.length < 2) return null;
      
      const songTitle = lines[0];
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
      
      return { title: songTitle, lyrics: lyrics, subject: message.getSubject(), date: message.getDate() };
    }
    
    const divPattern = /<div[^>]*>(.*?)<\/div>/gi;
    const divContents = [];
    let match;
    
    while ((match = divPattern.exec(emailBody)) !== null) {
      let content = match[1];
      
      if (content.match(/^\s*<br\s*\/?>\s*$/)) {
        divContents.push('||BREAK||');
      } else {
        content = content.replace(/<br\s*\/?>/gi, ' ').replace(/<[^>]*>/g, '').trim();
        
        if (content !== '') {
          content = content.replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/&lt;/g, '<')
                          .replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#39;/g, "'")
                          .replace(/&#\d+;/g, '').replace(/=E2=80=99/g, "'")
                          .replace(/=\r?\n/g, '').replace(/=[0-9A-F]{2}/gi, '');
          divContents.push(content);
        }
      }
    }
    
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
    
    if (currentSection.length > 0) {
      sections.push(currentSection.join('\n'));
    }
    
    if (sections.length >= 2) {
      return { title: sections[0], lyrics: sections.slice(1), subject: message.getSubject(), date: message.getDate() };
    }
    
    if (divContents.length > 0) {
      const contentOnly = divContents.filter(item => item !== '||BREAK||');
      if (contentOnly.length < 2) return null;
      
      const songTitle = contentOnly[0];
      const remainingLines = contentOnly.slice(1);
      const lyrics = [];
      
      if (remainingLines.length >= 14) {
        lyrics.push(remainingLines.slice(0, 2).join('\n'));
        lyrics.push(remainingLines.slice(2, 6).join('\n'));
        lyrics.push(remainingLines.slice(6, 10).join('\n'));
        lyrics.push(remainingLines.slice(10, 14).join('\n'));
      } else {
        let currentVerse = [];
        for (let i = 0; i < remainingLines.length; i++) {
          const line = remainingLines[i];
          currentVerse.push(line);
          if (line.match(/(silver|gold|will|within|sin)$/i) || (currentVerse.length === 4) || (currentVerse.length === 2 && i < 4)) {
            lyrics.push(currentVerse.join('\n'));
            currentVerse = [];
          }
        }
        if (currentVerse.length > 0) {
          lyrics.push(currentVerse.join('\n'));
        }
      }
      
      return { title: songTitle, lyrics: lyrics, subject: message.getSubject(), date: message.getDate() };
    }
    
    return null;
  } catch (error) {
    Logger.log('Error searching Gmail: ' + error.toString());
    return null;
  }
}

/**
 * Updates praise song slides in presentation
 */
function updatePraiseSongSlides(presentation, praiseData) {
  if (!praiseData) return;
  
  try {
    const slides = presentation.getSlides();
    
    slides.forEach((slide) => {
      const shapes = slide.getShapes();
      shapes.forEach(shape => {
        try {
          const textRange = shape.getText();
          if (!textRange) return;
          const text = textRange.asString();
          if (text.includes(PLACEHOLDERS.PRAISE_SONG)) {
            textRange.replaceAllText(PLACEHOLDERS.PRAISE_SONG, praiseData.title);
          }
        } catch (error) {
          // Skip
        }
      });
    });
    
    let templateSlide = null;
    
    for (let i = 0; i < slides.length; i++) {
      const slide = slides[i];
      const shapes = slide.getShapes();
      for (let shape of shapes) {
        try {
          const text = shape.getText().asString();
          if (text.includes(PLACEHOLDERS.PRAISE_LYRICS)) {
            templateSlide = slide;
            break;
          }
        } catch (error) {
          // Skip
        }
      }
      if (templateSlide) break;
    }
    
    if (!templateSlide) {
      for (let i = 0; i < slides.length; i++) {
        const slide = slides[i];
        const shapes = slide.getShapes();
        for (let shape of shapes) {
          try {
            const text = shape.getText().asString();
            if (text.includes(praiseData.title)) {
              templateSlide = slide;
              const firstVerse = praiseData.lyrics[0];
              if (firstVerse) {
                shape.getText().setText(firstVerse);
                adjustFontSizeToFitShape(shape, firstVerse);
                for (let j = 1; j < praiseData.lyrics.length; j++) {
                  const newSlide = slide.duplicate();
                  const newShapes = newSlide.getShapes();
                  for (let newShape of newShapes) {
                    try {
                      const textRange = newShape.getText();
                      if (textRange && textRange.asString().includes(firstVerse)) {
                        textRange.setText(praiseData.lyrics[j]);
                        adjustFontSizeToFitShape(newShape, praiseData.lyrics[j]);
                        break;
                      }
                    } catch (error) {
                      // Continue
                    }
                  }
                }
                return;
              }
              break;
            }
          } catch (error) {
            // Skip
          }
        }
        if (templateSlide) break;
      }
      if (!templateSlide) return;
    }
    
    const validParagraphs = praiseData.lyrics.filter(para => para && para.trim() !== '');
    if (validParagraphs.length === 0) return;
    
    const shapes = templateSlide.getShapes();
    shapes.forEach(shape => {
      try {
        const textRange = shape.getText();
        if (!textRange) return;
        const text = textRange.asString();
        if (text.includes(PLACEHOLDERS.PRAISE_LYRICS)) {
          textRange.replaceAllText(PLACEHOLDERS.PRAISE_LYRICS, validParagraphs[0].trim());
          adjustFontSizeToFitShape(shape, validParagraphs[0].trim());
        }
      } catch (error) {
        // Skip
      }
    });
    
    for (let i = 1; i < validParagraphs.length; i++) {
      const newSlide = templateSlide.duplicate();
      const newShapes = newSlide.getShapes();
      newShapes.forEach(shape => {
        try {
          const textRange = shape.getText();
          if (!textRange) return;
          const text = textRange.asString();
          if (text.includes(validParagraphs[0].trim())) {
            textRange.setText(validParagraphs[i].trim());
            adjustFontSizeToFitShape(shape, validParagraphs[i].trim());
          }
        } catch (error) {
          // Skip
        }
      });
    }
  } catch (error) {
    Logger.log('Error in updatePraiseSongSlides: ' + error.toString());
  }
}
