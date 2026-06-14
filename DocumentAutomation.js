function exportFilesSmartLinkSupport() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const doc = DocumentApp.create("DITMAWA");
  const body = doc.getBody();

  const startRow = 2; 
  const lastRow = 50;

  // Columns to read: label => column index (1-based)
  const columns = {
    J: 10,
    K: 11,
    L: 12,
    M: 13,
  };

  // Target size inside Google Docs (points)
  const MAX_WIDTH = 468;
  const MAX_HEIGHT = 648;

  // Insert a page break before each new file after the first one
  let hasAddedASection = false;

  for (let row = startRow; row <= lastRow; row++) {
    for (const [label, col] of Object.entries(columns)) {
      const cell = sheet.getRange(row, col);
      const displayValue = cell.getValue();
      const richText = cell.getRichTextValue();
      const linkUrl = richText ? richText.getLinkUrl() : null;

      let file = null;
      const fileName = displayValue ? String(displayValue).trim() : "";

      try {
        // 1) Resolve Drive file either from link or by name
        if (linkUrl) {
          Logger.log(`${label}${row}: Link URL detected: ${linkUrl}`);
          const fileId = extractDriveFileId_(linkUrl);
          if (fileId) {
            file = DriveApp.getFileById(fileId);
          } else {
            // Start a section for the notice (each notice still takes one page)
            if (hasAddedASection) body.appendPageBreak();
            body.appendParagraph(`Row ${row} — ${label}: ⚠️ Invalid Drive link`).setForegroundColor("#b26a00");
            hasAddedASection = true;
            continue;
          }
        } else if (fileName) {
          Logger.log(`${label}${row}: Looking by name "${fileName}"`);
          const files = DriveApp.getFilesByName(fileName);
          if (files.hasNext()) {
            file = files.next();
          } else {
            if (hasAddedASection) body.appendPageBreak();
            body.appendParagraph(`Row ${row} — ${label}: ❌ File not found: "${fileName}"`).setForegroundColor("#b00020");
            hasAddedASection = true;
            continue;
          }
        } else {
          // Skip empty cells silently
          continue;
        }

        // 2) New page per *document/file*
        if (hasAddedASection) body.appendPageBreak();

        // 3) Section header
        body.appendParagraph(`📦 Row ${row} — ${label}`).setHeading(DocumentApp.ParagraphHeading.HEADING3);

        // 4) Insert content based on MIME type
        const mime = file.getMimeType();

        if (mime.startsWith("image/")) {
          // Special handling for HEIC/HEIF which Docs can't embed directly
          if (mime === "image/heic" || mime === "image/heif") {
            try {
              const jpegBlob = file.getBlob().getAs("image/jpeg");
              const img = body.appendImage(jpegBlob);
              resizeImage_(img, MAX_WIDTH, MAX_HEIGHT);
              body.appendParagraph(`(Converted from HEIC/HEIF → JPEG: ${file.getName()})`).setFontSize(9).setItalic(true);
            } catch (err) {
              body.appendParagraph(`📸 HEIC/HEIF file cannot be embedded. Open in Drive:`)
                  .appendText(` ${file.getUrl()}`);
            }
          } else {
            const img = body.appendImage(file.getBlob());
            resizeImage_(img, MAX_WIDTH, MAX_HEIGHT);
          }

        } else if (mime === MimeType.PDF) {
          // Docs can't render PDF pages inline; leave a note + link
          body.appendParagraph(`📄 PDF File - ${file.getName()}`);
          body.appendParagraph(`(Open in Drive: ${file.getUrl()})`).setFontSize(9).setItalic(true);

        } else {
          body.appendParagraph(`⚠️ Unsupported type (${mime}) - ${file.getName()}`).setForegroundColor("#b26a00");
          body.appendParagraph(`(Open in Drive: ${file.getUrl()})`).setFontSize(9).setItalic(true);
        }

        hasAddedASection = true;

      } catch (e) {
        if (hasAddedASection) body.appendPageBreak();
        body.appendParagraph(`Row ${row} — ${label}: Error - ${e.message}`).setForegroundColor("#b00020");
        hasAddedASection = true;
      }
    }
  }

  Logger.log("✅ Document created: " + doc.getUrl());
}

/**
 * Resize image to fit max width and height (points).
 * Keeps aspect ratio; never upsizes beyond original.
 */
function resizeImage_(image, maxWidth, maxHeight) {
  const width = image.getWidth();
  const height = image.getHeight();
  const ratio = Math.min(maxWidth / width, maxHeight / height, 1);
  image.setWidth(Math.round(width * ratio));
  image.setHeight(Math.round(height * ratio));
}

/**
 * Extract a Google Drive fileId from common link shapes:
 * - https://drive.google.com/file/d/<ID>/view
 * - https://drive.google.com/open?id=<ID>
 * - https://drive.google.com/uc?id=<ID>
 * - Any URL containing a long Drive-like id
 */
function extractDriveFileId_(url) {
  if (!url) return null;

  // Try specific patterns first
  let m = url.match(/\/d\/([-\w]{25,})/);
  if (m && m[1]) return m[1];

  m = url.match(/[?&]id=([-\w]{25,})/);
  if (m && m[1]) return m[1];

  // Fallback: any 25+ char Drive-like id in the URL
  m = url.match(/[-\w]{25,}/);
  if (m && m[0]) return m[0];

  return null;
}
