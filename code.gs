/*******************************************************
 * Configuration
 *******************************************************/
const PHAB_BASE = "https://phabricator.wikimedia.org/api/maniphest.search";

/*******************************************************
 * Menu
 *******************************************************/
function buildHomepage() {
  return CardService.newCardBuilder()
    .setHeader(
      CardService.newCardHeader()
        .setTitle("Phabricator Link Updater")
    )
    .addSection(
      CardService.newCardSection()
        .addWidget(
          CardService.newTextParagraph()
            .setText("Update a Phabricator link at your cursor.")
        )
        .addWidget(
          CardService.newTextButton()
            .setText("Update link at cursor")
            .setOnClickAction(
              CardService.newAction()
                .setFunctionName("handleUpdateLink")
            )
        )
    )
    .build();
}


function handleUpdateLink(e) {
  try {
    updateLinkUnderCursor();

    return CardService.newActionResponseBuilder()
      .setNotification(
        CardService.newNotification()
          .setText("Link updated successfully ✅")
      )
      .build();

  } catch (err) {
    return CardService.newActionResponseBuilder()
      .setNotification(
        CardService.newNotification()
          .setText("Error: " + err.message)
      )
      .build();
  }
}


/*******************************************************
 * Update link under cursor
 *******************************************************/
function updateLinkUnderCursor() {
  const doc = DocumentApp.getActiveDocument();
  const cursor = doc.getCursor();

  if (!cursor) {
    throw new Error("Place your cursor inside a link.");
  }

  const element = cursor.getElement();
  if (!element.editAsText) {
    throw new Error("Cursor is not inside editable text.");
  }

  const text = element.asText();
  const offset = cursor.getOffset();
  const url = text.getLinkUrl(offset);

  if (!url) {
    throw new Error("Cursor is not inside a link.");
  }

  processSingleLink(text, offset, url);
}


/*******************************************************
 * Replace one hyperlink safely
 *******************************************************/
function processSingleLink(text, offset, url) {
  const match = url.match(/\/T(\d+)/);
  if (!match) return;

  const taskId = match[1];
  const title = fetchPhabricatorTitle(taskId);
  if (!title) return;

  // Find link boundaries
  let start = offset;
  while (start > 0 && text.getLinkUrl(start - 1) === url) {
    start--;
  }

  let end = offset;
  while (end < text.getText().length - 1 &&
         text.getLinkUrl(end + 1) === url) {
    end++;
  }

  text.deleteText(start, end);
  text.insertText(start, title);
  text.setLinkUrl(start, start + title.length - 1, url);
}

/*******************************************************
 * Fetch title from Phabricator
 *******************************************************/
function fetchPhabricatorTitle(taskId) {
  try {
    const url = `https://phabricator.wikimedia.org/T${taskId}`;
    const response = UrlFetchApp.fetch(url);
    const html = response.getContentText();

    const match = html.match(/<title>(.*?)<\/title>/i);
    if (!match) return null;

    return match[1]
      .replace(/^T\d+\s*/, "")
      .replace(/\s*·\s*Wikimedia Phabricator$/, "")
      .trim();

  } catch (e) {
    Logger.log(e);
    return null;
  }
}
