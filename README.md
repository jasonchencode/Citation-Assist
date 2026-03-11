# Word Citation Highlight Add-in

This project is a Word Office.js task pane add-in that lets you analyze selected text in a document, get citation metadata from a backend API, and visually track those citations with synchronized highlights and comments in the document.

## What the add-in does

- **Analyze selected text**: When you select text in a Word document and click **Analyze Selection** in the task pane, the add-in:
  - Reads the current selection using the **Word JavaScript API**.
  - Sends the text (along with a fixed `document_id` and `user_id`) to a backend **`POST /api/analyze`** endpoint.
  - Receives a citation response containing `source_id`, `citation_text`, `confidence`, and `url`.

- **Highlight and tag in Word**:
  - On a successful response with sufficient confidence, the add-in applies a highlight color to the selected text in the document.
  - Each analyzed selection uses the next color from a fixed palette, so different citations are visually distinguishable.

- **Task pane citation list**:
  - Every successful analysis is added to a **scrollable citation list** in the task pane.
  - Each card shows:
    - The **citation text**
    - **Confidence** as a percentage
    - An optional **Source** link (if a URL is present)
    - A **color dot** that matches the highlight color applied in the document
  - You can:
    - Click **Re-select** to jump Word’s selection back to the original passage.
    - Click **Remove** to remove the card and clear the corresponding highlight and comments from the document.

- **Error handling and retries**:
  - If no text is selected, the task pane shows a **“No text selected”** inline message instead of calling the API.
  - If the API returns an error, the task pane shows the error message inline and provides a **Retry** button so you can re-run the request without re-clicking Analyze.
  - If the backend returns a null or low-confidence result, the UI shows **“Unable to generate citation. Try refining your selection.”** instead of adding a citation card.

- **Sync with document changes**:
  - A **Refresh** button re-checks each citation’s text against the current document using `checkCitationExists` and removes cards whose text no longer exists (for example, if you deleted the highlighted text in Word).

## Running the project

### Prerequisites

- **Node.js** (LTS) installed (`node -v`, `npm -v` should work).
- **Word** as part of a Microsoft 365 subscription, or a compatible perpetual version.

### Start the dev server

From the project root:

```bash
npm install
npm run dev-server
```

This starts `webpack-dev-server` on `https://localhost:3000` with a development build.

### Sideload the add-in into Word

Using the **Office Add-ins Development Kit** (recommended):

1. Open this folder in VS Code / Cursor with the Office Add-ins Development Kit extension installed.
2. In the extension’s view, choose **Preview Your Office Add-in (F5)**.
3. When prompted, choose **Word Desktop (Edge Chromium)**.
4. Word will launch and sideload the add-in using `manifest.xml`.

Alternatively, you can use the `office-addin-debugging` scripts directly:

```bash
npm run start   # Starts dev server (if not already) and sideloads the add-in
npm run stop    # Stops debugging and removes the sideloaded add-in
```

## Customizing for your backend

By default, the project expects a backend reachable via `/api/analyze` (proxied from `webpack-dev-server` to your actual URL). To hook this up:

1. Configure your dev proxy or hosting so `/api/*` calls from `https://localhost:3000` reach your backend.
2. Ensure your backend implements:
   - `POST /api/analyze` with body `{ text, document_id, user_id }`.
   - Returns `{ source_id, citation_text, confidence, url }` or an error payload with an `error` string.
3. Add CORS rules allowing `https://localhost:3000`.
