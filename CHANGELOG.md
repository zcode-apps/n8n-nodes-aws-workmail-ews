# Changelog

All notable changes to this project will be documented in this file.

## [1.0.14] - 2026-01-20

### Added
- **Create Reply Draft**: Neue Operation zum Erstellen von Antwort-Entwuerfen. Die Antwort wird als Entwurf im Drafts-Ordner gespeichert und kann vor dem Senden ueberprueft werden. Ideal fuer AI-Agent-Workflows, die E-Mail-Antworten vorbereiten.

### Fixed
- **Email Body leer bei externen Mails**: Behoben, dass der Body bei HTML-Mails (z.B. von externen Absendern) leer war. Jetzt wird `RequestedBodyType = BodyType.Text` gesetzt, um den Inhalt zuverlaessig zu erhalten.
- **Body-Extraktion verbessert**: Robustere Extraktion des Body-Inhalts mit mehreren Fallback-Methoden (Body.Text, Body.text, Body.Content, toString).

### Technical Details
- `getMessage()` und `getMessages()` setzen jetzt explizit `propertySet.RequestedBodyType = ews.BodyType.Text`
- `convertMessageToJson()` versucht mehrere Wege, den Body-Text zu extrahieren
- BodyType wird korrekt von numerischen Werten (0=HTML, 1=Text) in Strings konvertiert
- Neue `createReplyDraft()` Methode im EwsClient

## [1.0.5] - 2026-01-20

### Fixed
- **Email Body Empty**: Fixed critical bug where email bodies were always empty in "Get Many Messages" operation. Messages are now loaded with full Body content using `EmailMessage.Bind`.
- **Trigger Fires Immediately**: Fixed trigger firing immediately with old emails on first workflow start. Trigger now starts from 5 minutes ago on first poll instead of 1 hour ago, preventing old emails from triggering the workflow.

### Changed
- `getMessages()` now uses `Bind` to load full message details including Body
- Trigger default lookback period reduced from 1 hour to 5 minutes on first poll

### Technical Details
- Added explicit `EmailMessage.Bind` call in `getMessages()` with PropertySet including Body schema
- Improved `convertMessageToJson` to properly extract Body.Text content
- Modified trigger polling logic to avoid triggering with historical emails

## [1.0.4] - 2026-01-20

### Fixed
- **Get Many Messages**: Fixed "You must load or assign this property before you can read its value" error when retrieving multiple messages. Body property is now safely accessed with proper error handling.
- **Get Many Calendars**: Fixed "Id is invalid" error. Now correctly returns the default calendar instead of trying to find sub-calendars (AWS WorkMail has one calendar per mailbox).

### Technical Details
- Added try-catch block in `convertMessageToJson` to handle properties that may not be loaded in FindItems results
- Changed `getCalendars()` to directly bind to the default Calendar folder instead of using FindFolders

## [1.0.3] - 2026-01-20

### Changed
- Switched from `node-ews` to `ews-javascript-api` library due to WSDL file requirements
- Updated EWS endpoint URL format from `mobile.mail` to `ews.mail`
- Set correct Exchange version to Exchange 2010 SP2 (AWS WorkMail standard)

### Fixed
- All TypeScript compilation errors resolved
- Fixed CommonJS module imports
- Fixed all CRUD operations to work with AWS WorkMail
- Contact email address API (`_setItem` instead of `set_Item`)
- Attachment collection iteration (`_getItem` instead of `getPropertyAtIndex`)
- Message/Event/Contact delete operations (use instance methods instead of service methods)

### Tested
- ✅ 100% operation success rate (25/25 operations)
- ✅ All CRUD operations validated with real AWS WorkMail instance

## [1.0.2] - 2026-01-20

### Fixed
- Added all missing dependencies explicitly (httpntlm, soap, when, request)

## [1.0.1] - 2026-01-20

### Fixed
- Added missing ntlm-client dependency

## [1.0.0] - 2026-01-20

### Added
- Initial release
- Support for Messages, Folders, Calendars, Events, Contacts, Attachments
- Basic Authentication with AWS WorkMail
- Polling trigger for new messages
- Binary data support for attachments
- Comprehensive error handling
