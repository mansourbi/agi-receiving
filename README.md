# AGI Receiving

Standalone QR receiving app for AGI Glass Factory. It talks only to a Google Apps Script web app (which writes to the Google Sheet) - it is NOT connected to the AGI server.

- index.html : the app (a self-contained PWA; add to home screen on the phone)
- Code.gs    : the Google Apps Script backend

Published with GitHub Pages. Managed by Deploy-AgiReceiving.ps1 - re-run it to publish updates.