# Parkhi website

Upload everything in this folder to the root of the GitHub repo (yashikasinghal13/Parkhi).

- `index.html` – the whole website (shop, product pages, size and fit picker, bag, checkout, reviews)
- `img/` – kurta photos and the UPI QR code
- `backend/Code.gs` – the Google Apps Script that saves orders, customers and reviews to the Google Sheet.
  It is kept here only as a backup copy; the live copy runs inside the Google Sheet (Extensions > Apps Script).
- `.nojekyll` – tells GitHub Pages to serve the files as they are

GitHub Pages: repo Settings > Pages > Source "Deploy from a branch", Branch "main", folder "/ (root)".

The site's connection to the Google Sheet is the `backendUrl` line near "shop settings" in `index.html`.
After changing `backend/Code.gs` in Apps Script: Deploy > Manage deployments > Edit > New version > Deploy (the URL stays the same).
