# Primary Market Dashboard

This repository contains:

- `index.html`: published GitHub Pages entry
- `primary_market_dashboard.html`: generated dashboard output
- `primary_market_dashboard_data.json`: exported data payload
- `build_primary_market_dashboard.py`: source code used to build the dashboard
- `secondary_market_dashboard.html`: generated Hanoi secondary market dashboard
- `secondary_market_dashboard_data.json`: exported secondary market data payload
- `build_secondary_market_dashboard.py`: source code used to build the secondary dashboard
- `secondary_report_dashboard.html`: generated Hanoi secondary report chart dashboard
- `secondary_report_dashboard_data.json`: exported secondary report chart data payload
- `build_secondary_report_dashboard.py`: source code used to build the secondary report dashboard

## Vercel Protected Deploy

This folder can also be deployed to Vercel with a shared-password login wall.

Required environment variables:

- `SITE_PASSWORD`: shared password for viewers
- `SESSION_SECRET`: long random string used to sign the login session cookie

Login route:

- `/login`

Logout route:

- `/api/logout`
