/**
 * ╔══════════════════════════════════════════════════════════╗
 * ║       INSURANCEDEKHO PARTNER DASHBOARD - CONFIG          ║
 * ║   Paste your Google Apps Script Web App URL below        ║
 * ╚══════════════════════════════════════════════════════════╝
 *
 * STEP:
 * 1. Deploy Code.gs as a Web App (see instructions inside Code.gs)
 * 2. Copy the URL that looks like:
 *    https://script.google.com/macros/s/AKfyc.../exec
 * 3. Paste it below replacing the placeholder
 */

const DASHBOARD_CONFIG = {
  // ← PASTE YOUR WEB APP URL HERE
  API_URL: "https://script.google.com/macros/s/AKfycbwuYGbPRvNIOFRMHO2C1emPfmwHVIjrVb57ask8goT6b4Yr4hDWwv9V9hHFqEUjMkSmZQ/exec",

  // Dashboard branding
  COMPANY_NAME: "InsuranceDekho",
  DASHBOARD_TITLE: "Partner Performance Dashboard",
  SUBTITLE: "All Partners · Live Data",

  // Auto-refresh interval in seconds (0 = disabled)
  AUTO_REFRESH_SECONDS: 120,

  // Default rows per page in the table
  DEFAULT_PAGE_SIZE: 25,

  // Month column keys (must match your sheet headers exactly)
  MONTH_COLS: [
    "April'2025", "May'25", "June'25", "July'25", "August'25",
    "September'25", "October'25", "November'25", "December'25",
    "January'26", "Feburary'26", "March'26", "April'26"
  ],

  // Short labels shown in chart X axis
  MONTH_LABELS: [
    "Apr'25", "May'25", "Jun'25", "Jul'25", "Aug'25",
    "Sep'25", "Oct'25", "Nov'25", "Dec'25",
    "Jan'26", "Feb'26", "Mar'26", "Apr'26"
  ],

  // Columns that should always be visible in the table
  PINNED_COLS: ["S.No", "GID/GCD", "NAME", "CITY", "STATE", "Zone", "Owner"],

  // Zones in your data
  ZONES: ["All", "North", "East & Central", "West", "South", "RON", "North East"],
};
