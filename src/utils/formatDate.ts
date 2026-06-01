/**
 * Shared date formatting utilities — used across all components.
 * Uses Intl.DateTimeFormat for locale-aware, consistent output.
 */

const DATE_FMT = new Intl.DateTimeFormat(undefined, {
  year: 'numeric',
  month: 'short',
  day: 'numeric',
});

const DATETIME_FMT = new Intl.DateTimeFormat(undefined, {
  year: 'numeric',
  month: 'short',
  day: 'numeric',
  hour: 'numeric',
  minute: '2-digit',
});

/** Format an ISO string as a short date: "May 30, 2026" */
export function formatDate(value?: string | null): string {
  if (!value) return '—';
  const d = new Date(value);
  return Number.isNaN(d.getTime()) ? value : DATE_FMT.format(d);
}

/** Format an ISO string as date + time: "May 30, 2026, 11:42 PM" */
export function formatDateTime(value?: string | null): string {
  if (!value) return '—';
  const d = new Date(value);
  return Number.isNaN(d.getTime()) ? value : DATETIME_FMT.format(d);
}
