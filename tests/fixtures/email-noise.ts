// Reusable, composable synthetic email-noise fixtures for the Email Noise Torture Test
// (docs/AI_HANDOFF.md "Email Noise Torture Test"). Entirely fictional names/organizations —
// no real personal or confidential correspondence. Small composable blocks, combined by
// buildEmail(), rather than one giant repeated blob per test.

export const EXTERNAL_SENDER_WARNING = `CAUTION: This email originated from outside the organization.
Do not click links or open attachments unless you recognize the sender.`;

export const SIGNATURE_BLOCK = `Jordan Smith
Director of Curriculum
Example Area School District
123 Main Street
Example, PA 17901
Phone: 570-555-1234
Fax: 570-555-5678
www.example.org`;

export const CONFIDENTIALITY_NOTICE = `This email and any attachments may contain confidential or privileged information intended solely for the named recipient. If you received this in error, please notify the sender and delete this message.`;

export const TEAMS_BOILERPLATE = `________________________________________
Join Microsoft Teams Meeting
Meeting ID: 234 567 890 12
Passcode: aB3xY9
Dial-in: +1 570-555-9000
Learn more about Teams | Meeting options`;

export const ZOOM_BOILERPLATE = `Join Zoom Meeting
https://zoom.us/j/1234567890
Meeting ID: 123 456 7890
Passcode: 654321
One tap mobile: +13126266799,,1234567890#`;

export const SOCIAL_LINKS = `Follow us: Facebook | Instagram | LinkedIn | YouTube`;

export const UNSUBSCRIBE_FOOTER = `Manage preferences
Unsubscribe
View in browser`;

export const IMAGE_PLACEHOLDERS = `[image]
[cid:image001.png]
[logo]
[EXTERNAL IMAGE]`;

export function autoReplyBody(returnDate: string): string {
  return `Thank you for your email. I am out of the office until ${returnDate} and will respond when I return.`;
}

export const CALENDAR_BLOCK = `When: Tuesday, September 15, 2026 2:00 PM
Where: Microsoft Teams Meeting
Organizer: Pat Alvarez
Required attendees: Sam Rivera; Jordan Smith`;

export function forwardedHeader(from: string, sent: string, to: string, subject: string): string {
  return `From: ${from}
Sent: ${sent}
To: ${to}
Subject: ${subject}`;
}

export const LEGAL_SECURITY_FOOTER = `This message was scanned by Example Secure Email Gateway. Do not forward outside the organization. Data Classification: Internal Use Only.`;

/** Wraps a current message with a quoted "-----Original Message-----" block, matching the common Outlook/Exchange shape. */
export function withQuotedThread(current: string, quotedSenderName: string, quotedDate: string, quotedBody: string): string {
  return `${current}

-----Original Message-----
From: ${quotedSenderName}
Sent: ${quotedDate}
${quotedBody}`;
}

/** Joins any number of email parts (paragraphs/blocks) with blank-line separation, skipping falsy parts. */
export function buildEmail(...parts: (string | false | undefined | null)[]): string {
  return parts.filter((part): part is string => Boolean(part)).join("\n\n");
}
