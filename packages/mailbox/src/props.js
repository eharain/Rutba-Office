// MAPI property tags and their types.
//
// Outlook does not store a message as RFC 5322. It stores a bag of numbered
// properties — subject is 0x0037, the HTML body is 0x1013 — and the same
// numbering is used by .msg files, by .pst files and over MAPI itself. So one
// table serves all three readers.
//
// The type is the low 16 bits of the tag and decides how to read the value; the
// id is the high 16 bits and decides what it means.

export const PT = {
  UNSPECIFIED: 0x0000,
  NULL: 0x0001,
  I2: 0x0002,
  LONG: 0x0003,
  R4: 0x0004,
  DOUBLE: 0x0005,
  CURRENCY: 0x0006,
  APPTIME: 0x0007,
  ERROR: 0x000a,
  BOOLEAN: 0x000b,
  OBJECT: 0x000d,
  I8: 0x0014,
  STRING8: 0x001e,
  UNICODE: 0x001f,
  SYSTIME: 0x0040,
  CLSID: 0x0048,
  BINARY: 0x0102,
  MV_STRING8: 0x101e,
  MV_UNICODE: 0x101f,
  MV_BINARY: 0x1102,
  MV_LONG: 0x1003,
};

/** The properties this suite actually reads, by id. */
export const PROPS = {
  0x0017: 'importance',
  0x001a: 'messageClass',
  0x0023: 'originatorDeliveryReportRequested',
  0x0026: 'priority',
  0x0029: 'readReceiptRequested',
  0x0037: 'subject',
  0x0039: 'clientSubmitTime',
  0x003b: 'sentRepresentingSearchKey',
  0x003d: 'subjectPrefix',
  0x003f: 'receivedByEntryId',
  0x0040: 'receivedByName',
  0x0042: 'sentRepresentingName',
  0x0044: 'receivedRepresentingName',
  0x0047: 'messageSubmissionId',
  0x004d: 'originalAuthorName',
  0x0064: 'sentRepresentingAddressType',
  0x0065: 'sentRepresentingEmailAddress',
  0x0070: 'conversationTopic',
  0x0071: 'conversationIndex',
  0x0072: 'originalDisplayBcc',
  0x0073: 'originalDisplayCc',
  0x0074: 'originalDisplayTo',
  0x0075: 'receivedByAddressType',
  0x0076: 'receivedByEmailAddress',
  0x007d: 'transportMessageHeaders',
  0x0c15: 'recipientType',
  0x0c1a: 'senderName',
  0x0c1e: 'senderAddressType',
  0x0c1f: 'senderEmailAddress',
  0x0e01: 'deleteAfterSubmit',
  0x0e02: 'displayBcc',
  0x0e03: 'displayCc',
  0x0e04: 'displayTo',
  0x0e06: 'messageDeliveryTime',
  0x0e07: 'messageFlags',
  0x0e08: 'messageSize',
  0x0e1d: 'normalizedSubject',
  0x0e20: 'attachSize',
  0x0e23: 'internetArticleNumber',
  0x0ff9: 'recordKey',
  0x1000: 'body',
  0x1009: 'rtfCompressed',
  0x1013: 'bodyHtml',
  0x1035: 'internetMessageId',
  0x1039: 'internetReferences',
  0x1042: 'inReplyToId',
  0x1080: 'iconIndex',
  0x1081: 'lastVerbExecuted',
  0x3001: 'displayName',
  0x3002: 'addressType',
  0x3003: 'emailAddress',
  0x3007: 'creationTime',
  0x3008: 'lastModificationTime',
  0x300b: 'searchKey',
  0x3701: 'attachData',
  0x3702: 'attachEncoding',
  0x3703: 'attachExtension',
  0x3704: 'attachFilename',
  0x3705: 'attachMethod',
  0x3707: 'attachLongFilename',
  0x370b: 'renderingPosition',
  0x370e: 'attachMimeTag',
  0x3712: 'attachContentId',
  0x3714: 'attachFlags',
  0x39fe: 'smtpAddress',
  0x39ff: 'displayNameSimple',
  0x3a00: 'account',
  0x3a06: 'givenName',
  0x3a11: 'surname',
  0x3a08: 'businessTelephoneNumber',
  0x3a1c: 'mobileTelephoneNumber',
  0x3a16: 'companyName',
  0x360a: 'hasSubfolders',
  0x3602: 'contentCount',
  0x3603: 'contentUnreadCount',
  0x35e0: 'ipmSubTreeEntryId',
  0x35e3: 'ipmWastebasketEntryId',
  0x35e7: 'finderEntryId',
  0x6633: 'pstHiddenCount',
  0x67f2: 'ltpRowId',
  0x67f3: 'ltpRowVer',
  0x5d01: 'senderSmtpAddress',
  0x5d02: 'sentRepresentingSmtpAddress',
  0x5ff6: 'recipientDisplayName',
  0x5fde: 'recipientResourceState',
  0x0c19: 'senderEntryId',
  0x0e28: 'primarySendAccount',
  0x0e29: 'nextSendAcct',
  0x8005: 'fileUnder',
};

export const NAME_TO_ID = Object.fromEntries(Object.entries(PROPS).map(([id, name]) => [name, Number(id)]));

/** Message flags — bit 0 is "read", which every mail list needs. */
export const MSG_FLAG = {
  READ: 0x01,
  UNMODIFIED: 0x02,
  SUBMIT: 0x04,
  UNSENT: 0x08,
  HAS_ATTACH: 0x10,
  FROM_ME: 0x20,
  ASSOCIATED: 0x40,
  RESEND: 0x80,
};

/** Recipient types, from PidTagRecipientType. */
export const RECIPIENT_TYPE = { 0: 'orig', 1: 'to', 2: 'cc', 3: 'bcc' };

/** FILETIME (100 ns ticks since 1601) as an ISO string. */
export function filetimeToIso(low, high) {
  const ticks = high * 2 ** 32 + low;
  if (!ticks) return null;
  const ms = ticks / 10000 - 11644473600000;
  if (!Number.isFinite(ms) || ms < -62135596800000 || ms > 253402300799999) return null;
  return new Date(ms).toISOString();
}

/**
 * Decode one property value from its bytes.
 * @param {number} type PT_* constant
 * @param {Uint8Array} bytes
 * @param {boolean} unicode whether strings are UTF-16LE
 */
export function decodeValue(type, bytes, unicode = true) {
  if (!bytes) return null;
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  switch (type) {
    case PT.I2:
      return bytes.length >= 2 ? dv.getInt16(0, true) : null;
    case PT.LONG:
    case PT.ERROR:
      return bytes.length >= 4 ? dv.getInt32(0, true) : null;
    case PT.R4:
      return bytes.length >= 4 ? dv.getFloat32(0, true) : null;
    case PT.DOUBLE:
    case PT.APPTIME:
      return bytes.length >= 8 ? dv.getFloat64(0, true) : null;
    case PT.BOOLEAN:
      return bytes.length >= 1 ? bytes[0] !== 0 : null;
    case PT.I8:
    case PT.CURRENCY:
      return bytes.length >= 8 ? Number(dv.getBigInt64(0, true)) : null;
    case PT.SYSTIME:
      return bytes.length >= 8 ? filetimeToIso(dv.getUint32(0, true), dv.getUint32(4, true)) : null;
    case PT.UNICODE:
      return new TextDecoder('utf-16le').decode(trimNul(bytes, 2));
    case PT.STRING8:
      return new TextDecoder(unicode ? 'utf-8' : 'windows-1252', { fatal: false }).decode(trimNul(bytes, 1));
    case PT.CLSID:
      return [...bytes].map((b) => b.toString(16).padStart(2, '0')).join('');
    case PT.BINARY:
    case PT.OBJECT:
      return bytes;
    default:
      return bytes;
  }
}

function trimNul(bytes, width) {
  let end = bytes.length;
  if (width === 2) {
    while (end >= 2 && bytes[end - 1] === 0 && bytes[end - 2] === 0) end -= 2;
  } else {
    while (end >= 1 && bytes[end - 1] === 0) end -= 1;
  }
  return bytes.subarray(0, end);
}

/** A tag as `0x0037001f` → { id, type, name }. */
export function describeTag(tag) {
  const id = (tag >>> 16) & 0xffff;
  const type = tag & 0xffff;
  return { id, type, name: PROPS[id] || `0x${id.toString(16).padStart(4, '0')}` };
}
