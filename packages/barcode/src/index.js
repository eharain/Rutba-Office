// QR Codes: the encoder the windows draw from. The reader (./decode) and the
// PNG reader (./png-read) are for the tests and the window checks.
export { encodeQr, LEVELS, MODES, modeFor, dataCodewords, formatBits, versionBits, reedSolomon, alignmentPositions } from './qr.js';
