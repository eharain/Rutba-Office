// The object identifiers S/MIME is written with, by name.

export const OID = {
  // CMS content types (RFC 5652, 5083)
  data: '1.2.840.113549.1.7.1',
  signedData: '1.2.840.113549.1.7.2',
  envelopedData: '1.2.840.113549.1.7.3',
  encryptedData: '1.2.840.113549.1.7.6',
  authEnvelopedData: '1.2.840.113549.1.9.16.1.23',
  // signed attributes
  contentType: '1.2.840.113549.1.9.3',
  messageDigest: '1.2.840.113549.1.9.4',
  signingTime: '1.2.840.113549.1.9.5',
  smimeCapabilities: '1.2.840.113549.1.9.15',
  // digests
  sha1: '1.3.14.3.2.26',
  sha256: '2.16.840.1.101.3.4.2.1',
  sha384: '2.16.840.1.101.3.4.2.2',
  sha512: '2.16.840.1.101.3.4.2.3',
  // public keys and signatures
  rsaEncryption: '1.2.840.113549.1.1.1',
  rsaesOaep: '1.2.840.113549.1.1.7',
  mgf1: '1.2.840.113549.1.1.8',
  rsassaPss: '1.2.840.113549.1.1.10',
  sha1WithRSA: '1.2.840.113549.1.1.5',
  sha256WithRSA: '1.2.840.113549.1.1.11',
  sha384WithRSA: '1.2.840.113549.1.1.12',
  sha512WithRSA: '1.2.840.113549.1.1.13',
  ecPublicKey: '1.2.840.10045.2.1',
  ecdsaWithSHA1: '1.2.840.10045.4.1',
  ecdsaWithSHA256: '1.2.840.10045.4.3.2',
  ecdsaWithSHA384: '1.2.840.10045.4.3.3',
  ecdsaWithSHA512: '1.2.840.10045.4.3.4',
  // content encryption
  aes128CBC: '2.16.840.1.101.3.4.1.2',
  aes192CBC: '2.16.840.1.101.3.4.1.22',
  aes256CBC: '2.16.840.1.101.3.4.1.42',
  aes128GCM: '2.16.840.1.101.3.4.1.6',
  aes192GCM: '2.16.840.1.101.3.4.1.26',
  aes256GCM: '2.16.840.1.101.3.4.1.46',
  desEde3CBC: '1.2.840.113549.3.7',
  rc2CBC: '1.2.840.113549.3.2',
  // PKCS #12 (RFC 7292) and PKCS #5 (RFC 8018)
  keyBag: '1.2.840.113549.1.12.10.1.1',
  pkcs8ShroudedKeyBag: '1.2.840.113549.1.12.10.1.2',
  certBag: '1.2.840.113549.1.12.10.1.3',
  x509Certificate: '1.2.840.113549.1.9.22.1',
  friendlyName: '1.2.840.113549.1.9.20',
  localKeyId: '1.2.840.113549.1.9.21',
  pbeSHA1And3DES: '1.2.840.113549.1.12.1.3',
  pbeSHA1And2DES: '1.2.840.113549.1.12.1.4',
  pbeSHA1And128RC2: '1.2.840.113549.1.12.1.5',
  pbeSHA1And40RC2: '1.2.840.113549.1.12.1.6',
  pbes2: '1.2.840.113549.1.5.13',
  pbkdf2: '1.2.840.113549.1.5.12',
  hmacSHA1: '1.2.840.113549.2.7',
  hmacSHA256: '1.2.840.113549.2.9',
  hmacSHA384: '1.2.840.113549.2.10',
  hmacSHA512: '1.2.840.113549.2.11',
  // certificates
  commonName: '2.5.4.3',
  emailAddress: '1.2.840.113549.1.9.1',
  subjectKeyIdentifier: '2.5.29.14',
  keyUsage: '2.5.29.15',
  subjectAltName: '2.5.29.17',
  basicConstraints: '2.5.29.19',
  nameConstraints: '2.5.29.30',
  extKeyUsage: '2.5.29.37',
  emailProtection: '1.3.6.1.5.5.7.3.4',
  anyExtendedKeyUsage: '2.5.29.37.0',
};

/** Node's name for a digest's OID. */
export const DIGEST = { [OID.sha1]: 'sha1', [OID.sha256]: 'sha256', [OID.sha384]: 'sha384', [OID.sha512]: 'sha512' };
export const DIGEST_OID = { sha1: OID.sha1, sha256: OID.sha256, sha384: OID.sha384, sha512: OID.sha512 };

/** A signature algorithm's digest, where the algorithm names one. */
export const SIGNATURE_DIGEST = {
  [OID.sha1WithRSA]: 'sha1',
  [OID.sha256WithRSA]: 'sha256',
  [OID.sha384WithRSA]: 'sha384',
  [OID.sha512WithRSA]: 'sha512',
  [OID.ecdsaWithSHA1]: 'sha1',
  [OID.ecdsaWithSHA256]: 'sha256',
  [OID.ecdsaWithSHA384]: 'sha384',
  [OID.ecdsaWithSHA512]: 'sha512',
};

/** A content cipher: Node's name and key length in bytes. */
export const CIPHER = {
  [OID.aes128CBC]: { name: 'aes-128-cbc', key: 16, label: 'AES-128' },
  [OID.aes192CBC]: { name: 'aes-192-cbc', key: 24, label: 'AES-192' },
  [OID.aes256CBC]: { name: 'aes-256-cbc', key: 32, label: 'AES-256' },
  [OID.aes128GCM]: { name: 'aes-128-gcm', key: 16, label: 'AES-128-GCM', gcm: true },
  [OID.aes192GCM]: { name: 'aes-192-gcm', key: 24, label: 'AES-192-GCM', gcm: true },
  [OID.aes256GCM]: { name: 'aes-256-gcm', key: 32, label: 'AES-256-GCM', gcm: true },
  [OID.desEde3CBC]: { name: 'des-ede3-cbc', key: 24, label: 'Triple DES' },
};
