// The document's title, in its core properties (docProps/core.xml) — what a
// screen reader announces when the file opens, and File → Info's Title.

const CORE_REL = 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties';
const CORE_CT = 'application/vnd.openxmlformats-package.core-properties+xml';

const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const unesc = (s) => String(s).replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');

/** The core-properties part's name, or null when the package has none. */
export function corePart(pkg) {
  for (const rel of pkg.rels('')) {
    if (rel.Type === CORE_REL && rel.Target) {
      const name = String(rel.Target).replace(/^\//, '');
      if (pkg.has(name)) return name;
    }
  }
  return pkg.has('docProps/core.xml') ? 'docProps/core.xml' : null;
}

export function readTitle(pkg) {
  const part = corePart(pkg);
  if (!part) return '';
  const m = /<dc:title\b[^>]*>([\s\S]*?)<\/dc:title>/.exec(pkg.text(part));
  return m ? unesc(m[1]).trim() : '';
}

/** Set the title, making the part when the package has none. Returns the part's name. */
export function writeTitle(pkg, title) {
  const text = String(title ?? '').trim();
  let part = corePart(pkg);
  if (!part) {
    part = 'docProps/core.xml';
    pkg.addPart(part,
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
      + 'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
      + 'xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
      + '</cp:coreProperties>', CORE_CT);
    pkg.addRelationshipTo('', CORE_REL, part);
  }
  let xml = pkg.text(part);
  const element = text ? `<dc:title>${esc(text)}</dc:title>` : '';
  if (/<dc:title\b[^>]*\/>|<dc:title\b[^>]*>[\s\S]*?<\/dc:title>/.test(xml)) {
    xml = xml.replace(/<dc:title\b[^>]*\/>|<dc:title\b[^>]*>[\s\S]*?<\/dc:title>/, element);
  } else if (element) {
    if (!/xmlns:dc=/.test(xml)) xml = xml.replace(/<cp:coreProperties\b/, '<cp:coreProperties xmlns:dc="http://purl.org/dc/elements/1.1/"');
    xml = xml.replace(/(<cp:coreProperties\b[^>]*>)/, `$1${element}`);
  }
  pkg.write_(part, xml);
  return part;
}
