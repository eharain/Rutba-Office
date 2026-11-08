// The Locker's strings go through the same t() as every other window. They are
// imported from here so the catalogue's list does not take them up until the
// Locker is opened to people; a build made without its switch must not carry
// a word of it, and that includes the translators' list.
export { t, tn } from '@rutba/office-ui';
