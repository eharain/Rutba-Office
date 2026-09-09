// The number formatter lives with the formula engine now
// (packages/formula/src/numfmt.js), so that TEXT() can format a date or a
// percentage with the same code the grid uses. This path stays, because the
// consumer line imports it here.
export * from '@rutba/formula/numfmt';
