// Data Files' ready-made databases: the tables a common job needs, related
// to each other, with a saved query or two, a form for each table and a
// report, as Access's own templates come. Each is made through the same
// calls a person's own design goes through, so every table, relationship
// and query is a numbered step like any other and can be changed after.
// The words a window shows for each (its name, what it is for) are the
// window's, to be translated there; here is only what goes in the file.

const ref = (table) => ({ type: 'integer', points: table });

const T = {
  contacts: {
    tables: [
      { name: 'Contacts', fields: [
        { name: 'First Name', type: 'text' }, { name: 'Last Name', type: 'text', required: true },
        { name: 'Company', type: 'text' }, { name: 'Job Title', type: 'text' },
        { name: 'E-mail', type: 'text' }, { name: 'Phone', type: 'text' }, { name: 'Mobile', type: 'text' },
        { name: 'Address', type: 'text' }, { name: 'City', type: 'text' }, { name: 'Postcode', type: 'text' }, { name: 'Country', type: 'text' },
        { name: 'Birthday', type: 'date' }, { name: 'Notes', type: 'longtext' },
      ] },
    ],
    queries: {
      'Contact list': { tables: ['Contacts'], columns: [
        { table: 'Contacts', field: 'Last Name', sort: 'asc' }, { table: 'Contacts', field: 'First Name', sort: 'asc' },
        { table: 'Contacts', field: 'Company' }, { table: 'Contacts', field: 'E-mail' }, { table: 'Contacts', field: 'Phone' },
      ] },
    },
    reports: [{ name: 'Phone list', table: 'Contacts', columns: ['Last Name', 'First Name', 'Phone', 'Mobile', 'E-mail'], sort: 'Last Name' }],
  },

  inventory: {
    tables: [
      { name: 'Suppliers', fields: [
        { name: 'Name', type: 'text', required: true }, { name: 'Contact', type: 'text' },
        { name: 'E-mail', type: 'text' }, { name: 'Phone', type: 'text' }, { name: 'Address', type: 'text' },
      ] },
      { name: 'Products', fields: [
        { name: 'Name', type: 'text', required: true }, { name: 'SKU', type: 'text', unique: true }, { name: 'Category', type: 'text' },
        { name: 'Supplier', ...ref('Suppliers') }, { name: 'Unit Price', type: 'currency' },
        { name: 'Units in Stock', type: 'integer', default: 0 }, { name: 'Reorder Level', type: 'integer', default: 5 },
        { name: 'Discontinued', type: 'yesno' },
      ] },
      { name: 'Stock Movements', fields: [
        { name: 'Product', ...ref('Products'), required: true }, { name: 'When', type: 'date' },
        { name: 'Quantity', type: 'integer' }, { name: 'Note', type: 'text' },
      ] },
    ],
    queries: {
      'Low stock': { tables: ['Products', 'Suppliers'], columns: [
        { table: 'Products', field: 'Name', as: 'Product', sort: 'asc' }, { table: 'Products', field: 'Units in Stock', criteria: ['< 10'] },
        { table: 'Products', field: 'Reorder Level' }, { table: 'Suppliers', field: 'Name', as: 'Supplier' }, { table: 'Suppliers', field: 'Phone' },
        { table: 'Products', field: 'Discontinued', criteria: ['No'], show: false },
      ] },
    },
    reports: [{ name: 'Products by supplier', table: 'Products', columns: ['Name', 'SKU', 'Category', 'Unit Price', 'Units in Stock'], group: 'Supplier', sort: 'Name', totals: [{ field: 'Units in Stock', fn: 'sum' }] }],
  },

  membership: {
    tables: [
      { name: 'Members', fields: [
        { name: 'First Name', type: 'text' }, { name: 'Last Name', type: 'text', required: true },
        { name: 'E-mail', type: 'text' }, { name: 'Phone', type: 'text' }, { name: 'Address', type: 'text' },
        { name: 'Membership Type', type: 'text', default: 'Standard' }, { name: 'Joined', type: 'date' },
        { name: 'Renewal Due', type: 'date' }, { name: 'Active', type: 'yesno', default: true },
      ] },
      { name: 'Payments', fields: [
        { name: 'Member', ...ref('Members'), required: true }, { name: 'Paid On', type: 'date' },
        { name: 'Amount', type: 'currency' }, { name: 'For', type: 'text' },
      ] },
    ],
    queries: {
      'Renewals due': { tables: ['Members'], columns: [
        { table: 'Members', field: 'Renewal Due', sort: 'asc', criteria: ['Is Not Null'] }, { table: 'Members', field: 'First Name' },
        { table: 'Members', field: 'Last Name' }, { table: 'Members', field: 'E-mail' }, { table: 'Members', field: 'Membership Type' },
        { table: 'Members', field: 'Active', criteria: ['Yes'], show: false },
      ] },
    },
    reports: [{ name: 'Payments by member', table: 'Payments', columns: ['Paid On', 'For', 'Amount'], group: 'Member', sort: 'Paid On', totals: [{ field: 'Amount', fn: 'sum' }] }],
  },

  crm: {
    tables: [
      { name: 'Companies', fields: [
        { name: 'Name', type: 'text', required: true }, { name: 'Industry', type: 'text' }, { name: 'Website', type: 'text' },
        { name: 'Phone', type: 'text' }, { name: 'City', type: 'text' }, { name: 'Country', type: 'text' },
      ] },
      { name: 'People', fields: [
        { name: 'Company', ...ref('Companies') }, { name: 'First Name', type: 'text' }, { name: 'Last Name', type: 'text', required: true },
        { name: 'Job Title', type: 'text' }, { name: 'E-mail', type: 'text' }, { name: 'Phone', type: 'text' },
      ] },
      { name: 'Opportunities', fields: [
        { name: 'Company', ...ref('Companies'), required: true }, { name: 'Title', type: 'text', required: true },
        { name: 'Stage', type: 'text', default: 'Lead' }, { name: 'Value', type: 'currency' },
        { name: 'Expected Close', type: 'date' }, { name: 'Owner', type: 'text' },
      ] },
      { name: 'Activities', fields: [
        { name: 'Person', ...ref('People') }, { name: 'When', type: 'datetime' },
        { name: 'Kind', type: 'text', default: 'Call' }, { name: 'Notes', type: 'longtext' },
      ] },
    ],
    queries: {
      'Open opportunities': { tables: ['Opportunities', 'Companies'], columns: [
        { table: 'Opportunities', field: 'Expected Close', sort: 'asc' }, { table: 'Companies', field: 'Name', as: 'Company' },
        { table: 'Opportunities', field: 'Title' }, { table: 'Opportunities', field: 'Stage', criteria: ['Not In ("Won", "Lost")'] },
        { table: 'Opportunities', field: 'Value' }, { table: 'Opportunities', field: 'Owner' },
      ] },
    },
    reports: [{ name: 'Pipeline by company', table: 'Opportunities', columns: ['Title', 'Stage', 'Expected Close', 'Value'], group: 'Company', sort: 'Expected Close', totals: [{ field: 'Value', fn: 'sum' }] }],
  },

  library: {
    tables: [
      { name: 'Books', fields: [
        { name: 'Title', type: 'text', required: true }, { name: 'Author', type: 'text' }, { name: 'ISBN', type: 'text' },
        { name: 'Publisher', type: 'text' }, { name: 'Year', type: 'integer' }, { name: 'Shelf', type: 'text' },
        { name: 'Copies', type: 'integer', default: 1 },
      ] },
      { name: 'Borrowers', fields: [
        { name: 'First Name', type: 'text' }, { name: 'Last Name', type: 'text', required: true },
        { name: 'E-mail', type: 'text' }, { name: 'Phone', type: 'text' },
      ] },
      { name: 'Loans', fields: [
        { name: 'Book', ...ref('Books'), required: true }, { name: 'Borrower', ...ref('Borrowers'), required: true },
        { name: 'Lent On', type: 'date' }, { name: 'Due Back', type: 'date' }, { name: 'Returned On', type: 'date' },
      ] },
    ],
    queries: {
      'On loan': { tables: ['Loans', 'Books', 'Borrowers'], columns: [
        { table: 'Loans', field: 'Due Back', sort: 'asc' }, { table: 'Books', field: 'Title' },
        { table: 'Borrowers', field: 'First Name' }, { table: 'Borrowers', field: 'Last Name' }, { table: 'Borrowers', field: 'Phone' },
        { table: 'Loans', field: 'Returned On', criteria: ['Is Null'], show: false },
      ] },
    },
    reports: [{ name: 'Loans by borrower', table: 'Loans', columns: ['Book', 'Lent On', 'Due Back', 'Returned On'], group: 'Borrower', sort: 'Lent On', totals: [{ field: 'Book', fn: 'count' }] }],
  },
};

/** The templates, in the order a window offers them. */
export const TEMPLATE_IDS = Object.keys(T);

/** A template's tables, as a list to show: [{ table, fields }]. */
export function templateTables(id) {
  const t = T[id];
  if (!t) throw new Error(`There is no template called ${id}.`);
  return t.tables.map((x) => ({ table: x.name, fields: x.fields.map((f) => f.name) }));
}

/**
 * Make template `id` in `base`, an empty database: its tables, the
 * relationships its pointing fields stand for, its queries, a form for each
 * table and its reports. Answers the names made.
 */
export function buildTemplate(base, id) {
  const t = T[id];
  if (!t) throw new Error(`There is no template called ${id}.`);
  if (base.tables().length) throw new Error('A template is made in an empty database.');
  for (const x of t.tables) {
    base.createTable({ name: x.name, fields: x.fields.map(({ points, ...f }) => f) });
  }
  for (const x of t.tables) {
    for (const f of x.fields.filter((g) => g.points)) {
      base.addRelationship({ from: { table: x.name, field: f.name }, to: { table: f.points, field: 'ID' } });
    }
  }
  for (const [name, design] of Object.entries(t.queries || {})) base.saveQuery(name, design);
  for (const x of t.tables) base.makeForm(x.name);
  for (const r of t.reports || []) {
    const made = base.makeReport({ table: r.table, name: r.name });
    const label = (field) => made.columns.find((c) => c.field === field)?.label || field;
    // A field that points at another table's record shows its words, so it stands on the left.
    const pointing = new Set(t.tables.find((x) => x.name === r.table).fields.filter((f) => f.points).map((f) => f.name));
    const types = Object.fromEntries(base.fields(r.table).map((f) => [f.name, pointing.has(f.name) ? 'text' : f.type]));
    base.saveReport({
      ...made,
      title: r.name,
      columns: r.columns.map((field) => ({ field, label: label(field), align: ['currency', 'number', 'integer'].includes(types[field]) ? 'right' : 'left' })),
      group: r.group || null,
      sort: r.sort || null,
      desc: false,
      totals: r.totals || [],
    });
  }
  return {
    tables: t.tables.map((x) => x.name),
    queries: Object.keys(t.queries || {}),
    reports: (t.reports || []).map((r) => r.name),
  };
}
