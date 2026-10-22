// The database engines Data Files links tables to (D-2026-10-22-13): the
// engine holds the records. PostgreSQL, MySQL or MariaDB; Microsoft SQL
// Server follows.

/** The engines offered, in the order the Link box lists them: { id, name, port, ready }. */
export const ENGINES = [
  { id: 'postgres', name: 'PostgreSQL', port: 5432, ready: true },
  { id: 'mysql', name: 'MySQL or MariaDB', port: 3306, ready: true },
  { id: 'mssql', name: 'Microsoft SQL Server', port: 1433, ready: false },
];

/** A connection to `engine` with `settings` ({ host, port, database, user, ssl }) and `password`. */
export async function connectEngine(engine, settings, password) {
  if (engine === 'postgres') return (await import('./postgres.js')).connect(settings, password);
  if (engine === 'mysql') return (await import('./mysql.js')).connect(settings, password);
  const known = ENGINES.find((e) => e.id === engine);
  throw new Error(known ? `${known.name} is not linked from Data Files yet: PostgreSQL and MySQL are.` : `${engine} is not an engine Data Files links to.`);
}
