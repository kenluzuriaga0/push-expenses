import 'dotenv/config';
import { getToken, fetchUnprocessedEmails, markProcessed } from './outlook';
import { connect, disconnect, findAccountId, isDuplicate, dollarsToUnits, insertTransaction } from './actual';

// ── Config ────────────────────────────────────────────────────────────────────

const REQUIRED = ['ACTUAL_SERVER_URL', 'ACTUAL_PASSWORD', 'ACTUAL_BUDGET_ID', 'ACTUAL_ACCOUNT_NAME', 'AZURE_CLIENT_ID'] as const;
const missing = REQUIRED.filter(k => !process.env[k]);
if (missing.length) {
  console.error('Faltan variables de entorno:', missing.join(', '));
  console.error('Copia .env.example a .env y completa los valores.');
  process.exit(1);
}

const config = {
  actual: {
    serverURL: process.env.ACTUAL_SERVER_URL!,
    password: process.env.ACTUAL_PASSWORD!,
    budgetId: process.env.ACTUAL_BUDGET_ID!,
    accountName: process.env.ACTUAL_ACCOUNT_NAME!,
    dataDir: process.env.ACTUAL_DATA_DIR ?? '.actual-data',
  },
  outlook: {
    clientId: process.env.AZURE_CLIENT_ID!,
    subjectRegex: process.env.EMAIL_SUBJECT_REGEX ?? 'Consumo Tarjeta de Crédito por (?:USD|\\$) ([\\d.]+)',
    processedLabel: process.env.PROCESSED_LABEL ?? 'ActualBudget',
    fromDate: process.env.EMAIL_FROM_DATE ?? null,
    limit: parseInt(process.env.EMAIL_LIMIT ?? '50', 10),
  },
};

// ── Main ──────────────────────────────────────────────────────────────────────

async function main(): Promise<void> {
  console.log('Autenticando en Outlook...');
  const token = await getToken(config.outlook.clientId);

  const { fromDate, limit } = config.outlook;
  console.log(fromDate ? `Leyendo emails desde ${fromDate}...` : `Leyendo los últimos ${limit} emails...`);
  const emails = await fetchUnprocessedEmails(token, config.outlook);

  if (emails.length === 0) {
    console.log('No hay emails nuevos para procesar.');
    return;
  }
  console.log(`Encontrados ${emails.length} email(s) para procesar.\n`);

  console.log('Conectando a Actual Budget...');
  await connect(config.actual);
  const accountId = await findAccountId(config.actual.accountName);

  let inserted = 0;
  let duplicates = 0;

  for (const email of emails) {
    const amountUnits = dollarsToUnits(email.amountUSD);

    if (await isDuplicate(accountId, email.date, amountUnits, email.notes)) {
      console.log(`  [dup]  ${email.date} | USD ${email.amountUSD.toFixed(2)} | ${email.notes} — ya existe`);
      duplicates++;
    } else {
      await insertTransaction(accountId, { date: email.date, amountUnits, notes: email.notes, payee: email.payee });
      const payeeTag = email.payee ? ` → ${email.payee}` : '';
      console.log(`  [ok]   ${email.date} | USD ${email.amountUSD.toFixed(2)} | ${email.notes}${payeeTag}`);
      inserted++;
    }

    await markProcessed(token, email.id, email.categories, config.outlook.processedLabel);
  }

  await disconnect();
  console.log(`\nResumen: ${inserted} insertado(s), ${duplicates} duplicado(s) omitido(s).`);
}

main().catch(err => {
  console.error('\nError:', err instanceof Error ? err.message : err);
  process.exit(1);
});
