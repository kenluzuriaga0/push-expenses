import 'dotenv/config';
import * as actual from '@actual-app/api';
import { connect, disconnect } from './actual';

const { ACTUAL_SERVER_URL, ACTUAL_PASSWORD, ACTUAL_BUDGET_ID, ACTUAL_ACCOUNT_NAME, ACTUAL_DATA_DIR } = process.env;

if (!ACTUAL_SERVER_URL || !ACTUAL_PASSWORD || !ACTUAL_BUDGET_ID || !ACTUAL_ACCOUNT_NAME) {
  console.error('Faltan variables de entorno de Actual Budget en .env');
  process.exit(1);
}

async function main(): Promise<void> {
  await connect({
    dataDir: ACTUAL_DATA_DIR ?? '.actual-data',
    serverURL: ACTUAL_SERVER_URL!,
    password: ACTUAL_PASSWORD!,
    budgetId: ACTUAL_BUDGET_ID!,
  });

  const accounts: Array<{ id: string; name: string }> = await actual.getAccounts();
  const account = accounts.find(a => a.name === ACTUAL_ACCOUNT_NAME);
  if (!account) {
    const available = accounts.map(a => a.name).join(', ');
    throw new Error(`Cuenta "${ACTUAL_ACCOUNT_NAME}" no encontrada. Disponibles: ${available}`);
  }

  const today = new Date();
  const startDate = new Date(today.getTime() - 30 * 86_400_000).toISOString().split('T')[0];
  const endDate = today.toISOString().split('T')[0];

  const txs: Array<{ date: string; amount: number; notes: string }> =
    await actual.getTransactions(account.id, startDate, endDate);

  const sorted = txs.sort((a, b) => b.date.localeCompare(a.date)).slice(0, 10);

  console.log(`\nÚltimas transacciones — ${ACTUAL_ACCOUNT_NAME} (últimos 30 días)\n`);
  console.log('─'.repeat(60));

  if (sorted.length === 0) {
    console.log('  Sin transacciones en este período.');
  } else {
    sorted.forEach((t, i) => {
      const amount = (Math.abs(t.amount) / 100).toFixed(2);
      const sign = t.amount < 0 ? '-' : '+';
      console.log(`[${i + 1}] ${t.date} | ${sign}USD ${amount}`);
      if (t.notes) console.log(`     ${t.notes}`);
      console.log('─'.repeat(60));
    });
  }

  await disconnect();
}

main().catch(err => {
  console.error('\nError:', err instanceof Error ? err.message : err);
  process.exit(1);
});
