import { mkdirSync } from 'fs';
import * as actual from '@actual-app/api';

// Suprime los logs internos de @actual-app/api durante una operación
function silenced<T>(fn: () => Promise<T>): Promise<T> {
  const noop = () => {};
  const orig = { log: console.log, warn: console.warn, error: console.error };
  console.log = noop; console.warn = noop; console.error = noop;
  return fn().finally(() => Object.assign(console, orig));
}

export interface ConnectOptions {
  dataDir: string;
  serverURL: string;
  password: string;
  budgetId: string;
}

export interface Transaction {
  date: string;
  amountUnits: number;
  notes: string;
  payee?: string;
}

// Gastos son negativos; 100 unidades = 1 dólar
export function dollarsToUnits(amount: number): number {
  return -Math.round(amount * 100);
}

export async function connect(opts: ConnectOptions): Promise<void> {
  mkdirSync(opts.dataDir, { recursive: true });
  await silenced(() => actual.init({ dataDir: opts.dataDir, serverURL: opts.serverURL, password: opts.password }));
  await silenced(() => actual.downloadBudget(opts.budgetId));
}

export async function disconnect(): Promise<void> {
  await actual.shutdown();
}

export async function findAccountId(name: string): Promise<string> {
  const accounts: Array<{ id: string; name: string }> = await actual.getAccounts();
  const account = accounts.find(a => a.name === name);
  if (!account) {
    const available = accounts.map(a => a.name).join(', ');
    throw new Error(`Cuenta "${name}" no encontrada. Disponibles: ${available}`);
  }
  return account.id;
}

export async function isDuplicate(
  accountId: string,
  date: string,
  amountUnits: number,
  notes: string,
): Promise<boolean> {
  const t = new Date(date).getTime();
  const pad = 3 * 86_400_000;
  const startDate = new Date(t - pad).toISOString().split('T')[0];
  const endDate = new Date(t + pad).toISOString().split('T')[0];
  const existing: Array<{ amount: number; notes: string }> = await actual.getTransactions(accountId, startDate, endDate);
  return existing.some(tx => tx.amount === amountUnits && tx.notes === notes);
}

export async function insertTransaction(accountId: string, tx: Transaction): Promise<void> {
  await silenced(() =>
    actual.addTransactions(accountId, [{
      date: tx.date,
      amount: tx.amountUnits,
      notes: tx.notes,
      ...(tx.payee ? { payee_name: tx.payee } : {}),
    }]),
  );
}
