import { config } from './config.js';
import { forceAuth } from './outlook/auth.js';
import { listUnreadMessages } from './outlook/inbox.js';
import { processMessage } from './pipeline.js';

const sleep = (ms: number) => new Promise(r => setTimeout(r, ms));
const QUEUE_BATCH_SIZE = 25;

function parseArgs(argv: string[]): { cmd: string; dryRun: boolean } {
  const positional = argv.filter(a => !a.startsWith('-'));
  const flags = new Set(argv.filter(a => a.startsWith('-')));
  return {
    cmd: positional[0] || 'start',
    dryRun: flags.has('--dry-run'),
  };
}

async function cmdAuth(): Promise<void> {
  await forceAuth();
}

// Process every unread message in the watched folder, oldest first, serially.
// `processMessage` handles its own errors (sending a friendly error reply +
// marking the source read) so the queue always advances.
async function cmdRunOnce(dryRun: boolean): Promise<void> {
  const messages = await listUnreadMessages(QUEUE_BATCH_SIZE);
  if (!messages.length) {
    console.log(`No unread messages in "${config.outlook.watchedFolder}".`);
    return;
  }
  console.log(`Processing ${messages.length} unread message(s) oldest-first.`);
  for (const m of messages) {
    try { await processMessage(m, { dryRun }); }
    catch (err) {
      // processMessage absorbs its own errors; only catastrophic failures
      // (e.g. Graph auth tokens evicted mid-loop) bubble up here.
      console.error(`Unhandled error on message ${m.id}:`, (err as Error).message);
    }
  }
}

async function cmdStart(dryRun: boolean): Promise<void> {
  console.log(`Watching Outlook folder "${config.outlook.watchedFolder}" for ${config.outlook.userEmail} (poll every ${config.pollIntervalMs / 1000}s).`);
  if (dryRun) console.log('DRY RUN mode — reports are generated but no replies are sent.');
  while (true) {
    try {
      const messages = await listUnreadMessages(QUEUE_BATCH_SIZE);
      for (const m of messages) {
        try { await processMessage(m, { dryRun }); }
        catch (err) {
          console.error(`Unhandled error on message ${m.id}:`, (err as Error).message);
        }
      }
    } catch (err) {
      console.error('Poll error:', (err as Error).message);
    }
    await sleep(config.pollIntervalMs);
  }
}

async function main(): Promise<void> {
  const { cmd, dryRun } = parseArgs(process.argv.slice(2));
  switch (cmd) {
    case 'auth': await cmdAuth(); break;
    case 'run-once': await cmdRunOnce(dryRun); break;
    case 'start': await cmdStart(dryRun); break;
    default:
      console.error(`Unknown command: ${cmd}`);
      console.error('Usage: decode-email-demo (auth | run-once | start) [--dry-run]');
      process.exit(2);
  }
}

main().catch(err => {
  console.error('Fatal:', err);
  process.exit(1);
});
