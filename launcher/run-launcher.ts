import * as childProcess from 'child_process';
import * as fs from 'fs';
import { resolveBinaryPath } from './resolve-binary-path';

type ProcessLike = Pick<NodeJS.Process, 'arch' | 'argv' | 'exit' | 'platform'> & {
  stdin: {
    isTTY?: boolean;
  };
  stderr: {
    write(buffer: string | Uint8Array): boolean;
  };
};

type LauncherDependencies = {
  existsSync: typeof fs.existsSync;
  readStdin(): Buffer;
  spawnSync(
    command: string,
    args: string[],
    options: childProcess.SpawnSyncOptions,
  ): Pick<childProcess.SpawnSyncReturns<Buffer>, 'error' | 'signal' | 'status'>;
};

function formatErrorMessage(error: unknown): string {
  if (error instanceof Error) {
    return error.message;
  }

  return String(error);
}

function runLauncher(
  baseDir: string,
  runtimeProcess: ProcessLike,
  dependencies: LauncherDependencies,
): void {
  const binaryPath = resolveBinaryPath(
    runtimeProcess.platform,
    runtimeProcess.arch,
    baseDir,
    dependencies.existsSync,
  );

  let result: Pick<childProcess.SpawnSyncReturns<Buffer>, 'error' | 'signal' | 'status'>;
  if (runtimeProcess.stdin.isTTY) {
    result = dependencies.spawnSync(binaryPath, runtimeProcess.argv.slice(2), {
      stdio: 'inherit'
    });
  } else {
    const stdinData = dependencies.readStdin();
    result = dependencies.spawnSync(binaryPath, runtimeProcess.argv.slice(2), {
      stdio: ['pipe', 'inherit', 'inherit'],
      input: stdinData,
    });
  }

  if (result.error !== undefined) {
    throw result.error;
  }

  if (result.signal !== null) {
    throw new Error(`Binary terminated by signal: ${result.signal}`);
  }

  if (typeof result.status !== 'number') {
    throw new Error(`Binary exited without a status code: ${binaryPath}`);
  }

  if (result.status !== 0) {
    runtimeProcess.exit(result.status);
  }
}

export function executeLauncher(baseDir: string, runtimeProcess: ProcessLike): void {
  try {
    runLauncher(baseDir, runtimeProcess, {
      existsSync: fs.existsSync,
      readStdin: () => fs.readFileSync(0),
      spawnSync: childProcess.spawnSync,
    });
  } catch (error) {
    runtimeProcess.stderr.write(`${formatErrorMessage(error)}\n`);
    runtimeProcess.exit(1);
  }
}
