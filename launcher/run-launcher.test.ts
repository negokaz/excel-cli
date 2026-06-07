type MockProcess = {
  arch: NodeJS.Architecture;
  argv: string[];
  exit: jest.MockedFunction<(code?: number) => never>;
  platform: NodeJS.Platform;
  stderr: {
    write: jest.MockedFunction<(buffer: string | Uint8Array) => boolean>;
  };
};

function createMockProcess(overrides?: Partial<MockProcess>): MockProcess {
  return {
    arch: 'x64',
    argv: ['node', 'dist/launcher.js', 'query', 'workbook.xlsx', '/'],
    exit: jest.fn((code?: number) => {
      throw new Error(`process.exit:${code}`);
    }),
    platform: 'linux',
    stderr: {
      write: jest.fn((buffer: string | Uint8Array) => {
        void buffer;
        return true;
      }),
    },
    ...overrides,
  };
}

describe('executeLauncher', () => {
  afterEach(() => {
    jest.resetModules();
  });

  it('should pass CLI arguments to the selected binary when the runtime target is supported', () => {
    jest.isolateModules(() => {
      const runtimeProcess = createMockProcess();
      const spawnSync = jest.fn(() => ({ status: 0, signal: null }));

      jest.doMock('fs', () => ({
        existsSync: () => true,
      }));
      jest.doMock('child_process', () => ({
        spawnSync,
      }));

      const { executeLauncher } = require('./run-launcher');

      executeLauncher('/repo/dist', runtimeProcess);

      expect(spawnSync).toHaveBeenCalledWith(
        expect.stringContaining('excel-cli_linux_amd64_v1'),
        ['query', 'workbook.xlsx', '/'],
        { stdio: 'inherit' },
      );
      expect(runtimeProcess.exit).not.toHaveBeenCalled();
    });
  });

  it('should return the child exit code when the selected binary exits with a non-zero status', () => {
    jest.isolateModules(() => {
      const runtimeProcess = createMockProcess({
        exit: jest.fn() as unknown as MockProcess['exit'],
      });
      const spawnSync = jest.fn(() => ({ status: 5, signal: null }));

      jest.doMock('fs', () => ({
        existsSync: () => true,
      }));
      jest.doMock('child_process', () => ({
        spawnSync,
      }));

      const { executeLauncher } = require('./run-launcher');

      executeLauncher('/repo/dist', runtimeProcess);
      expect(runtimeProcess.exit).toHaveBeenCalledWith(5);
      expect(runtimeProcess.stderr.write).not.toHaveBeenCalled();
    });
  });

  it('should write a clear error and exit with code 1 when the runtime platform is unsupported', () => {
    const { executeLauncher } = require('./run-launcher');
    const runtimeProcess = createMockProcess({ platform: 'freebsd' as NodeJS.Platform });

    expect(() => executeLauncher('/repo/dist', runtimeProcess)).toThrow('process.exit:1');

    expect(runtimeProcess.stderr.write).toHaveBeenCalledWith(
      'Unsupported platform: freebsd_x64 (platform=freebsd, arch=x64)\n',
    );
  });
});
