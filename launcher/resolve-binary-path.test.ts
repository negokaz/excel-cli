import * as path from 'path';
import { resolveBinaryPath } from './resolve-binary-path';

const FIXTURE_BASE_DIR = path.resolve('dist');
const existsAlways = (_filePath: string): boolean => true;
const existsNever = (_filePath: string): boolean => false;

describe('resolveBinaryPath', () => {
  it('should resolve the Windows launcher target when platform is win32 and arch is x64', () => {
    const result = resolveBinaryPath('win32', 'x64', FIXTURE_BASE_DIR, existsAlways);

    expect(result).toBe(path.resolve(FIXTURE_BASE_DIR, 'excel-cli_windows_amd64_v1', 'excel-cli.exe'));
  });

  it('should resolve the macOS launcher target when platform is darwin and arch is arm64', () => {
    const result = resolveBinaryPath('darwin', 'arm64', FIXTURE_BASE_DIR, existsAlways);

    expect(result).toBe(path.resolve(FIXTURE_BASE_DIR, 'excel-cli_darwin_arm64_v8.0', 'excel-cli'));
  });

  it('should fail fast when the runtime platform is unsupported', () => {
    expect(() => resolveBinaryPath('freebsd' as NodeJS.Platform, 'x64', FIXTURE_BASE_DIR, existsAlways))
      .toThrow('Unsupported platform: freebsd_x64');
  });

  it('should fail fast when the runtime architecture is unsupported', () => {
    expect(() => resolveBinaryPath('linux', 'mips', FIXTURE_BASE_DIR, existsAlways))
      .toThrow('Unsupported platform: linux_mips');
  });

  it('should fail fast when a 32-bit runtime is requested', () => {
    expect(() => resolveBinaryPath('win32', 'ia32', FIXTURE_BASE_DIR, existsAlways))
      .toThrow('Unsupported platform: win32_ia32');
  });

  it('should fail fast when the selected binary is missing', () => {
    expect(() => resolveBinaryPath('linux', 'x64', FIXTURE_BASE_DIR, existsNever))
      .toThrow(path.resolve(FIXTURE_BASE_DIR, 'excel-cli_linux_amd64_v1', 'excel-cli'));
  });
});
